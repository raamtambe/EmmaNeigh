#!/usr/bin/env python3
"""
EmmaNeigh - Task Detector
v5.3.0: Analyzes emails to identify actionable tasks for a first-year associate.
Uses Claude API to classify emails into task categories (collate, redline, sig packets, etc.)
and extract metadata (document names, signers, deadlines, priority).
"""

import os
import sys
import json
import re

try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False

DEFAULT_CLAUDE_MODEL = "claude-sonnet-4-20250514"
MODEL_ALIASES = {
    "claude-sonnet-4-6": "claude-sonnet-4-20250514",
}
FALLBACK_MODELS = [
    "claude-sonnet-4-20250514",
    "claude-3-5-sonnet-20241022",
    "claude-3-5-haiku-20241022",
]


def normalize_model_name(model):
    """Normalize model aliases and empty values."""
    value = (model or "").strip()
    if not value:
        return DEFAULT_CLAUDE_MODEL
    return MODEL_ALIASES.get(value, value)


def get_model_candidates():
    """Return model candidates in priority order, de-duplicated."""
    preferred = normalize_model_name(os.environ.get("CLAUDE_MODEL", DEFAULT_CLAUDE_MODEL))
    candidates = [preferred] + FALLBACK_MODELS
    seen = set()
    ordered = []
    for model in candidates:
        if model not in seen:
            seen.add(model)
            ordered.append(model)
    return ordered


def is_model_error(exc):
    """Heuristic for API errors caused by invalid/unavailable model IDs."""
    msg = str(exc).lower()
    return "model" in msg and ("not found" in msg or "invalid" in msg or "available" in msg or "access" in msg)


def emit(msg_type, **kwargs):
    """Output JSON message to stdout for the Electron app."""
    print(json.dumps({"type": msg_type, **kwargs}), flush=True)


def truncate_body(body, max_chars=500):
    """Truncate email body to save tokens while keeping enough for task detection."""
    if not body:
        return ""
    body = body.strip()
    if len(body) <= max_chars:
        return body
    return body[:max_chars] + "..."


def prepare_emails_for_detection(emails, max_emails=50):
    """
    Prepare emails for task detection. We include more body text than
    the NL search because task detection needs to understand context.
    """
    prepared = []
    for i, email in enumerate(emails[:max_emails]):
        prepared.append({
            "index": i,
            "from": email.get("from", "Unknown"),
            "to": email.get("to", ""),
            "cc": email.get("cc", ""),
            "subject": email.get("subject", "(No Subject)"),
            "body": truncate_body(email.get("body", "")),
            "date": email.get("date_received") or email.get("date_sent") or "",
            "attachments": email.get("attachments", ""),
            "has_attachments": email.get("has_attachments", False)
        })
    return prepared


TASK_DETECTION_PROMPT = """You are a legal transaction assistant analyzing emails sent to a first-year associate at a law firm. Your job is to identify actionable tasks that the associate needs to perform.

TASK CATEGORIES:
- COLLATE: Merge comments or track changes from multiple document versions into a single document. Keywords: "collate", "merge comments", "consolidate changes", "combine markups", "track changes from all parties"
- REDLINE: Compare two document versions to identify and mark up changes. Keywords: "redline", "compare", "blackline", "show changes", "amended version", "mark up differences"
- SIG_PACKETS: Extract signature pages from closing documents and organize them per signer. Keywords: "signature pages", "sig packets", "closing binder", "execution pages", "signing set"
- EXECUTION_VERSION: Merge signed/executed pages back into the original unsigned agreements. Keywords: "execution version", "conformed copy", "insert signed pages", "final executed"
- REVIEW: Review a document but no automated action is needed. Keywords: "please review", "take a look", "your thoughts on", "comments on"
- NONE: No actionable task detected (informational emails, FYIs, scheduling, etc.)

IMPORTANT RULES:
- Only detect tasks where there is a clear action being requested
- Do NOT classify FYI emails, scheduling emails, or status updates as tasks
- An email saying "attached is the agreement" without asking for action = NONE
- An email saying "please compare the attached against the original" = REDLINE
- Look for action verbs: "please collate", "can you compare", "extract signature pages", etc.

For each task detected, extract:
- task_type: one of COLLATE, REDLINE, SIG_PACKETS, EXECUTION_VERSION, REVIEW, NONE
- source_email_index: the email index this task came from
- documents: list of document filenames or descriptions referenced
- signers: list of person/entity names mentioned as signers (for SIG_PACKETS only)
- deadline: any deadline mentioned (date string or null)
- priority: HIGH (urgent/ASAP/today), MEDIUM (this week/soon), LOW (when you get a chance/no rush)
- summary: one-sentence description of what needs to be done

Respond with a JSON object:
{
    "tasks": [
        {
            "task_type": "COLLATE",
            "source_email_index": 3,
            "documents": ["Credit Agreement v3.docx", "Credit Agreement - Client Comments.docx"],
            "signers": [],
            "deadline": "2026-02-25",
            "priority": "HIGH",
            "summary": "Collate client comments on Credit Agreement into the base version"
        }
    ],
    "total_emails_analyzed": 15,
    "emails_with_tasks": 3
}

Respond ONLY with the JSON object, no other text."""


def detect_tasks(emails, api_key):
    """
    Use Claude API to detect actionable tasks from emails.

    Args:
        emails: List of email dictionaries
        api_key: Claude API key

    Returns:
        dict with tasks list and metadata
    """
    if not HAS_ANTHROPIC:
        return {
            "success": False,
            "error": "Anthropic SDK not installed. Please install with: pip install anthropic"
        }

    if not api_key:
        return {
            "success": False,
            "error": "No API key provided. Please add your Claude API key in Settings."
        }

    if not emails:
        return {
            "success": False,
            "error": "No emails to analyze"
        }

    try:
        client = anthropic.Anthropic(api_key=api_key)

        # Prepare email context
        email_context = prepare_emails_for_detection(emails)

        emit("progress", percent=20, message=f"Analyzing {len(email_context)} emails for tasks...")

        user_message = f"""{TASK_DETECTION_PROMPT}

EMAILS TO ANALYZE ({len(email_context)} emails):
{json.dumps(email_context, indent=2, default=str)}"""

        emit("progress", percent=40, message="Waiting for AI analysis...")

        message = None
        used_model = None
        model_errors = []
        for model_name in get_model_candidates():
            try:
                message = client.messages.create(
                    model=model_name,
                    max_tokens=4096,
                    messages=[
                        {"role": "user", "content": user_message}
                    ]
                )
                used_model = model_name
                break
            except anthropic.APIError as e:
                if is_model_error(e):
                    model_errors.append(f"{model_name}: {str(e)}")
                    continue
                raise

        if message is None:
            return {
                "success": False,
                "error": "No supported Claude model is available for this API key. "
                         f"Tried: {', '.join(get_model_candidates())}. "
                         f"Last model error: {model_errors[-1] if model_errors else 'unknown'}"
            }

        emit("progress", percent=75, message="Processing detected tasks...")

        response_text = message.content[0].text.strip()

        # Parse JSON from response (handle markdown code blocks)
        try:
            json_text = response_text
            if json_text.startswith("```"):
                lines = json_text.split("\n")
                json_lines = []
                in_json = False
                for line in lines:
                    if line.startswith("```json") or line.startswith("```"):
                        if not in_json:
                            in_json = True
                            continue
                        else:
                            in_json = False
                            continue
                    if in_json:
                        json_lines.append(line)
                json_text = "\n".join(json_lines)

            result = json.loads(json_text)
            tasks = result.get("tasks", [])

            # Filter out NONE and REVIEW tasks for the actionable list
            actionable_tasks = [t for t in tasks if t.get("task_type") not in ("NONE",)]

            emit("progress", percent=100, message=f"Found {len(actionable_tasks)} actionable tasks!")

            return {
                "success": True,
                "tasks": tasks,
                "actionable_tasks": actionable_tasks,
                "total_emails_analyzed": result.get("total_emails_analyzed", len(email_context)),
                "emails_with_tasks": result.get("emails_with_tasks", len(actionable_tasks)),
                "model_used": used_model
            }

        except json.JSONDecodeError:
            # If JSON parsing fails, try to extract tasks manually
            return {
                "success": True,
                "tasks": [],
                "actionable_tasks": [],
                "total_emails_analyzed": len(email_context),
                "emails_with_tasks": 0,
                "raw_response": response_text,
                "warning": "AI response could not be parsed as structured data"
            }

    except anthropic.AuthenticationError:
        return {
            "success": False,
            "error": "Invalid API key. Please check your API key in Settings."
        }
    except anthropic.RateLimitError:
        return {
            "success": False,
            "error": "Rate limit exceeded. Please try again in a few seconds."
        }
    except anthropic.APIError as e:
        return {
            "success": False,
            "error": f"API error: {str(e)}"
        }
    except Exception as e:
        return {
            "success": False,
            "error": f"Unexpected error: {str(e)}"
        }


def main():
    """CLI entry point."""
    if len(sys.argv) < 2:
        emit("error", message="Usage: task_detector.py <config_json_path>")
        sys.exit(1)

    config_path = sys.argv[1]

    if not os.path.isfile(config_path):
        emit("error", message=f"Config file not found: {config_path}")
        sys.exit(1)

    try:
        with open(config_path, 'r') as f:
            config = json.load(f)
    except json.JSONDecodeError as e:
        emit("error", message=f"Invalid JSON config: {str(e)}")
        sys.exit(1)

    emails = config.get('emails', [])
    api_key = config.get('api_key', '')

    if not emails:
        emit("error", message="No emails provided for task detection")
        sys.exit(1)

    emit("progress", percent=5, message="Starting task detection...")

    result = detect_tasks(emails, api_key)

    emit("result", **result)


if __name__ == "__main__":
    main()
