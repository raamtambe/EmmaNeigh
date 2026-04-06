#!/usr/bin/env python3
"""
EmmaNeigh - Natural Language Email Search
Uses Claude API to answer natural language questions about emails.
"""

import os
import sys
import json
import re
from datetime import datetime

# Try to import anthropic, provide helpful error if not available
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
QUERY_STOPWORDS = {
    "a", "an", "and", "are", "attachment", "attachments", "by", "can", "contains",
    "did", "do", "does", "email", "emails", "exact", "find", "for", "from", "has",
    "have", "i", "if", "in", "is", "it", "latest", "me", "message", "messages",
    "name", "named", "of", "on", "or", "our", "pull", "received", "search", "sent",
    "show", "that", "the", "their", "there", "these", "this", "title", "titles",
    "to", "up", "we", "what", "which", "with"
}


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


def truncate_email_body(body, max_chars=300):
    """Truncate email body to save tokens."""
    if not body:
        return ""
    body = body.strip()
    if len(body) <= max_chars:
        return body
    return body[:max_chars] + "..."


def clean_email_body_for_search(body):
    lines = []
    header_lines_seen = 0
    for raw_line in str(body or "").splitlines():
        line = raw_line.strip()
        if not line:
            continue
        lower_line = line.lower()
        if line.startswith(">"):
            continue
        if lower_line.startswith("-----original message-----"):
            break
        if re.match(r"^on .+ wrote:$", lower_line):
            break
        if re.match(r"^(from|sent|to|cc|subject):", lower_line):
            header_lines_seen += 1
            if header_lines_seen >= 2:
                break
            continue
        if "privileged and confidential" in lower_line or "confidentiality notice" in lower_line:
            break
        lines.append(line)
        if len(lines) >= 80:
            break
    return " ".join(lines)


def split_attachment_names(raw_value):
    if isinstance(raw_value, list):
        return [str(value).strip() for value in raw_value if str(value).strip()]

    text = str(raw_value or "").strip()
    if not text:
        return []
    if ";" in text:
        parts = text.split(";")
    elif "\n" in text:
        parts = text.splitlines()
    else:
        parts = text.split(",")
    return [part.strip() for part in parts if part.strip()]


def normalize_search_text(value):
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", str(value or "").lower())).strip()


def tokenize_query(query):
    tokens = []
    for token in normalize_search_text(query).split():
        if len(token) < 3 or token in QUERY_STOPWORDS:
            continue
        tokens.append(token)
    return tokens


def parse_email_timestamp(email):
    candidates = [
        email.get("date_received", ""),
        email.get("date_sent", ""),
        email.get("date", ""),
    ]
    for candidate in candidates:
        text = str(candidate or "").strip()
        if not text:
            continue
        try:
            return int(datetime.fromisoformat(text.replace("Z", "+00:00")).timestamp())
        except Exception:
            continue
    return 0


def score_attachment_title_match(attachment_names, query):
    raw_query = str(query or "").strip().lower()
    normalized_query = normalize_search_text(query)
    query_tokens = tokenize_query(query)
    score = 0
    matched_titles = []

    for attachment_name in attachment_names:
        title = str(attachment_name or "").strip()
        if not title:
            continue
        title_lower = title.lower()
        normalized_title = normalize_search_text(title)
        local_score = 0

        if raw_query and raw_query in title_lower:
            local_score = max(local_score, 90)
        if normalized_query and normalized_query in normalized_title:
            local_score = max(local_score, 110)

        if query_tokens:
            token_hits = [token for token in query_tokens if token in normalized_title]
            if token_hits:
                local_score = max(local_score, len(token_hits) * 14)
                if len(query_tokens) >= 2 and len(token_hits) == len(query_tokens):
                    local_score = max(local_score, 75)

        if local_score > 0:
            score = max(score, local_score)
            matched_titles.append(title)

    return {
        "score": score,
        "matched_attachment_titles": matched_titles[:5],
    }


def score_email_for_query(email, query):
    attachment_names = split_attachment_names(email.get("attachments", ""))
    attachment_match = score_attachment_title_match(attachment_names, query)
    normalized_query = normalize_search_text(query)
    query_tokens = tokenize_query(query)
    subject_text = normalize_search_text(email.get("subject", ""))
    sender_text = normalize_search_text(email.get("from", ""))
    body_text = normalize_search_text(clean_email_body_for_search(email.get("body", ""))[:4000])
    recipient_text = normalize_search_text(f"{email.get('to', '')} {email.get('cc', '')}")

    score = attachment_match["score"]
    if normalized_query:
        if normalized_query in subject_text:
            score += 24
        if normalized_query in body_text:
            score += 10
        if normalized_query in sender_text:
            score += 8
        if normalized_query in recipient_text:
            score += 5

    if query_tokens:
        searchable = " ".join([subject_text, body_text, sender_text, recipient_text])
        token_hits = sum(1 for token in query_tokens if token in searchable)
        score += token_hits * 3

    if attachment_match["score"] > 0:
        score += 5

    return {
        "score": score,
        "matched_attachment_titles": attachment_match["matched_attachment_titles"],
    }


def select_emails_for_query(emails, query, max_emails=120):
    normalized_max = max(10, min(int(max_emails or 120), 200))
    scored = []
    for index, email in enumerate(emails):
        match = score_email_for_query(email, query)
        scored.append({
            "index": index,
            "email": email,
            "score": match["score"],
            "timestamp": parse_email_timestamp(email),
            "matched_attachment_titles": match["matched_attachment_titles"],
        })

    selected = []
    seen = set()

    positive = sorted(
        (item for item in scored if item["score"] > 0),
        key=lambda item: (-item["score"], -item["timestamp"], item["index"]),
    )
    for item in positive:
        selected.append(item)
        seen.add(item["index"])
        if len(selected) >= normalized_max:
            break

    minimum_context = min(normalized_max, 40)
    if len(selected) < minimum_context:
        fallback = sorted(
            (item for item in scored if item["index"] not in seen),
            key=lambda item: (-item["timestamp"], item["index"]),
        )
        for item in fallback:
            selected.append(item)
            seen.add(item["index"])
            if len(selected) >= minimum_context:
                break

    return selected[:normalized_max]


def prepare_email_context(emails, query, max_emails=120):
    """
    Prepare emails for sending to Claude API.
    Pre-ranks likely matches, especially attachment-title matches, and preserves original indices.
    """
    context_emails = []

    for selected in select_emails_for_query(emails, query, max_emails=max_emails):
        email = selected["email"]
        context_emails.append({
            "index": selected["index"],
            "from": email.get("from", "Unknown"),
            "to": email.get("to", ""),
            "subject": email.get("subject", "(No Subject)"),
            "body_preview": truncate_email_body(clean_email_body_for_search(email.get("body", ""))),
            "date": email.get("date_received") or email.get("date_sent") or "",
            "attachments": email.get("attachments", ""),
            "attachment_titles": split_attachment_names(email.get("attachments", "")),
            "matched_attachment_titles": selected.get("matched_attachment_titles", []),
            "attachment_match_score": selected.get("score", 0),
            "has_attachments": email.get("has_attachments", False)
        })

    return context_emails


def perform_nl_search(emails, query, api_key):
    """
    Use Claude API to answer natural language questions about emails.

    Args:
        emails: List of email dictionaries
        query: Natural language question
        api_key: Claude API key

    Returns:
        dict with answer, relevant_emails, confidence
    """
    if not HAS_ANTHROPIC:
        return {
            "success": False,
            "error": "Anthropic SDK not installed. Please install with: pip install anthropic"
        }

    if not api_key:
        return {
            "success": False,
            "error": "No API key provided"
        }

    try:
        client = anthropic.Anthropic(api_key=api_key)

        # Prepare email context
        email_context = prepare_email_context(emails, query)

        emit("progress", percent=30, message="Analyzing emails with AI...")

        # Build the prompt
        prompt = f"""You are an email assistant analyzing a database of emails from a legal transaction.

User Question: {query}

Email Database ({len(email_context)} emails, already pre-ranked for likely relevance including attachment title matches):
{json.dumps(email_context, indent=2, default=str)}

Please analyze these emails and answer the user's question. Be specific and cite relevant emails by their index number.

Respond with a JSON object containing:
{{
    "answer": "Your detailed answer to the question",
    "relevant_email_indices": [0, 5, 12],  // indices of most relevant emails
    "confidence": 0.85,  // your confidence level 0.0-1.0
    "summary": "One-sentence summary of your finding"
}}

Important:
- If you can't find relevant information, say so clearly
- Be specific about which emails support your answer
- For version/latest questions, pay attention to dates
- For "did we receive X from Y" questions, check the from field carefully
- Exact attachment titles matter. If the query names a document, prioritize emails whose attachment_titles or matched_attachment_titles fit that document name.
- Use the provided index values exactly as written. They refer to the original loaded emails.

Respond ONLY with the JSON object, no other text."""

        emit("progress", percent=50, message="Waiting for AI response...")

        message = None
        used_model = None
        model_errors = []
        for model_name in get_model_candidates():
            try:
                message = client.messages.create(
                    model=model_name,
                    max_tokens=1024,
                    messages=[
                        {"role": "user", "content": prompt}
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

        emit("progress", percent=80, message="Processing response...")

        # Parse the response
        response_text = message.content[0].text.strip()

        # Try to extract JSON from the response
        try:
            # Handle case where response might have markdown code blocks
            if response_text.startswith("```"):
                lines = response_text.split("\n")
                json_lines = []
                in_json = False
                for line in lines:
                    if line.startswith("```json"):
                        in_json = True
                        continue
                    if line.startswith("```"):
                        in_json = False
                        continue
                    if in_json:
                        json_lines.append(line)
                response_text = "\n".join(json_lines)

            result = json.loads(response_text)

            emit("progress", percent=100, message="Complete!")

            return {
                "success": True,
                "answer": result.get("answer", "No answer provided"),
                "relevant_email_indices": result.get("relevant_email_indices", []),
                "confidence": result.get("confidence", 0.5),
                "summary": result.get("summary", ""),
                "model_used": used_model,
                "query": query
            }

        except json.JSONDecodeError:
            # If JSON parsing fails, return the raw text as the answer
            return {
                "success": True,
                "answer": response_text,
                "relevant_email_indices": [],
                "confidence": 0.5,
                "summary": "",
                "model_used": used_model,
                "query": query
            }

    except anthropic.AuthenticationError:
        return {
            "success": False,
            "error": "Invalid API key"
        }
    except anthropic.RateLimitError:
        return {
            "success": False,
            "error": "Rate limit exceeded. Please try again later."
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
        emit("error", message="Usage: email_nl_search.py <config_json_path>")
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
    query = config.get('query', '')
    api_key = config.get('api_key', '')

    if not query:
        emit("error", message="No query provided")
        sys.exit(1)

    if not emails:
        emit("error", message="No emails provided")
        sys.exit(1)

    emit("progress", percent=10, message="Starting natural language search...")

    result = perform_nl_search(emails, query, api_key)

    emit("result", **result)


if __name__ == "__main__":
    main()
