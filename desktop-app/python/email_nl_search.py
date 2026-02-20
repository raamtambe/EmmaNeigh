#!/usr/bin/env python3
"""
EmmaNeigh - Natural Language Email Search
Uses Claude API to answer natural language questions about emails.
"""

import os
import sys
import json

# Try to import anthropic, provide helpful error if not available
try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False


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


def prepare_email_context(emails, max_emails=100):
    """
    Prepare emails for sending to Claude API.
    Limits the number of emails and truncates bodies to manage token usage.
    """
    context_emails = []

    for i, email in enumerate(emails[:max_emails]):
        context_emails.append({
            "index": i,
            "from": email.get("from", "Unknown"),
            "to": email.get("to", ""),
            "subject": email.get("subject", "(No Subject)"),
            "body_preview": truncate_email_body(email.get("body", "")),
            "date": email.get("date_received") or email.get("date_sent") or "",
            "attachments": email.get("attachments", ""),
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
        email_context = prepare_email_context(emails)

        emit("progress", percent=30, message="Analyzing emails with AI...")

        # Build the prompt
        prompt = f"""You are an email assistant analyzing a database of emails from a legal transaction.

User Question: {query}

Email Database ({len(email_context)} emails):
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

Respond ONLY with the JSON object, no other text."""

        emit("progress", percent=50, message="Waiting for AI response...")

        message = client.messages.create(
            model=os.environ.get("CLAUDE_MODEL", "claude-sonnet-4-20250514"),
            max_tokens=1024,
            messages=[
                {"role": "user", "content": prompt}
            ]
        )

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
