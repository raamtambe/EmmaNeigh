#!/usr/bin/env python3
"""
Checklist Updater - Update transaction checklist based on email activity

This module parses a transaction checklist (Word document) and an email CSV export,
then updates the checklist status column based on detected email activity.

Usage:
    python checklist_updater.py <checklist_path> <email_csv_path> <output_folder> <api_key> [provider] [model] [provider_base_url]
"""

import sys
import os
import json
import re
import socket
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
from datetime import datetime
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

DEFAULT_CLAUDE_MODEL = "claude-sonnet-4-20250514"
DEFAULT_OPENAI_MODEL = "gpt-4.1-mini"
DEFAULT_HARVEY_BASE_URL = "https://api.harvey.ai"
DEFAULT_OLLAMA_BASE_URL = "http://127.0.0.1:11434"
DEFAULT_LMSTUDIO_BASE_URL = "http://127.0.0.1:1234"
DEFAULT_OLLAMA_MODEL = "llama3.1:8b"
DEFAULT_LMSTUDIO_MODEL = "local-model"
MAX_HTTP_ATTEMPTS = 3
RETRYABLE_HTTP_STATUS = {408, 425, 429, 500, 502, 503, 504}
RETRYABLE_ERROR_MARKERS = (
    "timeout",
    "temporarily unavailable",
    "connection reset",
    "connection aborted",
    "network is unreachable",
    "name or service not known",
    "temporary failure in name resolution",
)

LLM_STATUS_VALUES = {
    "Pending Draft",
    "Draft Circulated",
    "With Opposing Counsel",
    "Agreed Form",
    "Execution Version",
    "Executed",
}

STATUS_ALIASES = {
    "pending": "Pending Draft",
    "pending draft": "Pending Draft",
    "to be drafted": "Pending Draft",
    "draft circulated": "Draft Circulated",
    "with opposing counsel": "With Opposing Counsel",
    "under review": "With Opposing Counsel",
    "agreed form": "Agreed Form",
    "execution version": "Execution Version",
    "executed": "Executed",
    "fully executed": "Executed",
}

# Status detection patterns - order matters (more specific first)
STATUS_PATTERNS = [
    # Executed/Signed (highest priority)
    {
        'patterns': [
            r'\bfully\s+executed\b',
            r'\bexecuted\s+(?:copy|version|document)\b',
            r'\bsigned\s+(?:by\s+all|copy|version)\b',
            r'\b(?:all\s+)?signatures?\s+(?:received|obtained|collected)\b',
        ],
        'status': 'Executed',
        'priority': 5
    },
    # Execution Version Circulated
    {
        'patterns': [
            r'\bexecution\s+(?:version|copy)\b',
            r'\bfor\s+(?:signature|execution)\b',
            r'\b(?:ready|circulated?)\s+for\s+signature\b',
            r'\bsignature\s+pages?\s+(?:attached|circulated)\b',
        ],
        'status': 'Execution Version',
        'priority': 4
    },
    # Agreed Form
    {
        'patterns': [
            r'\bagreed\s+form\b',
            r'\bfinal\s+(?:form|version)\b',
            r'\bfinalized\b',
            r'\bno\s+(?:further\s+)?comments\b',
        ],
        'status': 'Agreed Form',
        'priority': 3
    },
    # Sent to Opposing Counsel / Under Review
    {
        'patterns': [
            r'\bsent\s+to\s+(?:opposing\s+)?counsel\b',
            r'\bsent\s+to\s+(?:counterparty|other\s+side|buyer|seller|lender|borrower)\b',
            r'\b(?:circulated|distributed)\s+(?:to|for)\s+(?:review|comment)\b',
            r'\bfor\s+(?:your\s+)?review\b',
            r'\bplease\s+(?:review|comment)\b',
            r'\bawaiting\s+(?:comments?|review|feedback)\b',
            r'\bunder\s+review\b',
        ],
        'status': 'With Opposing Counsel',
        'priority': 2
    },
    # Draft Circulated (internal)
    {
        'patterns': [
            r'\b(?:initial\s+)?draft\s+(?:attached|circulated)\b',
            r'\battached\s+(?:is\s+)?(?:a\s+)?draft\b',
            r'\bfirst\s+draft\b',
        ],
        'status': 'Draft Circulated',
        'priority': 1
    },
]

# Column name patterns for detecting checklist columns
DOCUMENT_COLUMN_PATTERNS = [
    r'^document',
    r'^doc\s*name',
    r'^agreement',
    r'^instrument',
    r'^deliverable',
    r'^item',
    r'^description',
    r'^closing\s*document',
]

STATUS_COLUMN_PATTERNS = [
    r'^status',
    r'^state',
    r'^progress',
    r'^current\s*status',
]


def parse_email_csv(csv_path):
    """
    Parse Outlook email CSV export.

    Returns list of email dicts with: subject, body, from, to, date
    """
    try:
        # Try different encodings
        for encoding in ['utf-8', 'latin-1', 'cp1252']:
            try:
                df = pd.read_csv(csv_path, encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError("Could not decode CSV file")

        # Normalize column names
        df.columns = df.columns.str.lower().str.strip()

        # Map common column names
        column_mapping = {
            'subject': ['subject', 'email subject', 'title'],
            'body': ['body', 'content', 'message', 'email body', 'notes'],
            'from': ['from', 'sender', 'from email', 'from address'],
            'to': ['to', 'recipient', 'to email', 'to address', 'recipients'],
            'date': ['date', 'sent', 'received', 'date sent', 'date received', 'sent date'],
        }

        emails = []
        for _, row in df.iterrows():
            email = {}
            for field, possible_cols in column_mapping.items():
                for col in possible_cols:
                    if col in df.columns and pd.notna(row.get(col)):
                        email[field] = str(row[col])
                        break
                else:
                    email[field] = ''

            # Combine subject and body for searching
            email['searchable'] = f"{email.get('subject', '')} {email.get('body', '')}".lower()
            emails.append(email)

        return emails

    except Exception as e:
        print(f"Error parsing email CSV: {e}", file=sys.stderr)
        return []


def find_column_index(headers, patterns):
    """Find column index matching any of the patterns."""
    for i, header in enumerate(headers):
        header_lower = header.lower().strip()
        for pattern in patterns:
            if re.match(pattern, header_lower):
                return i
    return -1


def parse_json_response_text(response_text):
    """
    Parse JSON from plain text or markdown fenced code block output.
    """
    text = (response_text or "").strip()
    if not text:
        raise ValueError("LLM returned empty response")

    if text.startswith("```"):
        lines = text.split("\n")
        json_lines = []
        in_block = False
        for line in lines:
            marker = line.strip()
            if marker.startswith("```"):
                if not in_block:
                    in_block = True
                    continue
                break
            if in_block:
                json_lines.append(line)
        if json_lines:
            text = "\n".join(json_lines).strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        fragment = extract_json_object_fragment(text)
        if fragment:
            return json.loads(fragment)
        raise ValueError("LLM response was not valid JSON.")


def extract_json_object_fragment(text):
    """Extract the first complete JSON object from a larger text blob."""
    source = str(text or "")
    start = source.find("{")
    if start == -1:
        return None

    depth = 0
    in_string = False
    escaped = False

    for idx in range(start, len(source)):
        ch = source[idx]

        if in_string:
            if escaped:
                escaped = False
            elif ch == "\\":
                escaped = True
            elif ch == '"':
                in_string = False
            continue

        if ch == '"':
            in_string = True
            continue
        if ch == "{":
            depth += 1
            continue
        if ch == "}":
            depth -= 1
            if depth == 0:
                return source[start:idx + 1]

    return None


def normalize_llm_status(raw_status):
    """
    Normalize model status text to one of the supported checklist statuses.
    """
    if not raw_status:
        return None

    status_text = str(raw_status).strip()
    if status_text in LLM_STATUS_VALUES:
        return status_text

    alias_key = status_text.lower()
    normalized = STATUS_ALIASES.get(alias_key)
    if normalized in LLM_STATUS_VALUES:
        return normalized

    return None


def canonical_doc_key(value):
    """Normalize document names for key matching."""
    text = re.sub(r'\s+', ' ', str(value or '')).strip().lower()
    return text


def normalize_api_key(api_key):
    return str(api_key or "").strip()


def normalize_provider(provider):
    value = str(provider or "").strip().lower()
    if value == "claude":
        return "anthropic"
    if value in ("anthropic", "openai", "harvey", "ollama", "lmstudio"):
        return value
    if value in ("lm studio", "lm-studio"):
        return "lmstudio"
    return "anthropic"


def infer_provider_from_api_key(api_key):
    value = normalize_api_key(api_key).lower()
    if not value:
        return None
    if value.startswith("sk-ant-"):
        return "anthropic"
    if value.startswith("harvey_") or value.startswith("hv_"):
        return "harvey"
    if value.startswith("sk-proj-") or value.startswith("sk-") or value.startswith("sess-"):
        return "openai"
    return None


def resolve_provider(provider, api_key):
    preferred = normalize_provider(provider)
    if preferred in ("ollama", "lmstudio"):
        return preferred
    inferred = infer_provider_from_api_key(api_key)
    return inferred or preferred


def provider_requires_api_key(provider):
    return normalize_provider(provider) in ("anthropic", "openai", "harvey")


def extract_openai_text(content):
    if isinstance(content, str):
        return content.strip()
    if not isinstance(content, list):
        return ""
    parts = []
    for item in content:
        if isinstance(item, str):
            parts.append(item)
            continue
        if isinstance(item, dict):
            text = item.get("text")
            if isinstance(text, str):
                parts.append(text)
    return "\n".join(parts).strip()


def parse_error_detail(raw_text, network_error=None):
    if network_error:
        return str(network_error)

    try:
        payload = json.loads(raw_text or "{}")
    except Exception:
        payload = {}

    if isinstance(payload, dict):
        err = payload.get("error")
        if isinstance(err, dict) and err.get("message"):
            return str(err.get("message"))
        if payload.get("message"):
            return str(payload.get("message"))

    return (raw_text or "").strip()[:300] or "Unknown error"


def should_retry_request(status_code, detail):
    detail_text = str(detail or "").lower()
    if status_code in RETRYABLE_HTTP_STATUS:
        return True
    return any(marker in detail_text for marker in RETRYABLE_ERROR_MARKERS)


def backoff_sleep(attempt_idx):
    # Short incremental backoff for transient API/network failures.
    time.sleep(min(0.5 * (attempt_idx + 1), 2.0))


def format_provider_error(provider_label, status_code, detail):
    if status_code in (401, 403):
        return (
            f"{provider_label} authentication failed ({status_code}). "
            f"Check API key and account/model access. Details: {detail}"
        )
    if status_code == 429:
        return f"{provider_label} rate limit reached (429). Please retry. Details: {detail}"
    if status_code == 0:
        return f"{provider_label} network error: {detail}"
    return f"{provider_label} error ({status_code}): {detail}"


def is_likely_model_error(status_code, detail):
    detail_text = str(detail or "").lower()
    if status_code == 404:
        return True
    if status_code == 400 and "model" in detail_text:
        return True
    return "model" in detail_text and any(
        part in detail_text
        for part in ("not found", "invalid", "unsupported", "available", "access")
    )


def perform_http_request(url, headers, body_bytes, timeout=90):
    req = urllib.request.Request(url, data=body_bytes, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            status_code = resp.getcode()
            raw_text = resp.read().decode("utf-8", errors="replace")
            return status_code, raw_text, None
    except urllib.error.HTTPError as e:
        raw_text = e.read().decode("utf-8", errors="replace")
        return e.code, raw_text, None
    except urllib.error.URLError as e:
        reason = getattr(e, "reason", e)
        return 0, "", f"Network error: {reason}"
    except (socket.timeout, TimeoutError):
        return 0, "", "Network error: request timed out"
    except Exception as e:
        return 0, "", f"Network error: {e}"


def perform_json_request(url, headers, payload, timeout=90):
    request_headers = dict(headers or {})
    request_headers["Content-Type"] = "application/json"
    body_bytes = json.dumps(payload).encode("utf-8")
    status_code, raw_text, network_error = perform_http_request(
        url, request_headers, body_bytes, timeout=timeout
    )
    try:
        parsed = json.loads(raw_text) if raw_text else {}
    except Exception:
        parsed = {}
    return status_code, parsed, raw_text, network_error


def build_multipart_form_data(fields):
    boundary = "----EmmaNeighBoundary" + uuid.uuid4().hex
    chunks = []
    for key, value in fields.items():
        chunks.append(f"--{boundary}".encode("utf-8"))
        chunks.append(
            f'Content-Disposition: form-data; name="{key}"'.encode("utf-8")
        )
        chunks.append(b"")
        chunks.append(str(value).encode("utf-8"))
    chunks.append(f"--{boundary}--".encode("utf-8"))
    body = b"\r\n".join(chunks) + b"\r\n"
    return boundary, body


def call_anthropic_prompt(prompt, api_key, model_name):
    api_key = normalize_api_key(api_key)
    model_candidates = []
    for candidate in (
        model_name,
        os.environ.get("CLAUDE_MODEL"),
        DEFAULT_CLAUDE_MODEL,
        "claude-3-5-sonnet-20241022",
        "claude-3-5-haiku-20241022",
    ):
        candidate_value = str(candidate or "").strip()
        if candidate_value and candidate_value not in model_candidates:
            model_candidates.append(candidate_value)

    last_detail = "Unknown error"
    for candidate in model_candidates:
        model_unavailable = False
        for attempt_idx in range(MAX_HTTP_ATTEMPTS):
            payload = {
                "model": candidate,
                "max_tokens": 2048,
                "temperature": 0,
                "messages": [{"role": "user", "content": prompt}],
            }
            status_code, parsed, raw_text, network_error = perform_json_request(
                "https://api.anthropic.com/v1/messages",
                {
                    "x-api-key": api_key,
                    "anthropic-version": "2023-06-01",
                },
                payload,
            )

            if status_code != 200:
                detail = parse_error_detail(raw_text, network_error=network_error)
                last_detail = detail
                if should_retry_request(status_code, detail) and attempt_idx < (MAX_HTTP_ATTEMPTS - 1):
                    backoff_sleep(attempt_idx)
                    continue
                if is_likely_model_error(status_code, detail):
                    model_unavailable = True
                    break
                raise RuntimeError(format_provider_error("Anthropic", status_code, detail))

            content = parsed.get("content") if isinstance(parsed, dict) else None
            if isinstance(content, list) and content:
                first = content[0]
                if isinstance(first, dict) and first.get("text"):
                    return str(first.get("text")).strip(), candidate
            raise RuntimeError("Anthropic response was missing message content.")

        if model_unavailable:
            continue

    raise RuntimeError(
        f"No supported Claude model is available for this API key. Last error: {last_detail}"
    )


def call_openai_prompt(prompt, api_key, model_name):
    api_key = normalize_api_key(api_key)
    model_candidates = []
    for candidate in (
        model_name,
        os.environ.get("OPENAI_MODEL"),
        DEFAULT_OPENAI_MODEL,
        "gpt-4.1-mini",
        "gpt-4o-mini",
        "gpt-4.1",
        "gpt-4o",
    ):
        candidate_value = str(candidate or "").strip()
        if candidate_value and candidate_value not in model_candidates:
            model_candidates.append(candidate_value)

    last_detail = "Unknown error"
    for candidate in model_candidates:
        model_unavailable = False
        for attempt_idx in range(MAX_HTTP_ATTEMPTS):
            payload = {
                "model": candidate,
                "max_tokens": 2048,
                "temperature": 0,
                "messages": [{"role": "user", "content": prompt}],
            }
            status_code, parsed, raw_text, network_error = perform_json_request(
                "https://api.openai.com/v1/chat/completions",
                {"Authorization": f"Bearer {api_key}"},
                payload,
            )

            if status_code != 200:
                detail = parse_error_detail(raw_text, network_error=network_error)
                last_detail = detail
                if should_retry_request(status_code, detail) and attempt_idx < (MAX_HTTP_ATTEMPTS - 1):
                    backoff_sleep(attempt_idx)
                    continue
                if is_likely_model_error(status_code, detail):
                    model_unavailable = True
                    break
                raise RuntimeError(format_provider_error("OpenAI", status_code, detail))

            choices = parsed.get("choices") if isinstance(parsed, dict) else None
            message = choices[0].get("message") if isinstance(choices, list) and choices else {}
            content = extract_openai_text(message.get("content") if isinstance(message, dict) else "")
            if content:
                return content, candidate
            raise RuntimeError("OpenAI response was missing message content.")

        if model_unavailable:
            continue

    raise RuntimeError(
        f"No supported OpenAI model is available for this API key. Last error: {last_detail}"
    )


def call_openai_compatible_prompt(prompt, api_key, model_candidates, base_url, provider_label):
    api_key = normalize_api_key(api_key)
    candidates = []
    for candidate in model_candidates:
        candidate_value = str(candidate or "").strip()
        if candidate_value and candidate_value not in candidates:
            candidates.append(candidate_value)

    last_detail = "Unknown error"
    for candidate in candidates:
        model_unavailable = False
        for attempt_idx in range(MAX_HTTP_ATTEMPTS):
            payload = {
                "model": candidate,
                "max_tokens": 2048,
                "temperature": 0,
                "messages": [{"role": "user", "content": prompt}],
            }
            headers = {}
            if api_key:
                headers["Authorization"] = f"Bearer {api_key}"
            status_code, parsed, raw_text, network_error = perform_json_request(
                f"{base_url}/v1/chat/completions",
                headers,
                payload,
            )

            if status_code != 200:
                detail = parse_error_detail(raw_text, network_error=network_error)
                last_detail = detail
                if should_retry_request(status_code, detail) and attempt_idx < (MAX_HTTP_ATTEMPTS - 1):
                    backoff_sleep(attempt_idx)
                    continue
                if is_likely_model_error(status_code, detail):
                    model_unavailable = True
                    break
                raise RuntimeError(format_provider_error(provider_label, status_code, detail))

            choices = parsed.get("choices") if isinstance(parsed, dict) else None
            message = choices[0].get("message") if isinstance(choices, list) and choices else {}
            content = extract_openai_text(message.get("content") if isinstance(message, dict) else "")
            if content:
                return content, candidate
            raise RuntimeError(f"{provider_label} response was missing message content.")

        if model_unavailable:
            continue

    raise RuntimeError(
        f"No supported {provider_label} model is available. Last error: {last_detail}"
    )


def call_ollama_prompt(prompt, api_key, model_name, provider_base_url):
    base_url = str(provider_base_url or os.environ.get("OLLAMA_BASE_URL") or DEFAULT_OLLAMA_BASE_URL).strip()
    if not base_url:
        base_url = DEFAULT_OLLAMA_BASE_URL
    if base_url.endswith("/"):
        base_url = base_url[:-1]
    return call_openai_compatible_prompt(
        prompt,
        api_key,
        [
            model_name,
            os.environ.get("OLLAMA_MODEL"),
            DEFAULT_OLLAMA_MODEL,
            "llama3.1:8b",
            "qwen2.5:7b",
        ],
        base_url,
        "Ollama",
    )


def call_lmstudio_prompt(prompt, api_key, model_name, provider_base_url):
    base_url = str(provider_base_url or os.environ.get("LMSTUDIO_BASE_URL") or DEFAULT_LMSTUDIO_BASE_URL).strip()
    if not base_url:
        base_url = DEFAULT_LMSTUDIO_BASE_URL
    if base_url.endswith("/"):
        base_url = base_url[:-1]
    return call_openai_compatible_prompt(
        prompt,
        api_key,
        [
            model_name,
            os.environ.get("LMSTUDIO_MODEL"),
            DEFAULT_LMSTUDIO_MODEL,
        ],
        base_url,
        "LM Studio",
    )


def call_harvey_prompt(prompt, api_key, max_tokens, provider_base_url):
    api_key = normalize_api_key(api_key)
    base_url = str(
        provider_base_url or os.environ.get("HARVEY_BASE_URL") or DEFAULT_HARVEY_BASE_URL
    ).strip()
    if not base_url:
        base_url = DEFAULT_HARVEY_BASE_URL
    if base_url.endswith("/"):
        base_url = base_url[:-1]

    boundary, body = build_multipart_form_data(
        {
            "prompt": prompt,
            "mode": "assist",
            "stream": "false",
            "max_tokens": str(max_tokens),
        }
    )

    parsed = {}
    for attempt_idx in range(MAX_HTTP_ATTEMPTS):
        status_code, raw_text, network_error = perform_http_request(
            f"{base_url}/api/v2/completion?include_citations=false",
            {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": f"multipart/form-data; boundary={boundary}",
            },
            body,
        )
        detail = parse_error_detail(raw_text, network_error=network_error)
        if status_code != 200:
            if should_retry_request(status_code, detail) and attempt_idx < (MAX_HTTP_ATTEMPTS - 1):
                backoff_sleep(attempt_idx)
                continue
            raise RuntimeError(format_provider_error("Harvey", status_code, detail))

        try:
            parsed = json.loads(raw_text) if raw_text else {}
        except Exception:
            parsed = {}
        break

    text = ""
    if isinstance(parsed, dict):
        text = parsed.get("response") or parsed.get("text") or ""
    if not text:
        raise RuntimeError("Harvey response was missing message content.")

    return str(text).strip(), "harvey-assist"


def call_provider_prompt_json(prompt, api_key, provider, model_name=None, provider_base_url=None):
    provider_name = resolve_provider(provider, api_key)
    if provider_name == "openai":
        response_text, _ = call_openai_prompt(prompt, api_key, model_name)
    elif provider_name == "harvey":
        response_text, _ = call_harvey_prompt(prompt, api_key, 2048, provider_base_url)
    elif provider_name == "ollama":
        response_text, _ = call_ollama_prompt(prompt, api_key, model_name, provider_base_url)
    elif provider_name == "lmstudio":
        response_text, _ = call_lmstudio_prompt(prompt, api_key, model_name, provider_base_url)
    else:
        response_text, _ = call_anthropic_prompt(prompt, api_key, model_name)

    parsed = parse_json_response_text(response_text)
    if not isinstance(parsed, dict):
        raise ValueError("LLM response was not a JSON object.")
    return parsed


def parse_checklist_table(doc):
    """
    Parse the first table in a Word document as a checklist.

    Returns:
        - headers: list of column headers
        - rows: list of row data (each row is a list of cell values)
        - table: the docx table object for modification
        - doc_col_idx: index of document name column
        - status_col_idx: index of status column
    """
    if not doc.tables:
        return None, None, None, -1, -1

    table = doc.tables[0]  # Use first table

    # Extract all rows
    all_rows = []
    for row in table.rows:
        row_data = [cell.text.strip() for cell in row.cells]
        all_rows.append(row_data)

    if not all_rows:
        return None, None, None, -1, -1

    headers = all_rows[0]
    rows = all_rows[1:]

    # Find document and status columns
    doc_col_idx = find_column_index(headers, DOCUMENT_COLUMN_PATTERNS)
    status_col_idx = find_column_index(headers, STATUS_COLUMN_PATTERNS)

    # If no status column found, we'll add one
    if status_col_idx == -1:
        # Look for a column we can repurpose or note that we need to add one
        status_col_idx = len(headers)  # Will add new column

    return headers, rows, table, doc_col_idx, status_col_idx


def detect_document_status(doc_name, emails):
    """
    Search emails for mentions of a document and detect its status.

    Returns:
        - status: detected status string or None
        - priority: status priority (higher = more advanced)
        - matching_emails: list of email subjects that matched
    """
    if not doc_name:
        return None, 0, []

    doc_name_lower = doc_name.lower()

    # Create search patterns from document name
    # Handle common variations
    doc_words = re.findall(r'\b\w+\b', doc_name_lower)

    best_status = None
    best_priority = 0
    matching_emails = []

    for email in emails:
        searchable = email.get('searchable', '')

        # Check if email mentions this document
        # Use fuzzy matching - at least 2 key words must match
        word_matches = sum(1 for word in doc_words if len(word) > 3 and word in searchable)

        if word_matches < 2 and doc_name_lower not in searchable:
            continue

        # Email mentions this document - check for status patterns
        for status_config in STATUS_PATTERNS:
            for pattern in status_config['patterns']:
                if re.search(pattern, searchable):
                    if status_config['priority'] > best_priority:
                        best_status = status_config['status']
                        best_priority = status_config['priority']
                        matching_emails.append(email.get('subject', 'No subject'))
                    break

    return best_status, best_priority, matching_emails


def match_documents_with_llm(
    checklist_items,
    emails,
    api_key,
    provider="anthropic",
    model_name=None,
    provider_base_url=None,
):
    """
    Use the selected LLM provider to match emails to documents and infer status.

    Args:
        checklist_items: List of document names from checklist
        emails: List of email dicts with subject, body, from, date
        api_key: Provider API key
        provider: LLM provider (anthropic/openai/harvey)
        model_name: Optional model override
        provider_base_url: Optional provider API base URL

    Returns:
        Dict mapping document_name to {status, matching_emails, confidence}
    """
    if provider_requires_api_key(provider) and not api_key:
        raise ValueError("API key is required for checklist LLM analysis.")

    # Prepare email context (limit and truncate for token efficiency)
    email_context = []
    for i, email in enumerate(emails[:100]):
        email_context.append({
            "index": i,
            "from": email.get("from", "")[:100],
            "subject": email.get("subject", "")[:200],
            "body_preview": email.get("body", "")[:200],
            "date": email.get("date_received") or email.get("date_sent") or email.get("date") or ""
        })

    prompt = f"""You are analyzing emails to update a transaction document checklist.

For each document in the checklist, find relevant emails and determine the current status.

CHECKLIST DOCUMENTS:
{json.dumps(checklist_items, indent=2)}

RECENT EMAILS:
{json.dumps(email_context, indent=2)}

For each document that has relevant email activity, determine its status from these options:
- "Pending Draft" (not started, needs drafting)
- "Draft Circulated" (initial draft sent out)
- "With Opposing Counsel" (sent to counterparty for review)
- "Agreed Form" (parties have agreed on the form)
- "Execution Version" (ready for signature)
- "Executed" (fully signed)

Return a JSON object mapping each document name (only those with email activity) to:
{{
    "Document Name": {{
        "status": "Execution Version",
        "matching_email_indices": [2, 15],
        "confidence": 0.85,
        "reasoning": "Email #2 mentions execution version is ready..."
    }}
}}

ONLY return the JSON object, no other text."""

    return call_provider_prompt_json(
        prompt,
        api_key,
        provider,
        model_name=model_name,
        provider_base_url=provider_base_url,
    )


def update_checklist(
    checklist_path,
    email_csv_path,
    output_folder,
    api_key=None,
    provider="anthropic",
    model_name=None,
    provider_base_url=None,
):
    """
    Main function to update checklist based on email activity.

    Args:
        checklist_path: Path to Word document with checklist table
        email_csv_path: Path to Outlook email CSV export
        output_folder: Folder to save updated checklist
        api_key: Provider API key
        provider: LLM provider (anthropic/openai/harvey)
        model_name: Optional provider model override
        provider_base_url: Optional provider API base URL

    Returns:
        dict with: success, output_path, items_updated, details
    """
    result = {
        'success': False,
        'output_path': None,
        'items_updated': 0,
        'details': [],
        'error': None
    }

    try:
        api_key = normalize_api_key(api_key)
        provider = resolve_provider(provider, api_key)

        if provider_requires_api_key(provider) and not api_key:
            result['error'] = 'LLM API key is required to update checklist from email activity.'
            return result

        # Parse email CSV
        emails = parse_email_csv(email_csv_path)
        if not emails:
            result['error'] = 'No emails found in CSV file'
            return result

        # Open checklist document
        doc = Document(checklist_path)

        # Parse the checklist table
        headers, rows, table, doc_col_idx, status_col_idx = parse_checklist_table(doc)

        if table is None:
            result['error'] = 'No table found in checklist document'
            return result

        if doc_col_idx == -1:
            result['error'] = 'Could not identify document name column in checklist'
            return result

        # Check if we need to add a status column
        needs_new_status_col = status_col_idx >= len(headers)

        if needs_new_status_col:
            # Add "Status" header to first row
            header_row = table.rows[0]
            # Add new cell (this is complex in python-docx, so we'll update existing empty column if available)
            # For now, we'll look for an empty column or skip
            for i, header in enumerate(headers):
                if not header.strip():
                    status_col_idx = i
                    header_row.cells[i].text = "Status (Updated)"
                    needs_new_status_col = False
                    break

            if needs_new_status_col:
                result['error'] = 'No status column found and cannot add new column. Please add a Status column to your checklist.'
                return result

        # Collect all document names for LLM matching
        doc_names = []
        for row_data in rows:
            if doc_col_idx >= len(row_data):
                continue
            doc_name = row_data[doc_col_idx]
            if doc_name.strip():
                doc_names.append(doc_name)

        if not doc_names:
            result['error'] = 'No checklist document names found to analyze.'
            return result

        # LLM matching is mandatory for checklist updates in this workflow.
        try:
            llm_matches = match_documents_with_llm(
                doc_names,
                emails,
                api_key,
                provider=provider,
                model_name=model_name,
                provider_base_url=provider_base_url,
            )
        except Exception as e:
            result['error'] = f'LLM checklist analysis failed: {e}'
            return result
        llm_matches_by_key = {
            canonical_doc_key(doc_name): match_data
            for doc_name, match_data in (llm_matches or {}).items()
        }

        # Process each row
        items_updated = 0
        details = []

        for row_idx, row_data in enumerate(rows):
            if doc_col_idx >= len(row_data):
                continue

            doc_name = row_data[doc_col_idx]
            if not doc_name.strip():
                continue

            # Get current status
            current_status = ''
            if status_col_idx < len(row_data):
                current_status = row_data[status_col_idx]

            # LLM-driven match only; no regex/rules fallback in this workflow.
            new_status = None
            matching_emails = []

            llm_result = llm_matches.get(doc_name) if llm_matches else None
            if not llm_result:
                llm_result = llm_matches_by_key.get(canonical_doc_key(doc_name))

            if llm_result:
                new_status = normalize_llm_status(llm_result.get('status'))
                # Get matching email subjects from indices
                email_indices = llm_result.get('matching_email_indices', [])
                for idx in email_indices[:3]:
                    if idx < len(emails):
                        matching_emails.append(emails[idx].get('subject', 'No subject'))

            if new_status and new_status != current_status:
                # Update the cell in the table
                table_row = table.rows[row_idx + 1]  # +1 to skip header
                if status_col_idx < len(table_row.cells):
                    cell = table_row.cells[status_col_idx]

                    # Preserve formatting, just update text
                    if cell.paragraphs:
                        cell.paragraphs[0].text = new_status
                    else:
                        cell.text = new_status

                    items_updated += 1
                    details.append({
                        'document': doc_name,
                        'old_status': current_status,
                        'new_status': new_status,
                        'emails': matching_emails[:3]  # Limit to 3 examples
                    })

        # Save updated document
        os.makedirs(output_folder, exist_ok=True)

        base_name = os.path.splitext(os.path.basename(checklist_path))[0]
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"{base_name}_updated_{timestamp}.docx"
        output_path = os.path.join(output_folder, output_filename)

        doc.save(output_path)

        result['success'] = True
        result['output_path'] = output_path
        result['items_updated'] = items_updated
        result['details'] = details
        result['emails_processed'] = len(emails)
        result['llm_documents_with_activity'] = len(llm_matches or {})

        return result

    except Exception as e:
        result['error'] = str(e)
        return result


def main():
    if len(sys.argv) < 4:
        print(json.dumps({
            'success': False,
            'error': 'Usage: checklist_updater.py <checklist_path> <email_csv_path> <output_folder> <api_key> [provider] [model] [provider_base_url]'
        }))
        sys.exit(1)

    checklist_path = sys.argv[1]
    email_csv_path = sys.argv[2]
    output_folder = sys.argv[3]

    # API key is required for LLM-driven checklist updates.
    api_key = sys.argv[4] if len(sys.argv) > 4 else (
        os.environ.get('ANTHROPIC_API_KEY')
        or os.environ.get('OPENAI_API_KEY')
        or os.environ.get('HARVEY_API_KEY')
    )
    provider = resolve_provider(
        sys.argv[5] if len(sys.argv) > 5 else os.environ.get('AI_PROVIDER', 'anthropic'),
        api_key,
    )
    model_name = sys.argv[6] if len(sys.argv) > 6 else (
        os.environ.get('OPENAI_MODEL') if provider == 'openai'
        else os.environ.get('OLLAMA_MODEL') if provider == 'ollama'
        else os.environ.get('LMSTUDIO_MODEL') if provider == 'lmstudio'
        else os.environ.get('CLAUDE_MODEL')
    )
    provider_base_url = sys.argv[7] if len(sys.argv) > 7 else (
        os.environ.get('HARVEY_BASE_URL', DEFAULT_HARVEY_BASE_URL) if provider == 'harvey'
        else os.environ.get('OLLAMA_BASE_URL', DEFAULT_OLLAMA_BASE_URL) if provider == 'ollama'
        else os.environ.get('LMSTUDIO_BASE_URL', DEFAULT_LMSTUDIO_BASE_URL) if provider == 'lmstudio'
        else None
    )

    result = update_checklist(
        checklist_path,
        email_csv_path,
        output_folder,
        api_key,
        provider=provider,
        model_name=model_name,
        provider_base_url=provider_base_url,
    )
    print(json.dumps(result))


if __name__ == '__main__':
    main()
