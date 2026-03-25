#!/usr/bin/env python3
"""
Checklist Updater - Update transaction checklist based on email activity

This module parses a transaction checklist (Word document) and an email activity dataset,
then updates the checklist status column based on detected email activity.

Usage:
    python checklist_updater.py <checklist_path> <email_input_path> <output_folder> <api_key> [provider] [model] [provider_base_url]
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
DEFAULT_LOCALAI_BASE_URL = "http://127.0.0.1:11435"
DEFAULT_OLLAMA_BASE_URL = "http://127.0.0.1:11434"
DEFAULT_LMSTUDIO_BASE_URL = "http://127.0.0.1:1234"
DEFAULT_LOCALAI_MODEL = "qwen2.5:1.5b"
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

PARTY_COLUMN_PATTERNS = [
    r'^party',
    r'^responsible',
    r'^owner',
    r'^assignee',
    r'^who',
    r'^counsel',
]

NOTES_COLUMN_PATTERNS = [
    r'^notes?',
    r'^comments?',
    r'^remarks?',
    r'^details?',
]


def normalize_email_record(raw_email):
    """Normalize one email record from CSV or JSON input."""
    if not isinstance(raw_email, dict):
        return None

    attachments = raw_email.get('attachments', '')
    if isinstance(attachments, list):
        attachment_text = '; '.join(str(item).strip() for item in attachments if str(item).strip())
    else:
        attachment_text = str(attachments or '').strip()

    has_attachments = raw_email.get('has_attachments')
    if isinstance(has_attachments, str):
        has_attachments = has_attachments.strip().lower() not in ('', '0', 'false', 'no', 'none', 'nan')
    else:
        has_attachments = bool(has_attachments)
    if not has_attachments and attachment_text:
        has_attachments = attachment_text.lower() not in ('none', 'false', 'no', 'n/a')

    email = {
        'subject': str(raw_email.get('subject', '') or '').strip(),
        'body': str(raw_email.get('body', '') or '').strip(),
        'from': str(raw_email.get('from', raw_email.get('sender', raw_email.get('sender_email', ''))) or '').strip(),
        'to': str(raw_email.get('to', '') or '').strip(),
        'cc': str(raw_email.get('cc', '') or '').strip(),
        'attachments': attachment_text,
        'has_attachments': has_attachments,
        'date': str(raw_email.get('date', '') or '').strip(),
    }

    sent_value = str(raw_email.get('date_sent', raw_email.get('sent', '')) or '').strip()
    received_value = str(raw_email.get('date_received', raw_email.get('received', '')) or '').strip()
    if sent_value:
        email['date_sent'] = sent_value
    else:
        email['date_sent'] = ''
    if received_value:
        email['date_received'] = received_value
    else:
        email['date_received'] = ''
    if not email['date']:
        email['date'] = email['date_received'] or email['date_sent']

    email['searchable'] = " ".join([
        email['subject'],
        email['body'],
        email['from'],
        email['to'],
        email['cc'],
        email['attachments'],
    ]).lower()
    return email


def parse_email_csv(csv_path):
    """
    Parse Outlook email CSV export.

    Returns list of email dicts with normalized metadata for LLM analysis:
    subject, body, from, to, cc, date, attachments, has_attachments, searchable
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

        # Map common Outlook/Graph export column names.
        column_mapping = {
            'subject': ['subject', 'email subject', 'title'],
            'body': ['body', 'content', 'message', 'email body', 'notes'],
            'from': ['from', 'sender', 'from email', 'from address'],
            'to': ['to', 'recipient', 'to email', 'to address', 'recipients'],
            'cc': ['cc', 'cc recipients', 'cc list'],
            'attachments': ['attachments', 'attachment', 'files', 'file names', 'attachment names'],
            'has_attachments': ['has attachments', 'has attachment', 'attachments?', 'hasattachments'],
            'date': ['date', 'sent', 'received', 'date sent', 'date received', 'sent date', 'received time', 'sent time'],
        }

        emails = []
        for _, row in df.iterrows():
            raw_email = {}
            for field, possible_cols in column_mapping.items():
                for col in possible_cols:
                    if col in df.columns and pd.notna(row.get(col)):
                        raw_value = str(row[col]).strip()
                        if field == 'has_attachments':
                            raw_email[field] = raw_value.lower() not in ('', '0', 'false', 'no', 'none', 'nan')
                        else:
                            raw_email[field] = raw_value
                        break
                else:
                    raw_email[field] = False if field == 'has_attachments' else ''

            email = normalize_email_record(raw_email)
            if email:
                emails.append(email)

        return emails

    except Exception as e:
        print(f"Error parsing email CSV: {e}", file=sys.stderr)
        return []


def parse_email_json(json_path):
    """Parse JSON email activity written by the Electron app."""
    try:
        with open(json_path, 'r', encoding='utf-8') as handle:
            payload = json.load(handle)
    except Exception as e:
        print(f"Error parsing email JSON: {e}", file=sys.stderr)
        return []

    if isinstance(payload, dict):
        raw_emails = payload.get('emails', [])
    elif isinstance(payload, list):
        raw_emails = payload
    else:
        raw_emails = []

    emails = []
    for raw_email in raw_emails:
        email = normalize_email_record(raw_email)
        if email:
            emails.append(email)
    return emails


def parse_email_input(email_input_path):
    """Load email activity from either CSV export or JSON dataset."""
    ext = os.path.splitext(str(email_input_path or ''))[1].lower()
    if ext == '.json':
        return parse_email_json(email_input_path)
    return parse_email_csv(email_input_path)


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
    if value in ("managed", "localai", "anthropic", "openai", "harvey", "ollama", "lmstudio"):
        return value
    if value in ("managed ai", "managed-ai", "managed remote", "managed-remote", "firm", "firm-managed", "firmmanaged"):
        return "managed"
    if value in ("local", "local ai", "local-ai", "local pack", "localpack", "built-in", "builtin"):
        return "localai"
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
    if preferred in ("managed", "localai", "ollama", "lmstudio"):
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


def call_localai_prompt(prompt, api_key, model_name, provider_base_url):
    base_url = str(provider_base_url or os.environ.get("LOCALAI_BASE_URL") or DEFAULT_LOCALAI_BASE_URL).strip()
    if not base_url:
        base_url = DEFAULT_LOCALAI_BASE_URL
    if base_url.endswith("/"):
        base_url = base_url[:-1]
    return call_openai_compatible_prompt(
        prompt,
        api_key,
        [
            model_name,
            os.environ.get("LOCALAI_MODEL"),
            DEFAULT_LOCALAI_MODEL,
            "qwen2.5:1.5b",
            "qwen2.5:7b",
            "gemma3:1b",
            "gemma3",
        ],
        base_url,
        "Local AI",
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


def call_managed_prompt(prompt, api_key, model_name, provider_base_url):
    endpoint = str(provider_base_url or os.environ.get("MANAGED_AI_PROMPT_URL") or "").strip()
    if not endpoint:
        raise RuntimeError("Managed AI prompt endpoint is not configured.")

    headers = {}
    token = normalize_api_key(api_key)
    if token:
        headers["Authorization"] = f"Bearer {token}"

    payload = {
        "prompt": prompt,
        "max_tokens": 2048,
    }
    if model_name:
        payload["model"] = str(model_name).strip()

    last_detail = "Unknown error"
    for attempt_idx in range(MAX_HTTP_ATTEMPTS):
        status_code, parsed, raw_text, network_error = perform_json_request(
            endpoint,
            headers,
            payload,
        )

        if status_code != 200:
            detail = parse_error_detail(raw_text, network_error=network_error)
            last_detail = detail
            if should_retry_request(status_code, detail) and attempt_idx < (MAX_HTTP_ATTEMPTS - 1):
                backoff_sleep(attempt_idx)
                continue
            raise RuntimeError(format_provider_error("Managed AI", status_code, detail))

        text = ""
        if isinstance(parsed, dict):
            text = str(parsed.get("text") or parsed.get("message") or "").strip()
        if text:
            return text, str(parsed.get("model") or model_name or "managed-ai").strip()
        raise RuntimeError("Managed AI response was missing message content.")

    raise RuntimeError(f"Managed AI error: {last_detail}")


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
    if provider_name == "managed":
        response_text, _ = call_managed_prompt(prompt, api_key, model_name, provider_base_url)
    elif provider_name == "localai":
        response_text, _ = call_localai_prompt(prompt, api_key, model_name, provider_base_url)
    elif provider_name == "openai":
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


def normalize_cell_text(value):
    return re.sub(r'\s+', ' ', str(value or '')).strip()


def infer_document_column(headers, rows):
    """
    Infer the document/title column when header names differ across templates/firms.
    """
    if not headers:
        return -1

    max_cols = max(len(headers), max((len(r) for r in rows), default=0))
    best_col = -1
    best_score = -1
    blocked_terms = ('status', 'state', 'progress', 'date', 'sent', 'received')

    for col_idx in range(max_cols):
        header_value = normalize_cell_text(headers[col_idx] if col_idx < len(headers) else '')
        header_lower = header_value.lower()
        if any(term in header_lower for term in blocked_terms):
            continue

        non_empty = 0
        rich_text = 0
        penalties = 0
        for row in rows[:150]:
            cell = normalize_cell_text(row[col_idx] if col_idx < len(row) else '')
            if not cell:
                continue
            non_empty += 1
            if re.search(r'[a-zA-Z]', cell):
                rich_text += 1
            if len(cell) > 220 or re.fullmatch(r'[\d\W_]+', cell):
                penalties += 1

        score = (non_empty * 2) + (rich_text * 3) - (penalties * 2)
        if score > best_score:
            best_score = score
            best_col = col_idx

    return best_col


def score_header_candidate(all_rows, header_idx):
    headers = all_rows[header_idx]
    non_empty_headers = [normalize_cell_text(h) for h in headers if normalize_cell_text(h)]
    if len(non_empty_headers) < 2:
        return None

    doc_col = find_column_index(headers, DOCUMENT_COLUMN_PATTERNS)
    status_col = find_column_index(headers, STATUS_COLUMN_PATTERNS)
    party_col = find_column_index(headers, PARTY_COLUMN_PATTERNS)
    notes_col = find_column_index(headers, NOTES_COLUMN_PATTERNS)

    avg_header_len = (
        sum(len(x) for x in non_empty_headers) / len(non_empty_headers)
        if non_empty_headers else 0
    )
    looks_like_header_text = 1 if avg_header_len <= 40 else -5

    # Score based on quality of the first rows after this candidate header.
    data_rows = all_rows[header_idx + 1: header_idx + 11]
    non_empty_data_rows = sum(1 for row in data_rows if any(normalize_cell_text(c) for c in row))

    score = len(non_empty_headers) + (non_empty_data_rows * 3) + looks_like_header_text
    if doc_col != -1:
        score += 40
    if status_col != -1:
        score += 20
    if party_col != -1:
        score += 6
    if notes_col != -1:
        score += 4

    return {
        'score': score,
        'header_idx': header_idx,
        'doc_col': doc_col,
        'status_col': status_col,
    }


def parse_checklist_table(doc):
    """
    Parse a checklist-like table from the document, allowing flexible header position
    and non-standard column names.

    Returns:
        headers, rows, table, doc_col_idx, status_col_idx, data_row_start_idx
    """
    if not doc.tables:
        return None, None, None, -1, -1, -1

    best_candidate = None

    for table in doc.tables:
        all_rows = []
        for row in table.rows:
            all_rows.append([normalize_cell_text(cell.text) for cell in row.cells])

        if not all_rows:
            continue

        max_probe = min(6, len(all_rows) - 1)
        for header_idx in range(max_probe + 1):
            candidate = score_header_candidate(all_rows, header_idx)
            if not candidate:
                continue
            candidate['all_rows'] = all_rows
            candidate['table'] = table
            if best_candidate is None or candidate['score'] > best_candidate['score']:
                best_candidate = candidate

    if not best_candidate:
        return None, None, None, -1, -1, -1

    table = best_candidate['table']
    all_rows = best_candidate['all_rows']
    header_idx = best_candidate['header_idx']
    headers = all_rows[header_idx]
    rows = [row for row in all_rows[header_idx + 1:] if any(normalize_cell_text(c) for c in row)]

    doc_col_idx = best_candidate['doc_col']
    if doc_col_idx == -1:
        doc_col_idx = infer_document_column(headers, rows)

    status_col_idx = best_candidate['status_col']
    if status_col_idx == -1:
        status_col_idx = len(headers)  # Signals "no existing status column".

    return headers, rows, table, doc_col_idx, status_col_idx, header_idx + 1


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


def _parse_row_id(value):
    try:
        return int(str(value).strip())
    except Exception:
        return None


def normalize_checklist_llm_matches(raw_payload):
    """
    Normalize LLM outputs into a stable structure:
      {
        "by_row": { row_id: {...} },
        "by_doc": { canonical_doc_key: {...} }
      }
    Supports both legacy doc-name keyed responses and new row-id keyed payloads.
    """
    normalized = {"by_row": {}, "by_doc": {}}
    payload = raw_payload if isinstance(raw_payload, dict) else {}

    def coerce_match(entry, default_doc_name=None):
        if not isinstance(entry, dict):
            return None
        status = normalize_llm_status(entry.get("status"))
        if not status:
            return None
        raw_indices = entry.get("matching_email_indices", [])
        email_indices = []
        if isinstance(raw_indices, list):
            for idx in raw_indices[:15]:
                try:
                    parsed_idx = int(idx)
                except Exception:
                    continue
                if parsed_idx >= 0:
                    email_indices.append(parsed_idx)

        confidence = entry.get("confidence", 0.5)
        try:
            confidence = float(confidence)
        except Exception:
            confidence = 0.5
        confidence = max(0.0, min(confidence, 1.0))

        doc_name = str(entry.get("document_name") or default_doc_name or "").strip()
        row_id = _parse_row_id(entry.get("row_id"))

        return {
            "row_id": row_id,
            "document_name": doc_name,
            "status": status,
            "matching_email_indices": email_indices,
            "confidence": confidence,
            "reasoning": str(entry.get("reasoning") or "").strip(),
        }

    # New format: { "matches": [ ... ] }
    if isinstance(payload.get("matches"), list):
        for item in payload.get("matches", []):
            parsed = coerce_match(item)
            if not parsed:
                continue
            row_id = parsed.get("row_id")
            doc_name = parsed.get("document_name", "")
            if row_id is not None:
                normalized["by_row"][row_id] = parsed
            if doc_name:
                normalized["by_doc"][canonical_doc_key(doc_name)] = parsed

    # Legacy format: { "Document Name": { ... } }
    for key, value in payload.items():
        if key == "matches":
            continue
        parsed = coerce_match(value, default_doc_name=key)
        if not parsed:
            continue
        row_id = parsed.get("row_id")
        doc_name = parsed.get("document_name", "") or str(key)
        if row_id is not None:
            normalized["by_row"][row_id] = parsed
        normalized["by_doc"][canonical_doc_key(doc_name)] = parsed

    return normalized


def match_documents_with_llm(
    checklist_items,
    emails,
    api_key,
    provider="anthropic",
    model_name=None,
    provider_base_url=None,
):
    """
    Use the selected LLM provider to match checklist rows to email activity and infer status.
    """
    if provider_requires_api_key(provider) and not api_key:
        raise ValueError("API key is required for checklist LLM analysis.")

    # Prepare richer email context so the model can infer "sent", "under review", etc.
    email_context = []
    for i, email in enumerate(emails[:180]):
        email_context.append({
            "index": i,
            "from": str(email.get("from", "") or "")[:160],
            "to": str(email.get("to", "") or "")[:220],
            "cc": str(email.get("cc", "") or "")[:220],
            "subject": str(email.get("subject", "") or "")[:260],
            "attachments": str(email.get("attachments", "") or "")[:260],
            "has_attachments": bool(email.get("has_attachments", False)),
            "date": email.get("date_received") or email.get("date_sent") or email.get("date") or "",
            "body_preview": str(email.get("body", "") or "")[:700],
        })

    checklist_context = []
    for item in checklist_items[:220]:
        checklist_context.append({
            "row_id": item.get("row_id"),
            "document_name": item.get("document_name", ""),
            "current_status": item.get("current_status", ""),
            "row_context": item.get("row_context", ""),
        })

    prompt = f"""You are a legal transaction assistant updating a checklist from email evidence.

Your task is to determine whether each checklist row shows real document activity in email and, if so, infer the best current status.

CHECKLIST ROWS:
{json.dumps(checklist_context, indent=2)}

EMAILS:
{json.dumps(email_context, indent=2)}

Status options (use exactly one):
- "Pending Draft"
- "Draft Circulated"
- "With Opposing Counsel"
- "Agreed Form"
- "Execution Version"
- "Executed"

Interpretation guidance:
- If a draft or markup was sent but external review is not clear: "Draft Circulated".
- If it was sent to counterparty/opposing counsel/client for comments/review: "With Opposing Counsel".
- If form is settled/final with no material open comments: "Agreed Form".
- If ready for signatures or signature pages circulated: "Execution Version".
- If fully signed / executed copies circulated: "Executed".

Matching guidance:
- Match by semantic meaning, abbreviations, related phrasing, and attachment names.
- Do NOT require exact column/header text matches.
- Ignore unrelated admin emails.

Return JSON only in this exact shape:
{{
  "matches": [
    {{
      "row_id": 12,
      "document_name": "Credit Agreement",
      "status": "With Opposing Counsel",
      "matching_email_indices": [3, 9],
      "confidence": 0.82,
      "reasoning": "Email 3 sends draft to counterparty counsel for review."
    }}
  ]
}}

Include only rows with meaningful evidence from email. Do not include rows with no evidence.
"""

    parsed = call_provider_prompt_json(
        prompt,
        api_key,
        provider,
        model_name=model_name,
        provider_base_url=provider_base_url,
    )
    return normalize_checklist_llm_matches(parsed)


def update_checklist(
    checklist_path,
    email_input_path,
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
        email_input_path: Path to CSV or JSON email activity input
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

        # Parse email activity
        emails = parse_email_input(email_input_path)
        if not emails:
            result['error'] = 'No emails found in the supplied email activity.'
            return result

        # Open checklist document
        doc = Document(checklist_path)

        # Parse the checklist table
        headers, rows, table, doc_col_idx, status_col_idx, data_row_start_idx = parse_checklist_table(doc)

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
            header_row_idx = data_row_start_idx - 1 if data_row_start_idx > 0 else 0
            header_row = table.rows[header_row_idx]
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

        # Build checklist row payloads for semantic LLM matching.
        checklist_items = []
        for row_idx, row_data in enumerate(rows):
            if doc_col_idx < 0 or doc_col_idx >= len(row_data):
                continue
            doc_name = normalize_cell_text(row_data[doc_col_idx])
            if not doc_name:
                continue

            current_status = ''
            if status_col_idx < len(row_data):
                current_status = normalize_cell_text(row_data[status_col_idx])

            context_parts = []
            for col_idx, cell_value in enumerate(row_data):
                if col_idx in (doc_col_idx, status_col_idx):
                    continue
                cleaned = normalize_cell_text(cell_value)
                if not cleaned:
                    continue
                header_label = normalize_cell_text(headers[col_idx] if col_idx < len(headers) else f"Column {col_idx + 1}")
                if header_label:
                    context_parts.append(f"{header_label}: {cleaned}")
                else:
                    context_parts.append(cleaned)

            checklist_items.append({
                'row_id': data_row_start_idx + row_idx,
                'document_name': doc_name,
                'current_status': current_status,
                'row_context': ' | '.join(context_parts[:8]),
                'row_data': row_data,
            })

        if not checklist_items:
            result['error'] = 'No checklist document names found to analyze.'
            return result

        # LLM matching is mandatory for checklist updates in this workflow.
        try:
            llm_matches = match_documents_with_llm(
                checklist_items,
                emails,
                api_key,
                provider=provider,
                model_name=model_name,
                provider_base_url=provider_base_url,
            )
        except Exception as e:
            result['error'] = f'LLM checklist analysis failed: {e}'
            return result

        # Process each row
        items_updated = 0
        details = []

        for item in checklist_items:
            doc_name = item['document_name']
            row_data = item.get('row_data', [])
            row_id = item.get('row_id')

            current_status = item.get('current_status', '')
            new_status = None
            matching_emails = []

            llm_result = None
            if isinstance(llm_matches, dict):
                by_row = llm_matches.get('by_row', {})
                by_doc = llm_matches.get('by_doc', {})
                llm_result = by_row.get(row_id) if isinstance(by_row, dict) else None
                if not llm_result and isinstance(by_doc, dict):
                    llm_result = by_doc.get(canonical_doc_key(doc_name))

            if llm_result:
                new_status = normalize_llm_status(llm_result.get('status'))
                # Get matching email subjects from indices
                email_indices = llm_result.get('matching_email_indices', [])
                for idx in email_indices[:3]:
                    if 0 <= idx < len(emails):
                        matching_emails.append(emails[idx].get('subject', 'No subject'))

            if new_status and new_status != current_status:
                # Update the cell in the table
                if row_id is None or row_id < 0 or row_id >= len(table.rows):
                    continue
                table_row = table.rows[row_id]
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
                        'confidence': llm_result.get('confidence') if isinstance(llm_result, dict) else None,
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
        result['llm_documents_with_activity'] = len((llm_matches or {}).get('by_row', {}))

        return result

    except Exception as e:
        result['error'] = str(e)
        return result


def main():
    if len(sys.argv) < 4:
        print(json.dumps({
            'success': False,
            'error': 'Usage: checklist_updater.py <checklist_path> <email_input_path> <output_folder> <api_key> [provider] [model] [provider_base_url]'
        }))
        sys.exit(1)

    checklist_path = sys.argv[1]
    email_input_path = sys.argv[2]
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
        else os.environ.get('LOCALAI_MODEL') if provider == 'localai'
        else os.environ.get('OLLAMA_MODEL') if provider == 'ollama'
        else os.environ.get('LMSTUDIO_MODEL') if provider == 'lmstudio'
        else os.environ.get('CLAUDE_MODEL')
    )
    provider_base_url = sys.argv[7] if len(sys.argv) > 7 else (
        os.environ.get('HARVEY_BASE_URL', DEFAULT_HARVEY_BASE_URL) if provider == 'harvey'
        else os.environ.get('LOCALAI_BASE_URL', DEFAULT_LOCALAI_BASE_URL) if provider == 'localai'
        else os.environ.get('OLLAMA_BASE_URL', DEFAULT_OLLAMA_BASE_URL) if provider == 'ollama'
        else os.environ.get('LMSTUDIO_BASE_URL', DEFAULT_LMSTUDIO_BASE_URL) if provider == 'lmstudio'
        else None
    )

    result = update_checklist(
        checklist_path,
        email_input_path,
        output_folder,
        api_key,
        provider=provider,
        model_name=model_name,
        provider_base_url=provider_base_url,
    )
    print(json.dumps(result))


if __name__ == '__main__':
    main()
