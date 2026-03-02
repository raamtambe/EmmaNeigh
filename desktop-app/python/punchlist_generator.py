#!/usr/bin/env python3
"""
Punchlist Generator - Generate daily punchlist from transaction checklist

This module parses a transaction checklist (Word document) and generates
a formatted punchlist showing only open/pending items organized by status.

A punchlist is a daily working document that shows:
- What items are still pending/open
- What needs attention today
- Grouped by status (pending draft, with counsel, awaiting signature, etc.)

Usage:
    python punchlist_generator.py <checklist_path> <output_folder> [status_filters_json] <api_key> [provider] [model] [harvey_base_url]
"""

import sys
import os
import json
import re
import urllib.error
import urllib.request
import uuid
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE

DEFAULT_CLAUDE_MODEL = "claude-sonnet-4-20250514"
DEFAULT_OPENAI_MODEL = "gpt-4.1-mini"
DEFAULT_HARVEY_BASE_URL = "https://api.harvey.ai"


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

    return json.loads(text)


def canonical_doc_key(value):
    """Normalize document names for key matching."""
    text = re.sub(r'\s+', ' ', str(value or '')).strip().lower()
    return text


def normalize_provider(provider):
    value = str(provider or "").strip().lower()
    if value == "claude":
        return "anthropic"
    if value in ("anthropic", "openai", "harvey"):
        return value
    return "anthropic"


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


def parse_error_detail(raw_text):
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
            return status_code, raw_text
    except urllib.error.HTTPError as e:
        raw_text = e.read().decode("utf-8", errors="replace")
        return e.code, raw_text


def perform_json_request(url, headers, payload, timeout=90):
    request_headers = dict(headers or {})
    request_headers["Content-Type"] = "application/json"
    body_bytes = json.dumps(payload).encode("utf-8")
    status_code, raw_text = perform_http_request(url, request_headers, body_bytes, timeout=timeout)
    try:
        parsed = json.loads(raw_text) if raw_text else {}
    except Exception:
        parsed = {}
    return status_code, parsed, raw_text


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
        payload = {
            "model": candidate,
            "max_tokens": 2048,
            "temperature": 0,
            "messages": [{"role": "user", "content": prompt}],
        }
        status_code, parsed, raw_text = perform_json_request(
            "https://api.anthropic.com/v1/messages",
            {
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
            },
            payload,
        )

        if status_code != 200:
            detail = parse_error_detail(raw_text)
            last_detail = detail
            if is_likely_model_error(status_code, detail):
                continue
            raise RuntimeError(f"Anthropic error ({status_code}): {detail}")

        content = parsed.get("content") if isinstance(parsed, dict) else None
        if isinstance(content, list) and content:
            first = content[0]
            if isinstance(first, dict) and first.get("text"):
                return str(first.get("text")).strip(), candidate
        raise RuntimeError("Anthropic response was missing message content.")

    raise RuntimeError(
        f"No supported Claude model is available for this API key. Last error: {last_detail}"
    )


def call_openai_prompt(prompt, api_key, model_name):
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
        payload = {
            "model": candidate,
            "max_tokens": 2048,
            "temperature": 0,
            "messages": [{"role": "user", "content": prompt}],
        }
        status_code, parsed, raw_text = perform_json_request(
            "https://api.openai.com/v1/chat/completions",
            {"Authorization": f"Bearer {api_key}"},
            payload,
        )

        if status_code != 200:
            detail = parse_error_detail(raw_text)
            last_detail = detail
            if is_likely_model_error(status_code, detail):
                continue
            raise RuntimeError(f"OpenAI error ({status_code}): {detail}")

        choices = parsed.get("choices") if isinstance(parsed, dict) else None
        message = choices[0].get("message") if isinstance(choices, list) and choices else {}
        content = extract_openai_text(message.get("content") if isinstance(message, dict) else "")
        if content:
            return content, candidate
        raise RuntimeError("OpenAI response was missing message content.")

    raise RuntimeError(
        f"No supported OpenAI model is available for this API key. Last error: {last_detail}"
    )


def call_harvey_prompt(prompt, api_key, max_tokens, harvey_base_url):
    base_url = str(
        harvey_base_url or os.environ.get("HARVEY_BASE_URL") or DEFAULT_HARVEY_BASE_URL
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

    status_code, raw_text = perform_http_request(
        f"{base_url}/api/v2/completion?include_citations=false",
        {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": f"multipart/form-data; boundary={boundary}",
        },
        body,
    )
    if status_code != 200:
        detail = parse_error_detail(raw_text)
        raise RuntimeError(f"Harvey error ({status_code}): {detail}")

    try:
        parsed = json.loads(raw_text) if raw_text else {}
    except Exception:
        parsed = {}

    text = ""
    if isinstance(parsed, dict):
        text = parsed.get("response") or parsed.get("text") or ""
    if not text:
        raise RuntimeError("Harvey response was missing message content.")

    return str(text).strip(), "harvey-assist"


def call_provider_prompt_json(prompt, api_key, provider, model_name=None, harvey_base_url=None):
    provider_name = normalize_provider(provider)
    if provider_name == "openai":
        response_text, _ = call_openai_prompt(prompt, api_key, model_name)
    elif provider_name == "harvey":
        response_text, _ = call_harvey_prompt(prompt, api_key, 2048, harvey_base_url)
    else:
        response_text, _ = call_anthropic_prompt(prompt, api_key, model_name)

    parsed = parse_json_response_text(response_text)
    if not isinstance(parsed, dict):
        raise ValueError("LLM response was not a JSON object.")
    return parsed


# Status categories for grouping (order determines display order)
STATUS_CATEGORIES = {
    'pending': {
        'title': 'PENDING DRAFTING',
        'patterns': [
            r'^pending',
            r'^to\s*be\s*drafted',
            r'^needs?\s*draft',
            r'^not\s*started',
            r'^tbd',
            r'^$',  # Empty status = pending
        ],
        'description': 'Documents that need initial drafts'
    },
    'review': {
        'title': 'UNDER REVIEW / WITH OPPOSING COUNSEL',
        'patterns': [
            r'with\s*(?:opposing\s*)?counsel',
            r'under\s*review',
            r'(?:awaiting|pending)\s*(?:comments?|review|feedback)',
            r'sent\s*(?:to|for)',
            r'circulated',
            r'draft\s*(?:circulated|sent)',
        ],
        'description': 'Documents out for review or with counterparty'
    },
    'signature': {
        'title': 'AWAITING SIGNATURE',
        'patterns': [
            r'execution\s*version',
            r'(?:ready|awaiting|pending)\s*(?:for\s*)?(?:signature|execution)',
            r'agreed\s*form',
            r'final\s*(?:form|version)',
            r'for\s*signature',
        ],
        'description': 'Documents ready for signature'
    },
    'executed': {
        'title': 'EXECUTED / COMPLETE',
        'patterns': [
            r'^executed',
            r'^signed',
            r'^complete[d]?',
            r'^done',
            r'fully\s*executed',
        ],
        'description': 'Completed documents (typically excluded from punchlist)'
    },
}

# Column name patterns
DOCUMENT_COLUMN_PATTERNS = [
    r'^document',
    r'^doc\s*name',
    r'^agreement',
    r'^instrument',
    r'^deliverable',
    r'^item',
    r'^description',
    r'^closing\s*document',
    r'^#',  # Sometimes just numbered
]

STATUS_COLUMN_PATTERNS = [
    r'^status',
    r'^state',
    r'^progress',
    r'^current\s*status',
]

PARTY_COLUMN_PATTERNS = [
    r'^part(?:y|ies)',
    r'^signator',
    r'^signer',
    r'^responsible',
    r'^who',
]

NOTES_COLUMN_PATTERNS = [
    r'^notes?',
    r'^comments?',
    r'^remarks?',
]


def find_column_index(headers, patterns):
    """Find column index matching any of the patterns."""
    for i, header in enumerate(headers):
        header_lower = header.lower().strip()
        for pattern in patterns:
            if re.match(pattern, header_lower):
                return i
    return -1


def categorize_status(status_text):
    """
    Determine which category a status belongs to.

    Returns category key (pending, review, signature, executed) or 'pending' as default
    """
    if not status_text:
        return 'pending'

    status_lower = status_text.lower().strip()

    for category, config in STATUS_CATEGORIES.items():
        for pattern in config['patterns']:
            if re.search(pattern, status_lower):
                return category

    # Default to pending if no match
    return 'pending'


def categorize_items_with_llm(
    items,
    api_key,
    provider="anthropic",
    model_name=None,
    harvey_base_url=None,
):
    """
    Use the selected LLM provider to categorize all items at once.

    Args:
        items: List of dicts with 'document_name' and 'status' keys
        api_key: Provider API key
        provider: LLM provider (anthropic/openai/harvey)
        model_name: Optional model override
        harvey_base_url: Optional Harvey API base URL

    Returns:
        Dict mapping document_name to category
    """
    if not api_key:
        raise ValueError("API key is required for punchlist LLM categorization.")

    # Prepare items for the prompt
    items_for_prompt = [
        {"doc": item.get('document_name', ''), "status": item.get('status', '')}
        for item in items[:100]  # Limit to 100 items
    ]

    prompt = f"""You are categorizing legal transaction documents by their current status.

For each document, classify its status into one of these categories:
- "pending": Needs drafting, not started, to be drafted, TBD
- "review": Under review, with counsel, sent to counterparty, awaiting comments, circulated
- "signature": Execution version, agreed form, ready for signature, final form
- "executed": Fully executed, signed, complete, done

Documents to categorize:
{json.dumps(items_for_prompt, indent=2)}

Return a JSON object mapping each document name to its category.
Include EVERY document from the input exactly once:
{{
    "Document Name 1": "pending",
    "Document Name 2": "signature",
    ...
}}

ONLY return the JSON object, no other text."""

    parsed = call_provider_prompt_json(
        prompt,
        api_key,
        provider,
        model_name=model_name,
        harvey_base_url=harvey_base_url,
    )

    normalized = {}
    for doc_name, category in parsed.items():
        if isinstance(category, str):
            normalized[canonical_doc_key(doc_name)] = category.strip().lower()
    return normalized


def parse_checklist_for_punchlist(doc):
    """
    Parse Word document checklist table.

    Returns dict with columns and rows data
    """
    if not doc.tables:
        return None

    table = doc.tables[0]

    # Extract all rows
    all_rows = []
    for row in table.rows:
        row_data = [cell.text.strip() for cell in row.cells]
        all_rows.append(row_data)

    if not all_rows:
        return None

    headers = all_rows[0]
    rows = all_rows[1:]

    # Find relevant columns
    doc_col = find_column_index(headers, DOCUMENT_COLUMN_PATTERNS)
    status_col = find_column_index(headers, STATUS_COLUMN_PATTERNS)
    party_col = find_column_index(headers, PARTY_COLUMN_PATTERNS)
    notes_col = find_column_index(headers, NOTES_COLUMN_PATTERNS)

    if doc_col == -1:
        # Try to use first non-empty column as document column
        for i, h in enumerate(headers):
            if h.strip():
                doc_col = i
                break

    return {
        'headers': headers,
        'rows': rows,
        'doc_col': doc_col,
        'status_col': status_col,
        'party_col': party_col,
        'notes_col': notes_col
    }


def extract_transaction_name(checklist_path, doc):
    """
    Try to extract transaction name from document.

    Looks in:
    1. Document title/heading
    2. Filename
    """
    # Try to find title in document
    for para in doc.paragraphs[:5]:  # Check first 5 paragraphs
        text = para.text.strip()
        if text and len(text) > 5 and len(text) < 100:
            # Skip if it looks like a column header
            if not any(h in text.lower() for h in ['document', 'status', 'party', 'item']):
                return text

    # Fall back to filename
    base_name = os.path.splitext(os.path.basename(checklist_path))[0]
    # Clean up common suffixes
    for suffix in ['_checklist', '_closing', '_documents', 'checklist', 'closing']:
        base_name = re.sub(f'{suffix}$', '', base_name, flags=re.IGNORECASE)

    return base_name.replace('_', ' ').replace('-', ' ').strip()


def generate_punchlist(
    checklist_path,
    output_folder,
    status_filters=None,
    api_key=None,
    provider="anthropic",
    model_name=None,
    harvey_base_url=None,
):
    """
    Generate punchlist document from checklist.

    Args:
        checklist_path: Path to Word document with checklist
        output_folder: Folder to save punchlist
        status_filters: List of status categories to include ['pending', 'review', 'signature']
                       If None, includes all except 'executed'
        api_key: Provider API key for LLM-based categorization
        provider: LLM provider (anthropic/openai/harvey)
        model_name: Optional provider model override
        harvey_base_url: Optional Harvey API base URL

    Returns:
        dict with: success, output_path, item_count, categories
    """
    result = {
        'success': False,
        'output_path': None,
        'item_count': 0,
        'categories': {},
        'error': None
    }

    # Default: include all except executed
    if status_filters is None:
        status_filters = ['pending', 'review', 'signature']

    try:
        if not api_key:
            result['error'] = 'LLM API key is required to generate punchlist.'
            return result

        # Open checklist
        doc = Document(checklist_path)

        # Parse table
        parsed = parse_checklist_for_punchlist(doc)
        if not parsed:
            result['error'] = 'No table found in checklist document'
            return result

        # Get transaction name
        transaction_name = extract_transaction_name(checklist_path, doc)

        # Collect all items first
        all_items = []
        for row in parsed['rows']:
            # Get document name
            doc_name = ''
            if parsed['doc_col'] >= 0 and parsed['doc_col'] < len(row):
                doc_name = row[parsed['doc_col']]

            if not doc_name.strip():
                continue

            # Get status
            status = ''
            if parsed['status_col'] >= 0 and parsed['status_col'] < len(row):
                status = row[parsed['status_col']]

            # Get party/responsible
            party = ''
            if parsed['party_col'] >= 0 and parsed['party_col'] < len(row):
                party = row[parsed['party_col']]

            # Get notes
            notes = ''
            if parsed['notes_col'] >= 0 and parsed['notes_col'] < len(row):
                notes = row[parsed['notes_col']]

            all_items.append({
                'document_name': doc_name,
                'status': status,
                'party': party,
                'notes': notes
            })

        if not all_items:
            result['error'] = 'No checklist items found to categorize.'
            return result

        # LLM categorization is mandatory for this workflow.
        try:
            llm_categories = categorize_items_with_llm(
                all_items,
                api_key,
                provider=provider,
                model_name=model_name,
                harvey_base_url=harvey_base_url,
            )
        except Exception as e:
            result['error'] = f'LLM punchlist categorization failed: {e}'
            return result

        # Categorize all items
        categorized_items = {cat: [] for cat in STATUS_CATEGORIES.keys()}
        missing_docs = []
        invalid_docs = []

        for item in all_items:
            doc_name = item['document_name']
            category = llm_categories.get(canonical_doc_key(doc_name))
            if not category:
                missing_docs.append(doc_name)
                continue
            if category not in STATUS_CATEGORIES:
                invalid_docs.append({'document': doc_name, 'category': category})
                continue

            categorized_items[category].append({
                'document': doc_name,
                'status': item['status'],
                'party': item['party'],
                'notes': item['notes']
            })

        if missing_docs:
            preview = ', '.join(missing_docs[:10])
            result['error'] = (
                f'LLM did not categorize all checklist items ({len(missing_docs)} missing). '
                f'Examples: {preview}'
            )
            return result

        if invalid_docs:
            preview = ', '.join([f"{x['document']}={x['category']}" for x in invalid_docs[:10]])
            result['error'] = f'LLM returned invalid category labels: {preview}'
            return result

        # Create punchlist document
        punchlist_doc = Document()

        # Set narrow margins
        for section in punchlist_doc.sections:
            section.left_margin = Inches(0.75)
            section.right_margin = Inches(0.75)
            section.top_margin = Inches(0.75)
            section.bottom_margin = Inches(0.75)

        # Title
        title = punchlist_doc.add_paragraph()
        title_run = title.add_run(f'DAILY PUNCHLIST')
        title_run.bold = True
        title_run.font.size = Pt(16)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Transaction name
        if transaction_name:
            txn_para = punchlist_doc.add_paragraph()
            txn_run = txn_para.add_run(transaction_name)
            txn_run.font.size = Pt(14)
            txn_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Date
        date_para = punchlist_doc.add_paragraph()
        date_run = date_para.add_run(f'Date: {datetime.now().strftime("%B %d, %Y")}')
        date_run.font.size = Pt(11)
        date_run.font.color.rgb = RGBColor(100, 100, 100)
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        punchlist_doc.add_paragraph()  # Spacer

        # Count total open items
        total_open = sum(
            len(items) for cat, items in categorized_items.items()
            if cat in status_filters
        )

        # Summary line
        summary = punchlist_doc.add_paragraph()
        summary_run = summary.add_run(f'OPEN ITEMS: {total_open}')
        summary_run.bold = True
        summary_run.font.size = Pt(12)

        punchlist_doc.add_paragraph()  # Spacer

        # Add each category
        category_counts = {}
        for category in ['pending', 'review', 'signature', 'executed']:
            if category not in status_filters:
                continue

            items = categorized_items[category]
            if not items:
                continue

            config = STATUS_CATEGORIES[category]
            category_counts[category] = len(items)

            # Category header
            cat_header = punchlist_doc.add_paragraph()
            cat_run = cat_header.add_run(f'{config["title"]} ({len(items)})')
            cat_run.bold = True
            cat_run.font.size = Pt(11)

            # Items
            for item in items:
                item_para = punchlist_doc.add_paragraph()
                item_para.paragraph_format.left_indent = Inches(0.25)

                # Checkbox character
                checkbox = item_para.add_run('☐ ')
                checkbox.font.size = Pt(11)

                # Document name
                doc_run = item_para.add_run(item['document'])
                doc_run.font.size = Pt(11)

                # Status/notes in lighter text if available
                extra_info = []
                if item['status'] and item['status'].lower() not in ['pending', 'tbd', '']:
                    extra_info.append(item['status'])
                if item['party']:
                    extra_info.append(f"({item['party']})")

                if extra_info:
                    info_run = item_para.add_run(f" - {', '.join(extra_info)}")
                    info_run.font.size = Pt(10)
                    info_run.font.color.rgb = RGBColor(100, 100, 100)

            punchlist_doc.add_paragraph()  # Spacer between categories

        # Footer
        footer = punchlist_doc.add_paragraph()
        footer_run = footer.add_run(f'Generated by EmmaNeigh on {datetime.now().strftime("%Y-%m-%d %H:%M")}')
        footer_run.font.size = Pt(9)
        footer_run.font.color.rgb = RGBColor(150, 150, 150)
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Save punchlist
        os.makedirs(output_folder, exist_ok=True)

        timestamp = datetime.now().strftime('%Y%m%d')
        output_filename = f'Punchlist_{timestamp}.docx'
        output_path = os.path.join(output_folder, output_filename)

        punchlist_doc.save(output_path)

        result['success'] = True
        result['output_path'] = output_path
        result['item_count'] = total_open
        result['categories'] = category_counts

        return result

    except Exception as e:
        result['error'] = str(e)
        return result


def main():
    if len(sys.argv) < 3:
        print(json.dumps({
            'success': False,
            'error': 'Usage: punchlist_generator.py <checklist_path> <output_folder> [status_filters_json] <api_key> [provider] [model] [harvey_base_url]'
        }))
        sys.exit(1)

    checklist_path = sys.argv[1]
    output_folder = sys.argv[2]

    # Parse status filters if provided
    status_filters = None
    if len(sys.argv) > 3:
        try:
            status_filters = json.loads(sys.argv[3])
        except json.JSONDecodeError:
            pass

    # API key is required for LLM-driven punchlist categorization.
    api_key = sys.argv[4] if len(sys.argv) > 4 else (
        os.environ.get('ANTHROPIC_API_KEY')
        or os.environ.get('OPENAI_API_KEY')
        or os.environ.get('HARVEY_API_KEY')
    )
    provider = normalize_provider(sys.argv[5] if len(sys.argv) > 5 else os.environ.get('AI_PROVIDER', 'anthropic'))
    model_name = sys.argv[6] if len(sys.argv) > 6 else (
        os.environ.get('OPENAI_MODEL') if provider == 'openai' else os.environ.get('CLAUDE_MODEL')
    )
    harvey_base_url = sys.argv[7] if len(sys.argv) > 7 else os.environ.get('HARVEY_BASE_URL', DEFAULT_HARVEY_BASE_URL)

    result = generate_punchlist(
        checklist_path,
        output_folder,
        status_filters,
        api_key,
        provider=provider,
        model_name=model_name,
        harvey_base_url=harvey_base_url,
    )
    print(json.dumps(result))


if __name__ == '__main__':
    main()
