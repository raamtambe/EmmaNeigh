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
    python punchlist_generator.py <checklist_path> <output_folder> [status_filters_json]
"""

import sys
import os
import json
import re
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE

# Try to import anthropic for LLM categorization
try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False

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


def categorize_items_with_llm(items, api_key):
    """
    Use Claude API to categorize all items at once.

    Args:
        items: List of dicts with 'document_name' and 'status' keys
        api_key: Claude API key

    Returns:
        Dict mapping document_name to category
    """
    if not HAS_ANTHROPIC or not api_key:
        return None

    try:
        client = anthropic.Anthropic(api_key=api_key)

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

Return a JSON object mapping each document name to its category:
{{
    "Document Name 1": "pending",
    "Document Name 2": "signature",
    ...
}}

ONLY return the JSON object, no other text."""

        response = client.messages.create(
            model=os.environ.get("CLAUDE_MODEL", "claude-sonnet-4-6"),
            max_tokens=2048,
            messages=[{"role": "user", "content": prompt}]
        )

        response_text = response.content[0].text.strip()

        # Handle markdown code blocks
        if response_text.startswith("```"):
            lines = response_text.split("\n")
            json_lines = []
            in_json = False
            for line in lines:
                if line.startswith("```json") or line.startswith("```"):
                    in_json = not in_json if not line.startswith("```json") else True
                    continue
                if in_json:
                    json_lines.append(line)
            response_text = "\n".join(json_lines)

        return json.loads(response_text)

    except Exception as e:
        print(f"LLM categorization failed: {e}", file=sys.stderr)
        return None


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


def generate_punchlist(checklist_path, output_folder, status_filters=None, api_key=None):
    """
    Generate punchlist document from checklist.

    Args:
        checklist_path: Path to Word document with checklist
        output_folder: Folder to save punchlist
        status_filters: List of status categories to include ['pending', 'review', 'signature']
                       If None, includes all except 'executed'
        api_key: Optional Claude API key for LLM-based categorization

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

        # Try LLM categorization if API key is available
        llm_categories = None
        if api_key and all_items:
            llm_categories = categorize_items_with_llm(all_items, api_key)

        # Categorize all items
        categorized_items = {cat: [] for cat in STATUS_CATEGORIES.keys()}

        for item in all_items:
            doc_name = item['document_name']
            status = item['status']

            # Use LLM category if available, otherwise fall back to regex
            if llm_categories and doc_name in llm_categories:
                category = llm_categories[doc_name]
                # Validate category
                if category not in STATUS_CATEGORIES:
                    category = categorize_status(status)
            else:
                category = categorize_status(status)

            categorized_items[category].append({
                'document': doc_name,
                'status': status,
                'party': item['party'],
                'notes': item['notes']
            })

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
            'error': 'Usage: punchlist_generator.py <checklist_path> <output_folder> [status_filters_json] [api_key]'
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

    # Get API key if provided
    api_key = None
    if len(sys.argv) > 4:
        api_key = sys.argv[4]

    # Also check environment variable
    if not api_key:
        api_key = os.environ.get('ANTHROPIC_API_KEY')

    result = generate_punchlist(checklist_path, output_folder, status_filters, api_key)
    print(json.dumps(result))


if __name__ == '__main__':
    main()
