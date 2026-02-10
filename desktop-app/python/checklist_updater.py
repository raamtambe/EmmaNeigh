#!/usr/bin/env python3
"""
Checklist Updater - Update transaction checklist based on email activity

This module parses a transaction checklist (Word document) and an email CSV export,
then updates the checklist status column based on detected email activity.

Usage:
    python checklist_updater.py <checklist_path> <email_csv_path> <output_folder>
"""

import sys
import os
import json
import re
from datetime import datetime
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Try to import anthropic for LLM matching
try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False

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


def match_documents_with_llm(checklist_items, emails, api_key):
    """
    Use Claude API to match emails to documents and infer status.

    Args:
        checklist_items: List of document names from checklist
        emails: List of email dicts with subject, body, from, date
        api_key: Claude API key

    Returns:
        Dict mapping document_name to {status, matching_emails, confidence}
    """
    if not HAS_ANTHROPIC or not api_key:
        return None

    try:
        client = anthropic.Anthropic(api_key=api_key)

        # Prepare email context (limit and truncate for token efficiency)
        email_context = []
        for i, email in enumerate(emails[:100]):
            email_context.append({
                "index": i,
                "from": email.get("from", "")[:100],
                "subject": email.get("subject", "")[:200],
                "body_preview": email.get("body", "")[:200],
                "date": email.get("date_received") or email.get("date_sent") or ""
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

Only include documents that have clear email activity. ONLY return the JSON object, no other text."""

        response = client.messages.create(
            model="claude-3-5-sonnet-20241022",
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
        print(f"LLM matching failed: {e}", file=sys.stderr)
        return None


def update_checklist(checklist_path, email_csv_path, output_folder, api_key=None):
    """
    Main function to update checklist based on email activity.

    Args:
        checklist_path: Path to Word document with checklist table
        email_csv_path: Path to Outlook email CSV export
        output_folder: Folder to save updated checklist

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

        # Try LLM matching if API key is available
        llm_matches = None
        if api_key and doc_names:
            llm_matches = match_documents_with_llm(doc_names, emails, api_key)

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

            # Try LLM match first, then fall back to regex matching
            new_status = None
            matching_emails = []

            if llm_matches and doc_name in llm_matches:
                llm_result = llm_matches[doc_name]
                new_status = llm_result.get('status')
                # Get matching email subjects from indices
                email_indices = llm_result.get('matching_email_indices', [])
                for idx in email_indices[:3]:
                    if idx < len(emails):
                        matching_emails.append(emails[idx].get('subject', 'No subject'))

            # Fall back to regex if no LLM match
            if not new_status:
                new_status, priority, matching_emails = detect_document_status(doc_name, emails)

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

        return result

    except Exception as e:
        result['error'] = str(e)
        return result


def main():
    if len(sys.argv) < 4:
        print(json.dumps({
            'success': False,
            'error': 'Usage: checklist_updater.py <checklist_path> <email_csv_path> <output_folder> [api_key]'
        }))
        sys.exit(1)

    checklist_path = sys.argv[1]
    email_csv_path = sys.argv[2]
    output_folder = sys.argv[3]

    # Get API key if provided
    api_key = None
    if len(sys.argv) > 4:
        api_key = sys.argv[4]

    # Also check environment variable
    if not api_key:
        api_key = os.environ.get('ANTHROPIC_API_KEY')

    result = update_checklist(checklist_path, email_csv_path, output_folder, api_key)
    print(json.dumps(result))


if __name__ == '__main__':
    main()
