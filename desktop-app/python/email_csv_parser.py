#!/usr/bin/env python3
"""
EmmaNeigh - Email CSV Parser
v5.1.6: Added boolean search with attachment filtering
Parses Outlook CSV exports and stores email metadata for checklist integration.
Supports boolean keyword search (AND, OR, NOT, quotes) and attachment filtering.
"""

import os
import sys
import json
import csv
import re
from datetime import datetime
from collections import defaultdict


def emit(msg_type, **kwargs):
    """Output JSON message to stdout for the Electron app."""
    print(json.dumps({"type": msg_type, **kwargs}), flush=True)


def parse_date(date_str):
    """Parse various date formats from Outlook exports."""
    if not date_str:
        return None

    formats = [
        "%m/%d/%Y %I:%M %p",
        "%m/%d/%Y %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%m/%d/%y %I:%M %p",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(date_str.strip(), fmt).isoformat()
        except ValueError:
            continue

    return date_str  # Return as-is if no format matches


def normalize_email(email):
    """Extract email address from various formats like 'Name <email@domain.com>'."""
    if not email:
        return ""

    match = re.search(r'<([^>]+)>', email)
    if match:
        return match.group(1).lower().strip()

    # If no angle brackets, assume it's just the email
    return email.lower().strip()


def extract_domain(email):
    """Extract domain from email address."""
    normalized = normalize_email(email)
    if '@' in normalized:
        return normalized.split('@')[1]
    return ""


def parse_outlook_csv(csv_path):
    """
    Parse an Outlook CSV export.

    Common Outlook CSV columns:
    - Subject, Body, From, To, CC, BCC
    - Date Sent, Date Received
    - Attachments

    Returns list of email dictionaries.
    """
    emails = []

    # Try different encodings (utf-8-sig first to handle BOM)
    encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']

    for encoding in encodings:
        try:
            with open(csv_path, 'r', encoding=encoding, newline='') as f:
                # Detect delimiter
                sample = f.read(2048)
                f.seek(0)

                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=',;\t')
                except csv.Error:
                    dialect = csv.excel

                reader = csv.DictReader(f, dialect=dialect)

                # Normalize column names (Outlook exports vary)
                for row in reader:
                    # Map various column names to standard fields
                    email_data = {
                        'subject': '',
                        'body': '',
                        'from': '',
                        'to': '',
                        'cc': '',
                        'date_sent': None,
                        'date_received': None,
                        'attachments': '',
                        'has_attachments': False
                    }

                    for key, value in row.items():
                        if not key:
                            continue

                        # Strip BOM and whitespace from key
                        key_lower = key.replace('\ufeff', '').lower().strip()
                        value = value.strip() if value else ''

                        if key_lower in ['subject', 'title']:
                            email_data['subject'] = value
                        elif key_lower in ['body', 'message', 'content', 'message body']:
                            email_data['body'] = value
                        elif key_lower in ['from', 'sender', 'from: (address)', 'from: (name)']:
                            if email_data['from']:
                                email_data['from'] += '; ' + value
                            else:
                                email_data['from'] = value
                        elif key_lower in ['to', 'recipient', 'to: (address)', 'to: (name)']:
                            if email_data['to']:
                                email_data['to'] += '; ' + value
                            else:
                                email_data['to'] = value
                        elif key_lower in ['cc', 'carbon copy']:
                            email_data['cc'] = value
                        elif key_lower in ['date sent', 'sent', 'send date']:
                            email_data['date_sent'] = parse_date(value)
                        elif key_lower in ['date received', 'received', 'receive date', 'date']:
                            email_data['date_received'] = parse_date(value)
                        elif key_lower == 'has attachments':
                            # Boolean column from Outlook - just TRUE/FALSE
                            email_data['has_attachments'] = bool(value and value.lower() not in ['no', 'false', '0', ''])
                        elif key_lower in ['attachments', 'attachment']:
                            # Actual attachment filename(s)
                            email_data['attachments'] = value
                            email_data['has_attachments'] = bool(value and value.lower() not in ['no', 'false', '0', ''])

                    # Extract domains for filtering
                    email_data['from_domain'] = extract_domain(email_data['from'])
                    email_data['to_domain'] = extract_domain(email_data['to'])

                    emails.append(email_data)

                break  # Successfully parsed

        except UnicodeDecodeError:
            continue
        except Exception as e:
            emit("progress", percent=0, message=f"Warning: Error with encoding {encoding}: {str(e)}")
            continue

    return emails


def parse_boolean_query(query):
    """
    Parse a boolean search query into tokens.

    Supports:
    - AND, OR, NOT operators (case insensitive)
    - Quoted phrases: "purchase agreement"
    - Attachment filter: attachment:filename or has:attachment
    - Implicit AND between terms

    Returns list of tokens with types:
    [{'type': 'term'|'phrase'|'and'|'or'|'not'|'attachment', 'value': str}, ...]
    """
    tokens = []
    query = query.strip()
    i = 0

    while i < len(query):
        # Skip whitespace
        while i < len(query) and query[i].isspace():
            i += 1
        if i >= len(query):
            break

        # Check for quoted phrase
        if query[i] == '"':
            end = query.find('"', i + 1)
            if end == -1:
                end = len(query)
            phrase = query[i+1:end].strip()
            if phrase:
                tokens.append({'type': 'phrase', 'value': phrase.lower()})
            i = end + 1
            continue

        # Extract word
        start = i
        while i < len(query) and not query[i].isspace() and query[i] != '"':
            i += 1
        word = query[start:i]

        if not word:
            continue

        word_upper = word.upper()
        word_lower = word.lower()

        # Check for operators
        if word_upper == 'AND':
            tokens.append({'type': 'and', 'value': 'AND'})
        elif word_upper == 'OR':
            tokens.append({'type': 'or', 'value': 'OR'})
        elif word_upper == 'NOT':
            tokens.append({'type': 'not', 'value': 'NOT'})
        # Check for attachment filter
        elif word_lower.startswith('attachment:'):
            attachment_query = word[11:]  # Remove "attachment:"
            tokens.append({'type': 'attachment', 'value': attachment_query.lower()})
        elif word_lower == 'has:attachment' or word_lower == 'has:attachments':
            tokens.append({'type': 'has_attachment', 'value': True})
        else:
            tokens.append({'type': 'term', 'value': word_lower})

    return tokens


def evaluate_boolean_query(email, tokens, search_fields):
    """
    Evaluate a boolean query against an email.

    Returns True if email matches the query.
    """
    if not tokens:
        return False

    # Build searchable text from specified fields
    searchable_parts = []
    for field in search_fields:
        value = email.get(field, '')
        if value:
            searchable_parts.append(value.lower())
    searchable_text = ' '.join(searchable_parts)

    # Get attachment text separately
    attachment_text = email.get('attachments', '').lower()
    has_attachments = email.get('has_attachments', False)

    # Evaluate tokens with simple boolean logic
    # Default is AND between consecutive terms
    results = []
    current_op = 'and'  # Default operator
    negate_next = False

    for token in tokens:
        if token['type'] == 'and':
            current_op = 'and'
            continue
        elif token['type'] == 'or':
            current_op = 'or'
            continue
        elif token['type'] == 'not':
            negate_next = True
            continue

        # Evaluate the term/phrase/attachment
        if token['type'] == 'term':
            match = token['value'] in searchable_text
        elif token['type'] == 'phrase':
            match = token['value'] in searchable_text
        elif token['type'] == 'attachment':
            match = token['value'] in attachment_text
        elif token['type'] == 'has_attachment':
            match = has_attachments
        else:
            match = False

        # Apply negation
        if negate_next:
            match = not match
            negate_next = False

        # Combine with previous results
        if not results:
            results.append(match)
        elif current_op == 'and':
            results[-1] = results[-1] and match
        elif current_op == 'or':
            results.append(match)

        # Reset to default AND
        current_op = 'and'

    # Final result: any OR group must be true
    return any(results) if results else False


def search_emails(emails, query, search_fields=None, attachment_only=False):
    """
    Search emails with boolean query support.

    v5.1.6: Supports boolean operators (AND, OR, NOT), quoted phrases,
    and attachment filtering.

    Args:
        emails: List of email dictionaries
        query: Search query string with optional boolean operators
               Examples:
               - "purchase agreement" (exact phrase)
               - credit AND agreement
               - loan OR credit
               - NOT draft
               - attachment:agreement.pdf
               - has:attachment
        search_fields: List of fields to search (default: subject, body, from, to, attachments)
        attachment_only: If True, only search in attachments field

    Returns list of matching emails with match info.
    """
    if not search_fields:
        search_fields = ['subject', 'body', 'from', 'to', 'attachments']

    if attachment_only:
        search_fields = ['attachments']

    # Parse the boolean query
    tokens = parse_boolean_query(query)

    if not tokens:
        return []

    results = []

    for email in emails:
        if evaluate_boolean_query(email, tokens, search_fields):
            # Determine which fields matched (for display)
            matched_fields = []
            for field in search_fields:
                value = email.get(field, '')
                if value:
                    # Check if any search term appears in this field
                    for token in tokens:
                        if token['type'] in ('term', 'phrase'):
                            if token['value'] in value.lower():
                                matched_fields.append(field)
                                break
                        elif token['type'] == 'attachment' and field == 'attachments':
                            if token['value'] in value.lower():
                                matched_fields.append('attachments')
                                break

            results.append({
                'email': email,
                'matched_fields': list(set(matched_fields))
            })

    return results


def search_attachments(emails, query):
    """
    Search specifically for attachments matching the query.

    v5.1.6: Convenience function for attachment-only search.

    Args:
        emails: List of email dictionaries
        query: Search query (can use boolean operators)

    Returns list of emails with matching attachments.
    """
    # First filter to only emails with attachments
    emails_with_attachments = [e for e in emails if e.get('has_attachments')]

    # Parse query
    tokens = parse_boolean_query(query)

    # If no special operators, treat as simple attachment search
    if all(t['type'] == 'term' for t in tokens):
        # Add implicit attachment: prefix to all terms
        tokens = [{'type': 'attachment', 'value': t['value']} for t in tokens]

    results = []
    for email in emails_with_attachments:
        attachment_text = email.get('attachments', '').lower()

        # Check if any term matches
        matches = False
        for token in tokens:
            if token['type'] in ('term', 'attachment', 'phrase'):
                if token['value'] in attachment_text:
                    matches = True
                    break

        if matches:
            results.append({
                'email': email,
                'matched_fields': ['attachments'],
                'attachments': email.get('attachments', '')
            })

    return results


def generate_summary(emails):
    """Generate summary statistics about the email dataset."""
    summary = {
        'total_emails': len(emails),
        'date_range': {'earliest': None, 'latest': None},
        'by_sender_domain': defaultdict(int),
        'by_recipient_domain': defaultdict(int),
        'with_attachments': 0,
        'unique_senders': set(),
        'unique_recipients': set()
    }

    dates = []

    for email in emails:
        # Count by domain
        if email.get('from_domain'):
            summary['by_sender_domain'][email['from_domain']] += 1
            summary['unique_senders'].add(normalize_email(email['from']))

        if email.get('to_domain'):
            summary['by_recipient_domain'][email['to_domain']] += 1
            summary['unique_recipients'].add(normalize_email(email['to']))

        if email.get('has_attachments'):
            summary['with_attachments'] += 1

        # Track dates
        date = email.get('date_sent') or email.get('date_received')
        if date:
            dates.append(date)

    # Date range
    if dates:
        dates.sort()
        summary['date_range']['earliest'] = dates[0]
        summary['date_range']['latest'] = dates[-1]

    # Convert sets to counts
    summary['unique_senders'] = len(summary['unique_senders'])
    summary['unique_recipients'] = len(summary['unique_recipients'])

    # Convert defaultdicts to regular dicts for JSON serialization
    summary['by_sender_domain'] = dict(summary['by_sender_domain'])
    summary['by_recipient_domain'] = dict(summary['by_recipient_domain'])

    return summary


def main():
    """CLI entry point."""
    if len(sys.argv) < 2:
        emit("error", message="Usage: email_csv_parser.py <config_json_path>")
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

    action = config.get('action', 'parse')

    if action == 'parse':
        # Parse CSV and return emails + summary
        csv_path = config.get('csv_path')

        if not csv_path or not os.path.isfile(csv_path):
            emit("error", message="CSV file not found")
            sys.exit(1)

        emit("progress", percent=10, message="Reading CSV file...")

        emails = parse_outlook_csv(csv_path)

        emit("progress", percent=70, message="Generating summary...")

        summary = generate_summary(emails)

        emit("progress", percent=100, message="Complete!")

        emit("result",
             success=True,
             action="parse",
             emails=emails,
             summary=summary)

    elif action == 'search':
        # Search within provided emails (v5.1.6: supports boolean operators)
        emails = config.get('emails', [])
        query = config.get('query', '')
        search_fields = config.get('search_fields')
        attachment_only = config.get('attachment_only', False)

        if not query:
            emit("error", message="No search query provided")
            sys.exit(1)

        results = search_emails(emails, query, search_fields, attachment_only)

        emit("result",
             success=True,
             action="search",
             query=query,
             results=results,
             total_matches=len(results))

    elif action == 'search_attachments':
        # v5.1.6: Search specifically for attachments
        emails = config.get('emails', [])
        query = config.get('query', '')

        if not query:
            emit("error", message="No search query provided")
            sys.exit(1)

        results = search_attachments(emails, query)

        emit("result",
             success=True,
             action="search_attachments",
             query=query,
             results=results,
             total_matches=len(results))

    else:
        emit("error", message=f"Unknown action: {action}")
        sys.exit(1)


if __name__ == "__main__":
    main()
