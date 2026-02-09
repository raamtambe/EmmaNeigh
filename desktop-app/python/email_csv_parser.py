#!/usr/bin/env python3
"""
EmmaNeigh - Email CSV Parser
Parses Outlook CSV exports and stores email metadata for checklist integration.
Supports keyword search and can be used with Claude API for semantic queries.
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

    # Try different encodings
    encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252']

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

                        key_lower = key.lower().strip()
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
                        elif key_lower in ['attachments', 'attachment', 'has attachments']:
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


def search_emails(emails, query, search_fields=None):
    """
    Search emails by keyword.

    Args:
        emails: List of email dictionaries
        query: Search query string
        search_fields: List of fields to search (default: subject, body, from, to)

    Returns list of matching emails with match info.
    """
    if not search_fields:
        search_fields = ['subject', 'body', 'from', 'to']

    query_lower = query.lower()
    results = []

    for email in emails:
        matches = []
        for field in search_fields:
            value = email.get(field, '')
            if value and query_lower in value.lower():
                matches.append(field)

        if matches:
            results.append({
                'email': email,
                'matched_fields': matches
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
        # Search within provided emails
        emails = config.get('emails', [])
        query = config.get('query', '')
        search_fields = config.get('search_fields')

        if not query:
            emit("error", message="No search query provided")
            sys.exit(1)

        results = search_emails(emails, query, search_fields)

        emit("result",
             success=True,
             action="search",
             query=query,
             results=results,
             total_matches=len(results))

    else:
        emit("error", message=f"Unknown action: {action}")
        sys.exit(1)


if __name__ == "__main__":
    main()
