#!/usr/bin/env python3
"""
EmmaNeigh - Time Tracker
Analyzes email CSV and calendar (ICS/CSV) exports to generate activity summaries.
Shows time breakdown by matter/client, activity type, and timeline.
"""

import os
import sys
import json
import csv
import re
from datetime import datetime, timedelta
from collections import defaultdict


def emit(msg_type, **kwargs):
    """Output JSON message to stdout for the Electron app."""
    print(json.dumps({"type": msg_type, **kwargs}), flush=True)


def parse_date(date_str):
    """Parse various date formats."""
    if not date_str:
        return None

    formats = [
        "%m/%d/%Y %I:%M %p",
        "%m/%d/%Y %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%dT%H:%M:%S.%f",  # ISO with microseconds
        "%Y-%m-%dT%H:%M:%SZ",
        "%d/%m/%Y %H:%M",
        "%m/%d/%y %I:%M %p",
        "%Y%m%dT%H%M%SZ",  # ICS format
        "%Y%m%dT%H%M%S",   # ICS local format
    ]

    for fmt in formats:
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except ValueError:
            continue

    return None


def parse_ics_file(ics_path):
    """Parse an ICS calendar file and extract events."""
    events = []

    try:
        with open(ics_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except:
        with open(ics_path, 'r', encoding='latin-1') as f:
            content = f.read()

    # Simple ICS parsing (without external dependencies)
    event_blocks = re.findall(r'BEGIN:VEVENT(.*?)END:VEVENT', content, re.DOTALL)

    for block in event_blocks:
        event = {
            'summary': '',
            'start': None,
            'end': None,
            'duration_minutes': 0,
            'location': '',
            'description': ''
        }

        # Extract fields
        summary_match = re.search(r'SUMMARY[^:]*:(.+?)(?:\r?\n(?!\s)|\Z)', block, re.DOTALL)
        if summary_match:
            event['summary'] = summary_match.group(1).replace('\r\n ', '').replace('\n ', '').strip()

        # Start time
        dtstart_match = re.search(r'DTSTART[^:]*:(\d{8}T?\d{0,6}Z?)', block)
        if dtstart_match:
            event['start'] = parse_date(dtstart_match.group(1))

        # End time
        dtend_match = re.search(r'DTEND[^:]*:(\d{8}T?\d{0,6}Z?)', block)
        if dtend_match:
            event['end'] = parse_date(dtend_match.group(1))

        # Calculate duration
        if event['start'] and event['end']:
            delta = event['end'] - event['start']
            event['duration_minutes'] = int(delta.total_seconds() / 60)

        # Location
        location_match = re.search(r'LOCATION[^:]*:(.+?)(?:\r?\n(?!\s)|\Z)', block, re.DOTALL)
        if location_match:
            event['location'] = location_match.group(1).replace('\r\n ', '').replace('\n ', '').strip()

        if event['start']:
            events.append(event)

    return events


def parse_calendar_csv(csv_path):
    """Parse Outlook calendar CSV export."""
    events = []

    encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252']

    for encoding in encodings:
        try:
            with open(csv_path, 'r', encoding=encoding) as f:
                reader = csv.DictReader(f)

                for row in reader:
                    event = {
                        'summary': '',
                        'start': None,
                        'end': None,
                        'duration_minutes': 0,
                        'location': '',
                        'description': ''
                    }

                    for key, value in row.items():
                        if not key:
                            continue

                        key_lower = key.lower().strip()
                        value = value.strip() if value else ''

                        if key_lower in ['subject', 'title', 'summary']:
                            event['summary'] = value
                        elif key_lower in ['start date', 'start', 'begin']:
                            event['start'] = parse_date(value)
                        elif key_lower in ['end date', 'end']:
                            event['end'] = parse_date(value)
                        elif key_lower in ['location']:
                            event['location'] = value
                        elif key_lower in ['description', 'body', 'notes']:
                            event['description'] = value
                        elif key_lower in ['duration', 'length']:
                            try:
                                event['duration_minutes'] = int(float(value))
                            except:
                                pass

                    # Calculate duration if not provided
                    if event['start'] and event['end'] and not event['duration_minutes']:
                        delta = event['end'] - event['start']
                        event['duration_minutes'] = int(delta.total_seconds() / 60)

                    if event['start']:
                        events.append(event)

                break
        except:
            continue

    return events


def parse_emails_for_activity(emails):
    """Extract activity patterns from emails."""
    activity = []

    for email in emails:
        date = None
        if email.get('date_sent'):
            date = parse_date(email['date_sent'])
        elif email.get('date_received'):
            date = parse_date(email['date_received'])

        if date:
            activity.append({
                'type': 'email',
                'timestamp': date,
                'subject': email.get('subject', ''),
                'from': email.get('from', ''),
                'to': email.get('to', ''),
                'direction': 'sent' if email.get('date_sent') else 'received'
            })

    return activity


def extract_matter_from_text(text, known_matters=None):
    """
    Try to extract matter/client name from text.
    Uses simple heuristics - can be enhanced with Claude API.
    """
    if not text:
        return 'General'

    text_lower = text.lower()

    # Common patterns
    patterns = [
        r're:\s*(.+?)(?:\s*-|\s*\||$)',  # Re: Matter Name -
        r'(?:matter|project|deal|transaction):\s*(.+?)(?:\s*-|\s*\||$)',
        r'\[(.+?)\]',  # [Matter Name]
    ]

    for pattern in patterns:
        match = re.search(pattern, text_lower)
        if match:
            matter = match.group(1).strip()
            if len(matter) > 3 and len(matter) < 50:
                return matter.title()

    # Check against known matters
    if known_matters:
        for matter in known_matters:
            if matter.lower() in text_lower:
                return matter

    return 'General'


def generate_activity_summary(emails, calendar_events, target_date=None, period='day'):
    """
    Generate an activity summary for a given period.

    Args:
        emails: List of email dictionaries
        calendar_events: List of calendar event dictionaries
        target_date: Date to summarize (default: today)
        period: 'day' or 'week'

    Returns:
        Summary dictionary
    """
    if target_date is None:
        target_date = datetime.now().date()
    elif isinstance(target_date, str):
        target_date = datetime.fromisoformat(target_date).date()

    # Determine date range
    if period == 'week':
        start_of_week = target_date - timedelta(days=target_date.weekday())
        date_range = [start_of_week + timedelta(days=i) for i in range(7)]
    else:
        date_range = [target_date]

    # Parse emails into activity
    email_activity = parse_emails_for_activity(emails)

    # Filter to date range
    def in_range(dt):
        if dt is None:
            return False
        if isinstance(dt, datetime):
            dt = dt.date()
        return dt in date_range

    filtered_emails = [a for a in email_activity if in_range(a['timestamp'])]
    filtered_events = [e for e in calendar_events if in_range(e['start'])]

    # Group by matter
    by_matter = defaultdict(lambda: {
        'emails_sent': 0,
        'emails_received': 0,
        'meetings': [],
        'meeting_minutes': 0
    })

    for email in filtered_emails:
        matter = extract_matter_from_text(email['subject'])
        if email['direction'] == 'sent':
            by_matter[matter]['emails_sent'] += 1
        else:
            by_matter[matter]['emails_received'] += 1

    for event in filtered_events:
        matter = extract_matter_from_text(event['summary'])
        by_matter[matter]['meetings'].append({
            'summary': event['summary'],
            'start': event['start'].isoformat() if event['start'] else None,
            'duration': event['duration_minutes']
        })
        by_matter[matter]['meeting_minutes'] += event['duration_minutes']

    # Build timeline
    timeline = []

    for event in sorted(filtered_events, key=lambda e: e['start'] or datetime.min):
        if event['start']:
            timeline.append({
                'time': event['start'].strftime('%I:%M %p'),
                'type': 'meeting',
                'description': event['summary'],
                'duration': event['duration_minutes']
            })

    # Group emails by hour (skip emails with no valid timestamp)
    email_by_hour = defaultdict(list)
    for email in filtered_emails:
        if email.get('timestamp') is None:
            continue
        hour = email['timestamp'].strftime('%I:00 %p')
        email_by_hour[hour].append(email)

    for hour, emails_in_hour in sorted(email_by_hour.items()):
        matters = set(extract_matter_from_text(e['subject']) for e in emails_in_hour)
        timeline.append({
            'time': hour,
            'type': 'emails',
            'description': f"Email activity ({len(emails_in_hour)} emails) - {', '.join(matters)}",
            'count': len(emails_in_hour)
        })

    # Calculate totals
    total_meeting_minutes = sum(e['duration_minutes'] for e in filtered_events)
    total_emails = len(filtered_emails)

    # Estimate active hours (rough: meetings + email time)
    email_minutes = total_emails * 3  # Assume 3 min per email
    total_active_minutes = total_meeting_minutes + email_minutes
    total_active_hours = round(total_active_minutes / 60, 1)

    # Format by matter for output
    matters_summary = []
    for matter, data in sorted(by_matter.items(), key=lambda x: -x[1]['meeting_minutes']):
        matter_minutes = data['meeting_minutes'] + (data['emails_sent'] + data['emails_received']) * 3
        matters_summary.append({
            'name': matter,
            'hours': round(matter_minutes / 60, 1),
            'percent': round(matter_minutes / max(total_active_minutes, 1) * 100),
            'emails_sent': data['emails_sent'],
            'emails_received': data['emails_received'],
            'meetings': len(data['meetings']),
            'meeting_hours': round(data['meeting_minutes'] / 60, 1)
        })

    return {
        'period': period,
        'date': target_date.isoformat(),
        'date_range': {
            'start': date_range[0].isoformat(),
            'end': date_range[-1].isoformat()
        },
        'total_active_hours': total_active_hours,
        'total_meetings': len(filtered_events),
        'total_meeting_hours': round(total_meeting_minutes / 60, 1),
        'total_emails': total_emails,
        'by_matter': matters_summary,
        'timeline': sorted(timeline, key=lambda x: x['time'])
    }


def main():
    """CLI entry point."""
    if len(sys.argv) < 2:
        emit("error", message="Usage: time_tracker.py <config_json_path>")
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

    emit("progress", percent=10, message="Loading data...")

    emails = config.get('emails', [])
    calendar_path = config.get('calendar_path')
    target_date = config.get('target_date')
    period = config.get('period', 'day')

    # Parse calendar if provided
    calendar_events = []
    if calendar_path and os.path.isfile(calendar_path):
        emit("progress", percent=30, message="Parsing calendar...")

        if calendar_path.lower().endswith('.ics'):
            calendar_events = parse_ics_file(calendar_path)
        elif calendar_path.lower().endswith('.csv'):
            calendar_events = parse_calendar_csv(calendar_path)

    emit("progress", percent=60, message="Generating summary...")

    summary = generate_activity_summary(
        emails,
        calendar_events,
        target_date,
        period
    )

    emit("progress", percent=100, message="Complete!")

    emit("result",
         success=True,
         summary=summary)


if __name__ == "__main__":
    main()
