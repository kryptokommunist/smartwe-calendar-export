#!/usr/bin/env python3
"""
Convert SmartWe CSV exports to iCal format.

Usage:
    python csv_to_ical.py dat/*.csv -o calendar.ics
"""

import argparse
import csv
from datetime import datetime
from pathlib import Path

from icalendar import Calendar, Event


def parse_datetime(date_str: str) -> datetime:
    """Parse German date format DD.MM.YYYY HH:MM"""
    return datetime.strptime(date_str.strip(), "%d.%m.%Y %H:%M")


def csv_to_events(csv_path: Path) -> list:
    """Parse a SmartWe CSV export file."""
    events = []

    with open(csv_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f, delimiter=';')

        for row in reader:
            try:
                event = {
                    'summary': row.get('Betreff', '').strip(),
                    'start': parse_datetime(row.get('Beginn', '')),
                    'end': parse_datetime(row.get('Ende', '')),
                    'location': row.get('Ort', '').strip(),
                    'uid': row.get('GGUID', '').strip(),
                    'notes': row.get('Notizen', '').strip(),
                    'category': row.get('Kategorie', '').strip(),
                }

                if event['summary'] and event['start']:
                    events.append(event)
            except Exception as e:
                print(f"Warning: Could not parse row: {e}")
                continue

    return events


def create_ical(events: list) -> Calendar:
    """Create iCal calendar from events."""
    cal = Calendar()
    cal.add('prodid', '-//SmartWe Calendar Export//hpi.de//')
    cal.add('version', '2.0')
    cal.add('calscale', 'GREGORIAN')
    cal.add('method', 'PUBLISH')
    cal.add('x-wr-calname', 'HPI SmartWe Calendar')

    for e in events:
        ev = Event()
        ev.add('summary', e['summary'])
        ev.add('dtstart', e['start'])
        ev.add('dtend', e['end'])
        ev.add('dtstamp', datetime.now())

        if e['location']:
            ev.add('location', e['location'])

        if e['notes']:
            ev.add('description', e['notes'])

        if e['uid']:
            ev.add('uid', f"{e['uid']}@smartwe.hpi.de")
        else:
            ev.add('uid', f"{abs(hash(str(e)))}@smartwe.hpi.de")

        cal.add_component(ev)

    return cal


def main():
    parser = argparse.ArgumentParser(description='Convert SmartWe CSV to iCal')
    parser.add_argument('csv_files', nargs='+', type=Path, help='CSV files to convert')
    parser.add_argument('-o', '--output', type=Path, default=Path('calendar_export.ics'),
                        help='Output iCal file')
    args = parser.parse_args()

    all_events = []

    for csv_file in args.csv_files:
        if csv_file.exists():
            print(f"📄 Reading {csv_file.name}...")
            events = csv_to_events(csv_file)
            print(f"   Found {len(events)} events")
            all_events.extend(events)
        else:
            print(f"⚠️  File not found: {csv_file}")

    print(f"\n📊 Total events: {len(all_events)}")

    cal = create_ical(all_events)

    with open(args.output, 'wb') as f:
        f.write(cal.to_ical())

    print(f"✅ Saved to {args.output}")


if __name__ == '__main__':
    main()
