#!/usr/bin/env python3
"""Convert HPI SmartWe CSV exports to ICS calendar format."""

import csv
import os
from datetime import datetime
from pathlib import Path

def parse_german_datetime(dt_str: str) -> datetime:
    """Parse German date format DD.MM.YYYY HH:MM to datetime."""
    return datetime.strptime(dt_str, "%d.%m.%Y %H:%M")

def format_ics_datetime(dt: datetime) -> str:
    """Format datetime to ICS format (YYYYMMDDTHHMMSS)."""
    return dt.strftime("%Y%m%dT%H%M%S")

def escape_ics_text(text: str) -> str:
    """Escape special characters for ICS format."""
    if not text:
        return ""
    return text.replace("\\", "\\\\").replace(",", "\\,").replace(";", "\\;").replace("\n", "\\n")

def csv_to_ics(csv_files: list[Path], output_file: Path) -> int:
    """Convert multiple CSV files to a single ICS file."""
    events = []

    for csv_file in csv_files:
        with open(csv_file, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f, delimiter=';')
            for row in reader:
                try:
                    start = parse_german_datetime(row['Beginn'])
                    end = parse_german_datetime(row['Ende'])

                    event = {
                        'uid': row.get('GGUID', ''),
                        'summary': row.get('Betreff', row.get('Titel', '')),
                        'location': row.get('Ort', ''),
                        'dtstart': format_ics_datetime(start),
                        'dtend': format_ics_datetime(end),
                        'description': row.get('Notizen', ''),
                        'typ': row.get('Typ', ''),
                    }
                    events.append(event)
                except (ValueError, KeyError) as e:
                    print(f"Skipping invalid row in {csv_file}: {e}")
                    continue

    # Generate ICS content
    ics_lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//HPI SmartWe Export//EN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:HPI Courses",
    ]

    for event in events:
        description = event['description']
        if event['typ']:
            description = f"Typ: {event['typ']}" + (f"\\n{description}" if description else "")

        ics_lines.extend([
            "BEGIN:VEVENT",
            f"UID:{event['uid']}@hpi.de",
            f"DTSTART:{event['dtstart']}",
            f"DTEND:{event['dtend']}",
            f"SUMMARY:{escape_ics_text(event['summary'])}",
            f"LOCATION:{escape_ics_text(event['location'])}",
        ])

        if description:
            ics_lines.append(f"DESCRIPTION:{escape_ics_text(description)}")

        ics_lines.append("END:VEVENT")

    ics_lines.append("END:VCALENDAR")

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("\r\n".join(ics_lines))

    return len(events)

def main():
    dat_dir = Path(__file__).parent / "dat"
    csv_files = list(dat_dir.glob("Termine*.csv"))

    if not csv_files:
        print("No CSV files found in dat/")
        return

    output_file = dat_dir / "hpi_courses.ics"
    count = csv_to_ics(csv_files, output_file)

    print(f"Converted {count} events from {len(csv_files)} CSV files")
    print(f"Output: {output_file}")

if __name__ == "__main__":
    main()
