#!/usr/bin/env python3
"""
SmartWe Calendar & Course Export Script

Extracts calendar entries from HPI SmartWe portal and exports them as iCal format.
Loops through all enrolled courses and extracts their Termine (appointments).
"""

import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict

from icalendar import Calendar, Event
from playwright.sync_api import sync_playwright, Page

PORTAL_URL = "https://sv-portal.hpi.de/SmartWe/"
OUTPUT_FILE = Path(__file__).parent / "calendar_export.ics"
CREDENTIALS_FILE = Path(__file__).parent / ".credentials"


def load_credentials(page_match: str = "login.microsoftonline") -> Optional[Dict[str, str]]:
    """Load credentials from .credentials file matching the page URL."""
    if CREDENTIALS_FILE.exists():
        try:
            content = CREDENTIALS_FILE.read_text().strip()
            for block in content.split('\n\n'):
                creds = {}
                for line in block.strip().split('\n'):
                    if ':' in line:
                        key, val = line.split(':', 1)
                        creds[key.strip()] = val.strip()
                if 'page' in creds and page_match in creds['page']:
                    if 'user' in creds:
                        creds['username'] = creds['user']
                    if 'pw' in creds:
                        creds['password'] = creds['pw']
                    return creds
            return None
        except Exception as e:
            print(f"   Warning: Could not load credentials: {e}")
    return None


def do_login(page: Page) -> bool:
    """Handle complete login flow: MS -> ADFS -> SmartWe."""
    for attempt in range(60):
        time.sleep(2)
        url = page.url

        try:
            body = page.inner_text("body", timeout=3000)
        except:
            continue

        # Success - app loaded
        if "sv-portal.hpi.de/SmartWe" in url:
            if any(kw in body for kw in ["Kalender", "Termine", "Apps"]) and "Anmelden" not in body:
                return True

            # SmartWe login form
            if "Anmelden" in body and "nicht korrekt" not in body:
                creds = load_credentials("sv-portal")
                if creds:
                    try:
                        page.locator('input[name="username"]').first.fill(creds['username'])
                        page.locator('input[name="password"]').first.fill(creds['password'])
                        try:
                            page.get_by_label("Automatisch anmelden").check()
                        except:
                            pass
                        page.get_by_role("button", name="Anmelden").click()
                        continue
                    except:
                        pass

        # Microsoft login
        if "login.microsoftonline" in url:
            if "Stay signed in" in body:
                page.click("#idSIButton9")
                continue
            if page.locator("#i0116").is_visible(timeout=1000):
                creds = load_credentials("login.microsoftonline")
                if creds:
                    page.fill("#i0116", creds['username'])
                    page.click("#idSIButton9")
                    continue

        # ADFS login
        if "adfs.hpi" in url:
            if page.locator('input[type="password"]').is_visible(timeout=1000):
                creds = load_credentials("adfs.hpi")
                if creds:
                    page.locator('input[type="password"]').fill(creds['password'])
                    page.get_by_role("button", name="Sign in").click()
                    continue

    return False


def extract_events_from_course(page: Page, course_name: str) -> List[Dict]:
    """Extract events from current course's Termine detail page."""
    events = []

    try:
        body = page.inner_text("body", timeout=5000)

        # Pattern: Title\nDD.MM.YYYY, HH:MM\nDD.MM.YYYY, HH:MM
        event_pattern = r'([^\n]{3,100})\n(\d{2})\.(\d{2})\.(\d{4}),\s*(\d{1,2}:\d{2})\n(\d{2})\.(\d{2})\.(\d{4}),\s*(\d{1,2}:\d{2})'
        matches = re.findall(event_pattern, body)

        for m in matches:
            title = m[0].strip()
            # Skip UI elements
            if any(skip in title.lower() for skip in ['betreff', 'beginn', 'ende', 'gesamt', 'daten geladen']):
                continue

            events.append({
                'title': title,
                'start_date': f"{m[3]}-{m[2]}-{m[1]}",
                'start_time': m[4],
                'end_date': f"{m[7]}-{m[6]}-{m[5]}",
                'end_time': m[8],
                'course': course_name
            })
    except Exception as e:
        print(f"   Error extracting events: {e}")

    return events


def create_ical(events: List[Dict]) -> Calendar:
    """Convert events to iCal format."""
    cal = Calendar()
    cal.add("prodid", "-//SmartWe Calendar Export//hpi.de//")
    cal.add("version", "2.0")
    cal.add("calscale", "GREGORIAN")
    cal.add("method", "PUBLISH")
    cal.add("x-wr-calname", "HPI SmartWe Calendar")

    for e in events:
        event = Event()
        event.add("summary", e['title'])

        try:
            start = datetime.strptime(f"{e['start_date']} {e['start_time']}", "%Y-%m-%d %H:%M")
            end = datetime.strptime(f"{e['end_date']} {e['end_time']}", "%Y-%m-%d %H:%M")
            event.add("dtstart", start)
            event.add("dtend", end)
        except:
            continue

        event.add("description", f"Course: {e['course']}")
        event.add("uid", f"{abs(hash(str(e)))}@smartwe.hpi.de")
        event.add("dtstamp", datetime.now())

        cal.add_component(event)

    return cal


def main():
    print("=" * 60)
    print("  HPI SmartWe Calendar Export Tool")
    print("=" * 60)

    all_events = []

    with sync_playwright() as p:
        print("\n🚀 Launching browser...")
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        context.set_default_timeout(15000)
        page = context.new_page()

        # Block Office 365 redirects
        page.route("**office.com**", lambda route: route.abort())
        page.route("**m365.cloud.microsoft**", lambda route: route.abort())

        print(f"   Navigating to {PORTAL_URL}")
        page.goto(PORTAL_URL, wait_until="domcontentloaded")

        print("\n🔐 Logging in...")
        if not do_login(page):
            print("❌ Login failed")
            browser.close()
            return

        print("✅ Login successful!\n")

        # Navigate to Meine Veranstaltungen
        print("📚 Loading enrolled courses...")
        page.locator("text=Veranstaltungen").first.click(force=True)
        time.sleep(2)
        page.locator("text=Meine Veranstaltungen").first.click(force=True)
        time.sleep(3)

        # Get course list
        body = page.inner_text("body")
        course_pattern = r'(SO\d+)\n([^\n]+)\nSO \d{4}'
        courses = re.findall(course_pattern, body)

        print(f"   Found {len(courses)} courses:\n")
        for code, name in courses:
            print(f"   • {name}")

        # Process each course
        for idx, (code, course_name) in enumerate(courses):
            print(f"\n📖 [{idx+1}/{len(courses)}] Processing: {course_name[:50]}...")

            try:
                # Ensure we're on course list
                current_body = page.inner_text("body", timeout=3000)
                if "Meine Veranstaltungen" not in current_body or code not in current_body:
                    page.locator("text=Veranstaltungen").first.click(force=True)
                    time.sleep(1)
                    page.locator("text=Meine Veranstaltungen").first.click(force=True)
                    time.sleep(2)

                # Click on course
                page.locator(f"text={code}").first.click()
                time.sleep(3)

                # Check for Termine section
                body = page.inner_text("body", timeout=3000)
                if "Termine" not in body:
                    print(f"   ⚠️  No Termine section")
                    page.locator("text=Meine Veranstaltungen").first.click(force=True)
                    time.sleep(2)
                    continue

                # Click Details in Termine section
                details_links = page.locator("text=Details").all()
                if details_links:
                    details_links[-1].click(force=True)
                    time.sleep(3)

                # Extract events
                events = extract_events_from_course(page, course_name)
                all_events.extend(events)
                print(f"   ✓ Found {len(events)} events")

                # Go back to course list
                page.locator("text=Meine Veranstaltungen").first.click(force=True)
                time.sleep(2)

            except Exception as e:
                print(f"   ❌ Error: {e}")
                try:
                    page.goto(f"{PORTAL_URL}#!app/smartdesign.campus.event", wait_until="domcontentloaded")
                    time.sleep(3)
                except:
                    pass

        print(f"\n📊 Total events collected: {len(all_events)}")

        # Create and save iCal
        cal = create_ical(all_events)
        with open(OUTPUT_FILE, "wb") as f:
            f.write(cal.to_ical())

        print(f"\n✅ Calendar exported to: {OUTPUT_FILE}")
        print(f"   Events exported: {len(all_events)}")

        browser.close()

    print("\n" + "=" * 60)
    print("  Export complete! Import the .ics file into your calendar app.")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    main()
