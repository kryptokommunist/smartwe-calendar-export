#!/usr/bin/env python3
"""
SmartWe Calendar & Course Export Script

Extracts calendar entries from HPI SmartWe portal and exports them as iCal format.
"""

import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict

from icalendar import Calendar, Event
from playwright.sync_api import sync_playwright, Page, TimeoutError as PlaywrightTimeout


PORTAL_URL = "https://sv-portal.hpi.de/SmartWe/"
CAMPUS_EVENTS_URL = "https://sv-portal.hpi.de/SmartWe/#!app/smartdesign.campus.event"
CALENDAR_URL = "https://sv-portal.hpi.de/SmartWe/#!app/smartdesign.calendar"
OUTPUT_FILE = Path(__file__).parent / "calendar_export.ics"
CREDENTIALS_FILE = Path(__file__).parent / ".credentials"


def load_credentials(page_match: str = "login.microsoftonline") -> Optional[Dict[str, str]]:
    """Load credentials from .credentials file matching the page URL."""
    if CREDENTIALS_FILE.exists():
        try:
            with open(CREDENTIALS_FILE, 'r') as f:
                content = f.read().strip()

            # Parse multiple credential blocks (separated by blank lines)
            blocks = content.split('\n\n')
            for block in blocks:
                lines = block.strip().split('\n')
                creds = {}
                for line in lines:
                    if ':' in line:
                        key, val = line.split(':', 1)
                        creds[key.strip()] = val.strip()
                    elif '=' in line:
                        key, val = line.split('=', 1)
                        creds[key.strip()] = val.strip()

                # Check if this block matches the page
                if 'page' in creds and page_match in creds['page']:
                    # Normalize keys
                    if 'user' in creds:
                        creds['username'] = creds['user']
                    if 'pw' in creds:
                        creds['password'] = creds['pw']

                    print(f"   Found credentials for: {creds.get('page', 'unknown')[:40]}...")
                    print(f"   Username: {creds.get('username', '')[:3]}***")
                    return creds

            print(f"   No credentials found matching '{page_match}'")
            return None
        except Exception as e:
            print(f"   Warning: Could not load credentials: {e}")
    return None


def do_full_login(page: Page) -> bool:
    """Handle complete login flow: MS -> ADFS -> SmartWe. Returns True when app is loaded."""
    max_attempts = 60  # 2 minutes max

    for attempt in range(max_attempts):
        page.wait_for_timeout(2000)
        url = page.url

        try:
            body = page.inner_text("body", timeout=3000)
        except:
            print(f"   [{attempt}] Loading...")
            continue

        print(f"   [{attempt}] {url[:55]}...")

        # Check if app is loaded (success)
        if "sv-portal.hpi.de/SmartWe" in url:
            if any(kw in body for kw in ["Kalender", "Termine", "Apps", "Veranstaltungen"]) and "Anmelden" not in body:
                return True

            # SmartWe login form
            if "Anmelden" in body and "nicht korrekt" not in body:
                user_field = page.locator('input[name="username"]').first
                pw_field = page.locator('input[name="password"]').first
                if user_field.is_visible(timeout=1000) and pw_field.is_visible(timeout=1000):
                    creds = load_credentials("sv-portal")
                    if creds:
                        print("   SmartWe: Logging in...")
                        user_field.fill(creds['username'])
                        pw_field.fill(creds['password'])
                        # Check "Automatisch anmelden" if available
                        try:
                            auto_login_cb = page.get_by_label("Automatisch anmelden")
                            if auto_login_cb.is_visible(timeout=1000):
                                auto_login_cb.check()
                        except:
                            pass
                        page.get_by_role("button", name="Anmelden").click()
                        page.wait_for_timeout(3000)
                        continue

            # Wrong credentials - stop
            if "nicht korrekt" in body:
                print("   ⚠️ SmartWe credentials incorrect!")
                return False

        # Microsoft login - email entry
        if "login.microsoftonline" in url:
            if "Stay signed in" in body:
                print("   MS: Stay signed in -> Yes")
                page.click("#idSIButton9")
                continue

            email_field = page.locator("#i0116")
            if email_field.is_visible(timeout=1000):
                creds = load_credentials("login.microsoftonline")
                if creds:
                    print("   MS: Entering email...")
                    email_field.fill(creds['username'])
                    page.click("#idSIButton9")
                    continue

        # HPI ADFS login
        if "adfs.hpi" in url:
            pw_field = page.locator('input[type="password"]')
            if pw_field.is_visible(timeout=1000):
                creds = load_credentials("adfs.hpi")
                if creds:
                    print("   ADFS: Entering password...")
                    pw_field.fill(creds['password'])
                    page.get_by_role("button", name="Sign in").click()
                    continue

    return False


def wait_for_login(page: Page) -> None:
    """Handle full login flow: MS -> ADFS -> SmartWe."""
    print("\n🔐 Logging in (MS -> ADFS -> SmartWe)...")

    if do_full_login(page):
        print("✅ Login successful!\n")
    else:
        print("⚠️  Auto-login incomplete, please complete manually...")
        # Wait for manual completion
        for _ in range(60):
            if "sv-portal.hpi.de/SmartWe" in page.url:
                body = page.inner_text("body", timeout=3000)
                if any(kw in body for kw in ["Kalender", "Termine", "Apps"]):
                    print("✅ Login successful!\n")
                    return
            page.wait_for_timeout(5000)


def extract_courses(page: Page) -> List[Dict]:
    """Extract list of signed-up courses from campus events page."""
    print("📚 Extracting courses from current page...")

    # The courses are shown on the main dashboard under "Veranstaltungen"
    # Wait for content to load
    page.wait_for_timeout(3000)

    courses = []

    # Look for course list items - they typically have course name and course ID
    # Based on screenshot: items like "Betriebssysteme II\nSO 2026, SO250000307"
    try:
        # Find all clickable list items in the Veranstaltungen section
        # These appear to be in a list with course names and IDs
        all_text = page.inner_text("body")

        # Parse course entries from the page - look for patterns like course names followed by semester codes
        import re
        # Match course entries: Name followed by semester code pattern
        course_pattern = r'([A-Za-z][^\n]{5,80})\n(SO \d{4}, [A-Z0-9]+(?:, [A-Za-z]+)?)'
        matches = re.findall(course_pattern, all_text)

        for i, (name, code) in enumerate(matches):
            name = name.strip()
            # Filter out UI elements
            if any(skip in name.lower() for skip in ['suche', 'details', 'ansichten', 'alle anzeigen', 'meine']):
                continue
            courses.append({
                "id": i,
                "name": name,
                "code": code.strip(),
                "full_text": f"{name} ({code})"
            })

        if courses:
            print(f"   Found {len(courses)} courses")

    except Exception as e:
        print(f"   Error extracting courses: {e}")

    # Fallback: take screenshot for debugging
    if not courses:
        print("   Using fallback extraction method...")
        page.screenshot(path="debug_courses.png")
        print("   📸 Screenshot saved to debug_courses.png for inspection")

    return courses


def select_courses(courses: List[Dict]) -> List[Dict]:
    """Present course selection to user via CLI prompt."""
    if not courses:
        print("\n⚠️  No courses found. Proceeding with all calendar events.\n")
        return []

    print("\n📋 Found the following courses/events:\n")
    for i, course in enumerate(courses, 1):
        name = course["name"][:60] + "..." if len(course["name"]) > 60 else course["name"]
        print(f"   [{i}] {name}")

    print(f"\n   [0] Select ALL courses")
    print(f"   [q] Quit\n")

    while True:
        selection = input("Enter course numbers (comma-separated) or 0 for all: ").strip()

        if selection.lower() == 'q':
            print("Exiting...")
            sys.exit(0)

        if selection == '0':
            return courses

        try:
            indices = [int(x.strip()) for x in selection.split(",")]
            selected = [courses[i-1] for i in indices if 1 <= i <= len(courses)]
            if selected:
                print(f"\n✅ Selected {len(selected)} course(s)\n")
                return selected
            else:
                print("Invalid selection. Please try again.")
        except (ValueError, IndexError):
            print("Invalid input. Enter numbers separated by commas, 0 for all, or q to quit.")


def extract_calendar_events(page: Page, selected_courses: List[Dict]) -> List[Dict]:
    """Extract calendar events via Termine -> Meine Termine."""
    print("📅 Navigating to Termine -> Meine Termine...")

    events = []

    try:
        # Click on the Termine tile
        page.locator("text=Termine").first.click(force=True)
        page.wait_for_timeout(3000)

        # Click "Meine Termine"
        page.locator("text=Meine Termine").first.click(force=True)
        page.wait_for_timeout(5000)

        # Get total event count
        body_text = page.inner_text("body")
        total_match = re.search(r'Gesamt:\s*(\d+)', body_text)
        expected_total = int(total_match.group(1)) if total_match else 0
        print(f"   Total events in list: {expected_total}")

        # Pattern to extract events
        event_pattern = r'([^\n]{3,100})\n(\d{2})\.(\d{2})\.(\d{4}),\s*(\d{1,2}:\d{2})\n(\d{2})\.(\d{2})\.(\d{4}),\s*(\d{1,2}:\d{2})'
        skip_words = ['betreff', 'beginn', 'ende', 'daten geladen', 'gesamt', 'meine termine', 'geändert']
        seen = set()

        # Scroll through the virtualized table to collect all events
        print("   Scrolling to collect all events...")
        for scroll in range(100):  # Max 100 scroll iterations
            body_text = page.inner_text("body")
            matches = re.findall(event_pattern, body_text)

            for match in matches:
                title = match[0].strip()
                start_day, start_month, start_year, start_time = match[1], match[2], match[3], match[4]
                end_day, end_month, end_year, end_time = match[5], match[6], match[7], match[8]

                if any(skip in title.lower() for skip in skip_words):
                    continue
                if len(title) < 3:
                    continue

                key = f"{title}-{start_day}.{start_month}.{start_year}-{start_time}"
                if key in seen:
                    continue
                seen.add(key)

                events.append({
                    "id": len(events),
                    "title": title,
                    "date": f"{start_year}-{start_month}-{start_day}",
                    "start_time": start_time,
                    "end_time": end_time,
                    "full_text": f"{title} ({start_day}.{start_month}.{start_year}, {start_time} - {end_time})"
                })

            # Check if we have all events
            if len(events) >= expected_total:
                break

            # Scroll down
            page.keyboard.press("End")
            page.wait_for_timeout(500)

            if scroll % 10 == 0:
                print(f"   ... collected {len(events)}/{expected_total} events")

        page.screenshot(path="debug_meine_termine.png")
        print(f"   Found {len(events)} events (expected {expected_total})")

    except Exception as e:
        print(f"   Error extracting events: {e}")
        import traceback
        traceback.print_exc()
        page.screenshot(path="debug_error.png")

    # Filter by selected courses if provided
    if selected_courses and events:
        course_names = [c["name"].lower() for c in selected_courses]
        filtered = []
        for event in events:
            event_title = event.get("title", "").lower()
            for course_name in course_names:
                words = [w for w in course_name.split() if len(w) > 3]
                if any(word in event_title for word in words):
                    filtered.append(event)
                    break
        if filtered:
            print(f"   Filtered to {len(filtered)} events matching selected courses")
            events = filtered

    return events


def parse_datetime(text: str) -> Optional[datetime]:
    """Try to parse datetime from various formats."""
    formats = [
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%d.%m.%Y %H:%M",
        "%d.%m.%Y",
        "%Y-%m-%d",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(text.strip(), fmt)
        except ValueError:
            continue
    return None


def create_ical(events: List[Dict]) -> Calendar:
    """Convert events to iCal format."""
    cal = Calendar()
    cal.add("prodid", "-//SmartWe Calendar Export//hpi.de//")
    cal.add("version", "2.0")
    cal.add("calscale", "GREGORIAN")
    cal.add("method", "PUBLISH")
    cal.add("x-wr-calname", "HPI SmartWe Calendar")

    for i, event_data in enumerate(events):
        event = Event()

        # Summary/Title
        title = event_data.get("title", f"Event {i}")
        event.add("summary", title)

        # UID - use hash of title + date + time for uniqueness
        uid_base = f"{title}-{event_data.get('date', '')}-{event_data.get('start_time', '')}"
        uid = f"smartwe-{abs(hash(uid_base))}@hpi.de"
        event.add("uid", uid)

        # Timestamps
        now = datetime.now()
        event.add("dtstamp", now)

        # Parse start/end times
        start = None
        end = None

        date_str = event_data.get("date", now.strftime("%Y-%m-%d"))
        start_time = event_data.get("start_time", "")
        end_time = event_data.get("end_time", "")

        if start_time:
            try:
                start = datetime.strptime(f"{date_str} {start_time}", "%Y-%m-%d %H:%M")
            except:
                start = now

        if end_time:
            try:
                end = datetime.strptime(f"{date_str} {end_time}", "%Y-%m-%d %H:%M")
            except:
                end = start if start else now

        if not start:
            start = now
        if not end:
            end = start

        event.add("dtstart", start)
        event.add("dtend", end)

        # Description
        desc = event_data.get("full_text", "")
        if desc:
            event.add("description", desc)

        # Location
        if event_data.get("location"):
            event.add("location", event_data["location"])

        cal.add_component(event)

    return cal


def main():
    print("=" * 60)
    print("  HPI SmartWe Calendar Export Tool")
    print("=" * 60)

    browser = None
    try:
        with sync_playwright() as p:
            # Launch browser in non-headless mode for login
            print("\n🚀 Launching browser...")
            browser = p.chromium.launch(headless=False, slow_mo=100)
            context = browser.new_context(
                viewport={"width": 1280, "height": 800}
            )
            page = context.new_page()

            # Navigate to portal
            print(f"   Navigating to {PORTAL_URL}")
            page.goto(PORTAL_URL, wait_until="domcontentloaded")

            # Wait for user to login
            wait_for_login(page)

            # Extract courses
            courses = extract_courses(page)

            # Let user select courses
            selected_courses = select_courses(courses)

            # Extract calendar events
            events = extract_calendar_events(page, selected_courses)

            if not events:
                print("\n⚠️  No calendar events found.")
                print("   This might be due to the page structure.")
                print("   Check debug_calendar.png for the page layout.\n")
            else:
                print(f"\n📊 Found {len(events)} calendar event(s)")

            # Create iCal
            cal = create_ical(events)

            # Save to file
            with open(OUTPUT_FILE, "wb") as f:
                f.write(cal.to_ical())

            print(f"\n✅ Calendar exported to: {OUTPUT_FILE}")
            print(f"   Events exported: {len(events)}")

            # Cleanup
            browser.close()

    except KeyboardInterrupt:
        print("\n\n⚠️  Interrupted by user")
        if browser:
            browser.close()
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Error: {e}")
        if browser:
            try:
                browser.close()
            except:
                pass
        raise

    print("\n" + "=" * 60)
    print("  Export complete! Import the .ics file into your calendar app.")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    main()
