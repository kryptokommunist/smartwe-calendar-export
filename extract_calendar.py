#!/usr/bin/env python3
"""
SmartWe Calendar & Course Export Script

Extracts calendar entries from HPI SmartWe portal and exports them as iCal format.
"""

import re
import sys
from datetime import datetime
from pathlib import Path

from icalendar import Calendar, Event
from playwright.sync_api import sync_playwright, Page, TimeoutError as PlaywrightTimeout


PORTAL_URL = "https://sv-portal.hpi.de/SmartWe/"
CAMPUS_EVENTS_URL = "https://sv-portal.hpi.de/SmartWe/#!app/smartdesign.campus.event"
CALENDAR_URL = "https://sv-portal.hpi.de/SmartWe/#!app/smartdesign.calendar"
OUTPUT_FILE = Path(__file__).parent / "calendar_export.ics"


def wait_for_login(page: Page) -> None:
    """Wait for user to complete Microsoft SAML login."""
    print("\n🔐 Please complete the Microsoft login in the browser window...")
    print("   (Waiting for authentication to complete...)\n")

    # Wait until we're on the SmartWe portal (not a login page)
    while True:
        try:
            page.wait_for_url("**/SmartWe/**", timeout=5000)
            # Check if we're past the login by looking for the main app content
            if page.locator(".smartwe-app, .app-container, [class*='main'], .sw-main").first.is_visible(timeout=2000):
                break
        except PlaywrightTimeout:
            pass

        # Check if still on Microsoft login
        if "login.microsoftonline.com" in page.url or "login.live.com" in page.url:
            continue

        # Additional check - wait for any navigation element that indicates logged in state
        try:
            page.wait_for_selector(".sw-navigation, .navigation, nav, [class*='sidebar']", timeout=2000)
            break
        except PlaywrightTimeout:
            pass

    print("✅ Login successful!\n")


def extract_courses(page: Page) -> list[dict]:
    """Extract list of signed-up courses from campus events page."""
    print("📚 Navigating to campus events page...")
    page.goto(CAMPUS_EVENTS_URL)
    page.wait_for_load_state("networkidle")

    # Wait for the course list to load
    print("   Waiting for course list to load...")
    page.wait_for_timeout(3000)  # Give the SPA time to render

    courses = []

    # Try multiple selectors to find course items
    selectors = [
        ".event-item", ".course-item", ".campus-event",
        "[class*='event']", "[class*='course']",
        ".list-item", ".sw-list-item", "tr[class*='row']",
        ".card", ".sw-card"
    ]

    for selector in selectors:
        elements = page.locator(selector).all()
        if elements:
            print(f"   Found {len(elements)} items with selector: {selector}")
            for i, elem in enumerate(elements):
                try:
                    text = elem.inner_text(timeout=1000)
                    if text.strip():
                        courses.append({
                            "id": i,
                            "name": text.strip().split("\n")[0][:80],  # First line, truncated
                            "full_text": text.strip(),
                            "selector": selector,
                            "index": i
                        })
                except:
                    pass
            if courses:
                break

    # Fallback: extract from page content
    if not courses:
        print("   Using fallback extraction method...")
        content = page.content()
        # Look for event-related data in the page
        page.screenshot(path="debug_courses.png")
        print("   📸 Screenshot saved to debug_courses.png for inspection")

    return courses


def select_courses(courses: list[dict]) -> list[dict]:
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


def extract_calendar_events(page: Page, selected_courses: list[dict]) -> list[dict]:
    """Extract calendar events for selected courses."""
    print("📅 Navigating to calendar page...")
    page.goto(CALENDAR_URL)
    page.wait_for_load_state("networkidle")

    print("   Waiting for calendar to load...")
    page.wait_for_timeout(3000)

    events = []

    # Try to find calendar events with various selectors
    selectors = [
        ".calendar-event", ".event", ".fc-event",  # FullCalendar style
        "[class*='calendar'] [class*='event']",
        ".appointment", ".sw-appointment",
        ".fc-event-container", ".fc-content",
        "[data-event]", "[class*='schedule']"
    ]

    for selector in selectors:
        elements = page.locator(selector).all()
        if elements:
            print(f"   Found {len(elements)} calendar items with selector: {selector}")
            for i, elem in enumerate(elements):
                try:
                    text = elem.inner_text(timeout=1000)

                    # Try to extract time information from element or attributes
                    event_data = {
                        "id": i,
                        "title": text.strip().split("\n")[0] if text.strip() else f"Event {i}",
                        "full_text": text.strip(),
                        "start": None,
                        "end": None,
                        "location": "",
                        "description": text.strip()
                    }

                    # Try to get data attributes
                    for attr in ["data-start", "data-end", "data-date", "data-time"]:
                        try:
                            val = elem.get_attribute(attr)
                            if val:
                                event_data[attr] = val
                        except:
                            pass

                    events.append(event_data)
                except:
                    pass
            if events:
                break

    # Filter events by selected courses if we have course names
    if selected_courses:
        course_names = [c["name"].lower() for c in selected_courses]
        filtered_events = []
        for event in events:
            event_text = (event.get("title", "") + " " + event.get("full_text", "")).lower()
            for course_name in course_names:
                # Match if any significant part of course name is in event
                words = [w for w in course_name.split() if len(w) > 3]
                if any(word in event_text for word in words) or not words:
                    filtered_events.append(event)
                    break

        if filtered_events:
            print(f"   Filtered to {len(filtered_events)} events matching selected courses")
            events = filtered_events

    # Take screenshot for debugging if no events found
    if not events:
        page.screenshot(path="debug_calendar.png")
        print("   📸 Screenshot saved to debug_calendar.png for inspection")

    return events


def parse_datetime(text: str) -> datetime | None:
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


def create_ical(events: list[dict]) -> Calendar:
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

        # UID
        uid = f"smartwe-{i}-{hash(title)}@hpi.de"
        event.add("uid", uid)

        # Timestamps
        now = datetime.now()
        event.add("dtstamp", now)

        # Try to parse start/end times
        start = None
        end = None

        for key in ["data-start", "start", "data-date"]:
            if key in event_data and event_data[key]:
                start = parse_datetime(event_data[key])
                if start:
                    break

        for key in ["data-end", "end"]:
            if key in event_data and event_data[key]:
                end = parse_datetime(event_data[key])
                if end:
                    break

        # Default to today if no date found
        if not start:
            start = now

        if not end:
            end = start

        event.add("dtstart", start)
        event.add("dtend", end)

        # Description
        if event_data.get("description"):
            event.add("description", event_data["description"])

        # Location
        if event_data.get("location"):
            event.add("location", event_data["location"])

        cal.add_component(event)

    return cal


def main():
    print("=" * 60)
    print("  HPI SmartWe Calendar Export Tool")
    print("=" * 60)

    with sync_playwright() as p:
        # Launch browser in non-headless mode for login
        print("\n🚀 Launching browser...")
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        # Navigate to portal
        print(f"   Navigating to {PORTAL_URL}")
        page.goto(PORTAL_URL)

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

    print("\n" + "=" * 60)
    print("  Export complete! Import the .ics file into your calendar app.")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    main()
