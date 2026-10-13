"""Calendar files (.ics) for Ignite events.

An .ics file is the one calendar format every phone understands: iPhone (Apple
Calendar), Samsung, Outlook and Google Calendar all import it. Each file carries
two reminders, so the member's phone alerts them a day before and two hours
before, whichever calendar app they use.

Used by app.py (download button) and server.py (the /calendar/<id>.ics link).
"""
import os
import sqlite3
from contextlib import closing
from datetime import date, datetime, timedelta, timezone

DB_FILENAME = "ignite_network.db"
PERSISTENT_DIR = "/data"
REMINDERS = [("-P1D", "Tomorrow"), ("-PT2H", "In 2 hours")]   # before the start


def resolve_db_path() -> str:
    """Same rule as app.py: IGNITE_DB_PATH, then Render's /data disk, then next to this file."""
    override = os.environ.get("IGNITE_DB_PATH", "").strip()
    if override:
        return override
    if os.path.isdir(PERSISTENT_DIR):
        return os.path.join(PERSISTENT_DIR, DB_FILENAME)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), DB_FILENAME)


def load_event(event_id: int):
    """Read one event straight from the database (read-only)."""
    path = resolve_db_path()
    if not os.path.exists(path):
        return None
    with closing(sqlite3.connect(f"file:{path}?mode=ro", uri=True, timeout=10)) as conn:
        conn.row_factory = sqlite3.Row
        row = conn.execute("SELECT * FROM events WHERE id = ?", (int(event_id),)).fetchone()
        return dict(row) if row else None


def _escape(text) -> str:
    return (str(text or "").replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,")
            .replace("\r\n", "\\n").replace("\n", "\\n"))


def _fold(line: str) -> str:
    """Calendar lines may be at most 75 bytes; longer ones continue on a line starting with a space."""
    raw = line.encode("utf-8")
    if len(raw) <= 75:
        return line
    parts, chunk = [], b""
    for ch in line:
        b = ch.encode("utf-8")
        if len(chunk) + len(b) > (75 if not parts else 74):
            parts.append(chunk.decode("utf-8"))
            chunk = b""
        chunk += b
    parts.append(chunk.decode("utf-8"))
    return "\r\n ".join(parts)


def event_times(ev: dict):
    """(start, end, all_day). Ghana time is UTC all year, so times are written in UTC."""
    if not ev.get("event_date"):
        return None
    try:
        day = date.fromisoformat(ev["event_date"])
    except ValueError:
        return None
    try:
        start = datetime.combine(day, datetime.strptime(ev.get("start_time") or "", "%H:%M").time())
    except ValueError:
        return day, day + timedelta(days=1), True
    try:
        end = datetime.combine(day, datetime.strptime(ev.get("end_time") or "", "%H:%M").time())
        if end <= start:              # e.g. 10 PM to 1 AM
            end += timedelta(days=1)
    except ValueError:
        end = start + timedelta(hours=3)
    return start, end, False


def build_ics(ev: dict, page_url: str = "", organiser: str = "Ignite Prayer Network") -> bytes:
    times = event_times(ev)
    if not times:
        raise ValueError("This event has no date yet.")
    start, end, all_day = times
    if all_day:
        when = [f"DTSTART;VALUE=DATE:{start:%Y%m%d}", f"DTEND;VALUE=DATE:{end:%Y%m%d}"]
    else:
        when = [f"DTSTART:{start:%Y%m%dT%H%M%S}Z", f"DTEND:{end:%Y%m%dT%H%M%S}Z"]
    details = []
    if ev.get("theme"):
        details.append(f"Theme: {ev['theme']}")
    if ev.get("about"):
        details.append(ev["about"])
    if ev.get("dress_code") or ev.get("dress_colors"):
        details.append("Dress code: " + " · ".join(x for x in (ev.get("dress_code"), ev.get("dress_colors")) if x))
    if page_url:
        details.append(page_url)
    lines = [
        "BEGIN:VCALENDAR", "VERSION:2.0", f"PRODID:-//{organiser}//Events//EN", "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH", "BEGIN:VEVENT",
        f"UID:ignite-event-{ev['id']}@igniteprayernetwork",
        f"DTSTAMP:{datetime.now(timezone.utc):%Y%m%dT%H%M%SZ}",
        *when,
        f"SUMMARY:{_escape(ev['event_name'])}",
    ]
    if ev.get("venue"):
        lines.append(f"LOCATION:{_escape(ev['venue'])}")
    if details:
        lines.append(f"DESCRIPTION:{_escape(chr(10).join(details))}")
    if page_url:
        lines.append(f"URL:{page_url}")
    for trigger, label in REMINDERS:
        lines += ["BEGIN:VALARM", "ACTION:DISPLAY", f"DESCRIPTION:{_escape(ev['event_name'])} · {label}",
                  f"TRIGGER:{trigger}", "END:VALARM"]
    lines += ["END:VEVENT", "END:VCALENDAR"]
    return ("\r\n".join(_fold(x) for x in lines) + "\r\n").encode("utf-8")


def ics_filename(ev: dict) -> str:
    safe = "".join(c if c.isalnum() else "_" for c in ev.get("event_name", "event")).strip("_")[:40]
    return f"{safe or 'event'}.ics"
