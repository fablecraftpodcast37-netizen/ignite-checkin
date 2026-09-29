"""
Ignite Prayer Network: Event Registration & Check-In System
===========================================================

One Streamlit app, three pages:

  * Day-of check-in (default). Each event's check-in QR code opens
        https://<your-app>.streamlit.app/?event=<id>
  * Pre-registration, shared before the programme (e.g. on WhatsApp):
        https://<your-app>.streamlit.app/?event=<id>&mode=register
  * Admin portal (not linked from any public page, bookmark it):
        https://<your-app>.streamlit.app/?view=admin

How records are kept apart
  * Every event (Asteri, Shekinah Glory, Impromptu) has its own event_id, and
    every record is written and read with it, so events never mix.
  * Pre-registrations live in their own table (pre_registrations), separate
    from the day-of check-ins (attendance). Each has its own ledger and export.
  * Both share one members list, so someone who pre-registers only needs their
    phone number to check in on the day.

Optional settings (Streamlit Cloud: Settings -> Secrets):
    ADMIN_USERNAME = "admin"                      # main administrator
    ADMIN_PASSWORD = "your-strong-password"
    APP_BASE_URL   = "https://your-app-name.streamlit.app"

Team accounts: the main administrator (and any Admin) can give trusted
people their own sign-in from the Team tab, as an Admin or an Usher.
Every change made in the admin portal is written to an activity log.
"""

import hashlib
import hmac
import html
import io
import os
import re
import sqlite3
import tempfile
import time
from contextlib import closing
from datetime import date, datetime, timezone
from urllib.parse import quote, urlsplit

import pandas as pd
import qrcode
import streamlit as st
from PIL import Image, ImageDraw, ImageFont, ImageOps, UnidentifiedImageError

# ---------------------------------------------------------------------------
# Settings (change here, or override with Streamlit secrets)
# ---------------------------------------------------------------------------
APP_NAME = "Ignite Prayer Network"
DB_PATH = "ignite_network.db"
DEFAULT_ADMIN_USERNAME = "admin"              # main administrator's username
DEFAULT_ADMIN_PASSWORD = "IgniteAdmin2026"     # change before going live (use Secrets)
TEAM_ROLES = {
    "Admin": "Full access: events, records, exports, backups and the team",
    "Usher": "Event-day helper: live check-in list and manual check-in only",
}
EVENT_TYPES = ["Asteri", "Shekinah Glory", "Impromptu"]

CONFIRM_SECONDS = 6          # how long the "Submitted, thank you" screen stays up
ADMIN_SESSION_MINUTES = 30   # admin is logged out after this much idle time
MAX_LOGIN_ATTEMPTS = 5       # wrong passwords allowed before a cool-down
LOCKOUT_SECONDS = 60
FLYER_MAX_UPLOAD_MB = 10
FLYER_MAX_WIDTH = 1200
LEDGER_REFRESH_SECONDS = 15

BRAND = "#E4572E"            # Ignite flame orange
BRAND_DARK = "#B83A17"
ACCENT = "#F3A712"           # gold
# What members see for each programme type. Change these freely.
PUBLIC_TYPE_NAMES = {"Asteri": "Asteri", "Shekinah Glory": "Shekinah Glory", "Impromptu": "Special Meeting"}
TYPE_COLOURS = {"Asteri": "#6C4AB6", "Shekinah Glory": "#B8860B", "Impromptu": "#1B998B"}
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def get_setting(key: str, default: str = "") -> str:
    """Read a value from Streamlit secrets, falling back to the default."""
    try:
        return str(st.secrets[key])
    except Exception:
        return default


# ---------------------------------------------------------------------------
# Database
# ---------------------------------------------------------------------------
SCHEMA = """
CREATE TABLE IF NOT EXISTS events (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    event_name  TEXT NOT NULL,
    event_type  TEXT NOT NULL
                CHECK (event_type IN ('Asteri', 'Shekinah Glory', 'Impromptu')),
    event_date  TEXT,                         -- YYYY-MM-DD, optional
    venue       TEXT,
    is_open     INTEGER NOT NULL DEFAULT 1,   -- 1 = day-of check-in is open
    prereg_open INTEGER NOT NULL DEFAULT 1,   -- 1 = pre-registration link works
    created_at  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    flyer_bytes BLOB                          -- programme flyer, stored as JPEG
);

CREATE TABLE IF NOT EXISTS members (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    full_name               TEXT NOT NULL,
    phone_number            TEXT NOT NULL UNIQUE,
    whatsapp_number         TEXT,
    email                   TEXT,
    emergency_contact_name  TEXT,
    emergency_contact_phone TEXT,
    parent_guardian_phone   TEXT,
    is_minor                INTEGER NOT NULL DEFAULT 0,
    consent_at              DATETIME,         -- when they agreed to data storage
    created_at              DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
);

-- Day-of check-ins. Timestamps come from the database clock, never from the
-- member's phone. One check-in per person per event, so the first arrival
-- time is the one that stands.
CREATE TABLE IF NOT EXISTS attendance (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    member_id           INTEGER NOT NULL REFERENCES members(id),
    event_id            INTEGER NOT NULL REFERENCES events(id),
    check_in_timestamp  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (member_id, event_id)
);

-- Registrations made before the day through the shared link.
-- Kept in their own table so they never mix with day-of check-ins.
CREATE TABLE IF NOT EXISTS pre_registrations (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    member_id      INTEGER NOT NULL REFERENCES members(id),
    event_id       INTEGER NOT NULL REFERENCES events(id),
    registered_at  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (member_id, event_id)
);

-- Trusted team members with their own sign-in (the main administrator
-- signs in with the username/password from Streamlit Secrets instead).
CREATE TABLE IF NOT EXISTS admins (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    full_name      TEXT NOT NULL,
    username       TEXT NOT NULL UNIQUE COLLATE NOCASE,
    password_hash  TEXT NOT NULL,                 -- salted PBKDF2, never the password
    role           TEXT NOT NULL CHECK (role IN ('Admin', 'Usher')),
    is_active      INTEGER NOT NULL DEFAULT 1,
    created_by     TEXT,
    created_at     DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    last_login     DATETIME
);

-- Who did what, and when, in the admin portal.
CREATE TABLE IF NOT EXISTS activity_log (
    id      INTEGER PRIMARY KEY AUTOINCREMENT,
    at      DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    actor   TEXT NOT NULL,
    action  TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_attendance_event  ON attendance(event_id);
CREATE INDEX IF NOT EXISTS idx_attendance_member ON attendance(member_id);
CREATE INDEX IF NOT EXISTS idx_prereg_event      ON pre_registrations(event_id);
CREATE INDEX IF NOT EXISTS idx_prereg_member     ON pre_registrations(member_id);

-- Arrival times can't be edited. Records can only be removed as a whole,
-- by an admin deleting an event or a member.
CREATE TRIGGER IF NOT EXISTS attendance_no_update
BEFORE UPDATE ON attendance
BEGIN
    SELECT RAISE(ABORT, 'Attendance records are read-only');
END;
"""

# Columns added in later versions. Older databases get them automatically.
MIGRATIONS = [
    ("events", "flyer_bytes", "BLOB"),
    ("events", "event_date", "TEXT"),
    ("events", "venue", "TEXT"),
    ("events", "is_open", "INTEGER NOT NULL DEFAULT 1"),
    ("events", "prereg_open", "INTEGER NOT NULL DEFAULT 1"),
    ("members", "is_minor", "INTEGER NOT NULL DEFAULT 0"),
    ("members", "consent_at", "DATETIME"),
]


def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, timeout=15, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def migrate(conn: sqlite3.Connection):
    """Create missing tables first, add missing columns, then the rest
    (indexes and triggers), so older databases upgrade cleanly."""
    existing_tables = {r["name"] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    if "events" in existing_tables or "members" in existing_tables:
        for table, column, decl in MIGRATIONS:
            if table in existing_tables:
                cols = {r["name"] for r in conn.execute(f"PRAGMA table_info({table})")}
                if column not in cols:
                    conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {decl}")
    conn.executescript(SCHEMA)


@st.cache_resource
def init_db() -> bool:
    """Create or upgrade the tables once per server process."""
    with closing(get_conn()) as conn, conn:
        migrate(conn)
    return True


def query_df(sql: str, params=()) -> pd.DataFrame:
    with closing(get_conn()) as conn:
        return pd.read_sql_query(sql, conn, params=params)


# ----- events ---------------------------------------------------------------
EVENT_COLUMNS = """e.id, e.event_name, e.event_type, e.event_date, e.venue, e.is_open,
                   e.prereg_open, e.created_at, e.flyer_bytes IS NOT NULL AS has_flyer,
                   (SELECT COUNT(*) FROM attendance a WHERE a.event_id = e.id) AS attendees,
                   (SELECT COUNT(*) FROM pre_registrations p WHERE p.event_id = e.id) AS preregs"""


def get_events(open_only: bool = False) -> list[dict]:
    where = "WHERE e.is_open = 1" if open_only else ""
    with closing(get_conn()) as conn:
        rows = conn.execute(
            f"""SELECT {EVENT_COLUMNS} FROM events e {where}
                ORDER BY COALESCE(e.event_date, date(e.created_at)) DESC, e.id DESC"""
        ).fetchall()
        return [dict(r) for r in rows]


def get_event(event_id: int):
    with closing(get_conn()) as conn:
        row = conn.execute(f"SELECT {EVENT_COLUMNS} FROM events e WHERE e.id = ?", (event_id,)).fetchone()
        return dict(row) if row else None


def create_event(name, event_type, event_date=None, venue=None, flyer=None,
                 is_open=True, prereg_open=True) -> int:
    with closing(get_conn()) as conn, conn:
        cur = conn.execute(
            """INSERT INTO events (event_name, event_type, event_date, venue, flyer_bytes,
                                   is_open, prereg_open)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (name.strip(), event_type, event_date, (venue or "").strip() or None, flyer,
             int(is_open), int(prereg_open)),
        )
    get_flyer.clear()
    return cur.lastrowid


def update_event(event_id, name, event_type, event_date, venue, is_open, prereg_open):
    with closing(get_conn()) as conn, conn:
        conn.execute(
            """UPDATE events SET event_name = ?, event_type = ?, event_date = ?, venue = ?,
                                 is_open = ?, prereg_open = ?
               WHERE id = ?""",
            (name.strip(), event_type, event_date, (venue or "").strip() or None,
             int(is_open), int(prereg_open), event_id),
        )


def set_event_flag(event_id: int, column: str, value: bool):
    assert column in ("is_open", "prereg_open")
    with closing(get_conn()) as conn, conn:
        conn.execute(f"UPDATE events SET {column} = ? WHERE id = ?", (int(value), event_id))


def delete_event(event_id: int) -> tuple[int, int]:
    """Remove an event with its check-ins and pre-registrations in one
    transaction. Members stay, because they may belong to other events.
    Returns (check-ins removed, pre-registrations removed)."""
    with closing(get_conn()) as conn, conn:
        checkins = conn.execute("DELETE FROM attendance WHERE event_id = ?", (event_id,)).rowcount
        preregs = conn.execute("DELETE FROM pre_registrations WHERE event_id = ?", (event_id,)).rowcount
        conn.execute("DELETE FROM events WHERE id = ?", (event_id,))
    get_flyer.clear()
    return checkins, preregs


def set_event_flyer(event_id: int, flyer):
    with closing(get_conn()) as conn, conn:
        conn.execute("UPDATE events SET flyer_bytes = ? WHERE id = ?", (flyer, event_id))
    get_flyer.clear()


@st.cache_data(ttl=600, show_spinner=False)
def get_flyer(event_id: int):
    """Cached, so a queue of people doesn't reload the image on every tap.
    Cleared whenever a flyer changes."""
    with closing(get_conn()) as conn:
        row = conn.execute("SELECT flyer_bytes FROM events WHERE id = ?", (event_id,)).fetchone()
        return bytes(row["flyer_bytes"]) if row and row["flyer_bytes"] else None


# ----- members ---------------------------------------------------------------
def find_member_id(phone: str):
    with closing(get_conn()) as conn:
        row = conn.execute("SELECT id FROM members WHERE phone_number = ?", (phone,)).fetchone()
        return row["id"] if row else None


def create_member(data: dict) -> int:
    with closing(get_conn()) as conn, conn:
        cur = conn.execute(
            """INSERT INTO members (full_name, phone_number, whatsapp_number, email,
                                    emergency_contact_name, emergency_contact_phone,
                                    parent_guardian_phone, is_minor, consent_at)
               VALUES (:full_name, :phone_number, :whatsapp_number, :email,
                       :emergency_contact_name, :emergency_contact_phone,
                       :parent_guardian_phone, :is_minor, CURRENT_TIMESTAMP)""",
            data,
        )
        return cur.lastrowid


def get_member(member_id: int):
    with closing(get_conn()) as conn:
        row = conn.execute("SELECT * FROM members WHERE id = ?", (member_id,)).fetchone()
        return dict(row) if row else None


def update_member(member_id: int, data: dict):
    """Raises sqlite3.IntegrityError if the new phone number belongs to someone else."""
    with closing(get_conn()) as conn, conn:
        conn.execute(
            """UPDATE members SET full_name = :full_name, phone_number = :phone_number,
                   whatsapp_number = :whatsapp_number, email = :email,
                   emergency_contact_name = :emergency_contact_name,
                   emergency_contact_phone = :emergency_contact_phone,
                   parent_guardian_phone = :parent_guardian_phone, is_minor = :is_minor
               WHERE id = :id""",
            {**data, "id": member_id},
        )


def delete_member(member_id: int) -> int:
    """Remove a person with all their check-ins and pre-registrations."""
    with closing(get_conn()) as conn, conn:
        removed = conn.execute("DELETE FROM attendance WHERE member_id = ?", (member_id,)).rowcount
        removed += conn.execute("DELETE FROM pre_registrations WHERE member_id = ?", (member_id,)).rowcount
        conn.execute("DELETE FROM members WHERE id = ?", (member_id,))
    return removed


def search_members(term: str) -> pd.DataFrame:
    term = (term or "").strip()
    digits = re.sub(r"\D", "", term)
    return query_df(
        """SELECT m.id,
                  m.full_name AS "Full Name",
                  m.phone_number AS "Phone",
                  m.email AS "Email",
                  CASE WHEN m.is_minor THEN 'Yes' ELSE '' END AS "Under 18",
                  (SELECT COUNT(*) FROM attendance a WHERE a.member_id = m.id) AS "Events Attended",
                  (SELECT COUNT(*) FROM pre_registrations p WHERE p.member_id = m.id) AS "Pre-Registrations",
                  (SELECT strftime('%Y-%m-%d %H:%M', MAX(a.check_in_timestamp))
                     FROM attendance a WHERE a.member_id = m.id) AS "Last Check-In",
                  strftime('%Y-%m-%d', m.created_at) AS "Registered"
           FROM members m
           WHERE (? = '' OR m.full_name LIKE ? OR m.phone_number LIKE ?)
           ORDER BY m.full_name COLLATE NOCASE""",
        (term, f"%{term}%", f"%{digits or term}%"),
    )


def all_members_export() -> pd.DataFrame:
    return query_df(
        """SELECT m.full_name AS "Full Name", m.phone_number AS "Phone",
                  m.whatsapp_number AS "WhatsApp", m.email AS "Email",
                  CASE WHEN m.is_minor THEN 'Yes' ELSE '' END AS "Under 18",
                  m.parent_guardian_phone AS "Parent/Guardian Phone",
                  m.emergency_contact_name AS "Emergency Contact",
                  m.emergency_contact_phone AS "Emergency Phone",
                  (SELECT COUNT(*) FROM attendance a WHERE a.member_id = m.id) AS "Events Attended",
                  strftime('%Y-%m-%d', m.created_at) AS "Registered"
           FROM members m ORDER BY m.full_name COLLATE NOCASE"""
    )


def member_history(member_id: int) -> pd.DataFrame:
    return query_df(
        """SELECT e.event_name AS "Event", e.event_type AS "Type",
                  COALESCE(strftime('%Y-%m-%d %H:%M', p.registered_at), '') AS "Pre-Registered",
                  COALESCE(strftime('%Y-%m-%d %H:%M:%S', a.check_in_timestamp), '') AS "Checked In"
           FROM events e
           LEFT JOIN pre_registrations p ON p.event_id = e.id AND p.member_id = :m
           LEFT JOIN attendance a        ON a.event_id = e.id AND a.member_id = :m
           WHERE p.id IS NOT NULL OR a.id IS NOT NULL
           ORDER BY COALESCE(e.event_date, date(e.created_at)) DESC, e.id DESC""",
        {"m": member_id},
    )


# ----- check-ins and pre-registrations ---------------------------------------
def record_check_in(member_id: int, event_id: int) -> bool:
    """True for a new check-in, False if they were already checked in."""
    try:
        with closing(get_conn()) as conn, conn:
            conn.execute("INSERT INTO attendance (member_id, event_id) VALUES (?, ?)", (member_id, event_id))
        return True
    except sqlite3.IntegrityError:
        return False


def record_pre_registration(member_id: int, event_id: int) -> bool:
    """True for a new pre-registration, False if they had already registered."""
    try:
        with closing(get_conn()) as conn, conn:
            conn.execute("INSERT INTO pre_registrations (member_id, event_id) VALUES (?, ?)",
                         (member_id, event_id))
        return True
    except sqlite3.IntegrityError:
        return False


def is_pre_registered(member_id: int, event_id: int) -> bool:
    with closing(get_conn()) as conn:
        return conn.execute("SELECT 1 FROM pre_registrations WHERE member_id = ? AND event_id = ?",
                            (member_id, event_id)).fetchone() is not None


def get_ledger(event_id: int) -> pd.DataFrame:
    """Day-of check-ins for ONE event only."""
    df = query_df(
        """SELECT strftime('%Y-%m-%d %H:%M:%S', a.check_in_timestamp) AS "Check-In Time",
                  m.full_name               AS "Full Name",
                  m.phone_number            AS "Phone",
                  m.whatsapp_number         AS "WhatsApp",
                  m.email                   AS "Email",
                  CASE WHEN m.is_minor THEN 'Yes' ELSE '' END AS "Under 18",
                  m.parent_guardian_phone   AS "Parent/Guardian Phone",
                  m.emergency_contact_name  AS "Emergency Contact",
                  m.emergency_contact_phone AS "Emergency Phone",
                  CASE WHEN p.id IS NOT NULL THEN 'Yes' ELSE '' END AS "Pre-Registered",
                  CASE WHEN a.id = (SELECT MIN(a2.id) FROM attendance a2
                                    WHERE a2.member_id = a.member_id)
                       THEN 'Yes' ELSE '' END AS "First Visit"
           FROM attendance a
           JOIN members m ON m.id = a.member_id
           LEFT JOIN pre_registrations p ON p.member_id = a.member_id AND p.event_id = a.event_id
           WHERE a.event_id = ?
           ORDER BY a.check_in_timestamp ASC, a.id ASC""",
        (event_id,),
    )
    df.insert(0, "#", range(1, len(df) + 1))
    return df


def get_prereg_ledger(event_id: int) -> pd.DataFrame:
    """Pre-registrations for ONE event, with whether each person has arrived."""
    df = query_df(
        """SELECT strftime('%Y-%m-%d %H:%M:%S', p.registered_at) AS "Registered At",
                  m.full_name               AS "Full Name",
                  m.phone_number            AS "Phone",
                  m.whatsapp_number         AS "WhatsApp",
                  m.email                   AS "Email",
                  CASE WHEN m.is_minor THEN 'Yes' ELSE '' END AS "Under 18",
                  m.parent_guardian_phone   AS "Parent/Guardian Phone",
                  m.emergency_contact_name  AS "Emergency Contact",
                  m.emergency_contact_phone AS "Emergency Phone",
                  COALESCE(strftime('%Y-%m-%d %H:%M:%S', a.check_in_timestamp), '') AS "Arrived At"
           FROM pre_registrations p
           JOIN members m ON m.id = p.member_id
           LEFT JOIN attendance a ON a.member_id = p.member_id AND a.event_id = p.event_id
           WHERE p.event_id = ?
           ORDER BY p.registered_at ASC, p.id ASC""",
        (event_id,),
    )
    df.insert(0, "#", range(1, len(df) + 1))
    return df


def overview_stats() -> dict:
    with closing(get_conn()) as conn:
        return dict(conn.execute(
            """SELECT (SELECT COUNT(*) FROM events) AS events,
                      (SELECT COUNT(*) FROM events WHERE is_open = 1) AS open_events,
                      (SELECT COUNT(*) FROM members) AS members,
                      (SELECT COUNT(*) FROM pre_registrations) AS preregs,
                      (SELECT COUNT(*) FROM attendance) AS checkins,
                      (SELECT COUNT(*) FROM attendance
                        WHERE date(check_in_timestamp) = date('now')) AS today"""
        ).fetchone())


def events_summary() -> pd.DataFrame:
    return query_df(
        """SELECT e.id,
                  e.event_name AS "Event", e.event_type AS "Type", e.event_date AS "Date",
                  (SELECT COUNT(*) FROM pre_registrations p WHERE p.event_id = e.id) AS "Pre-Registered",
                  (SELECT COUNT(*) FROM attendance a WHERE a.event_id = e.id) AS "Checked In",
                  (SELECT COUNT(*) FROM pre_registrations p JOIN attendance a
                     ON a.member_id = p.member_id AND a.event_id = p.event_id
                    WHERE p.event_id = e.id) AS "Pre-Registered & Came",
                  (SELECT COUNT(*) FROM attendance a
                    WHERE a.event_id = e.id
                      AND a.id = (SELECT MIN(a2.id) FROM attendance a2 WHERE a2.member_id = a.member_id)
                  ) AS "First-Timers",
                  CASE WHEN e.prereg_open THEN 'On' ELSE 'Off' END AS "Pre-Registration",
                  CASE WHEN e.is_open THEN 'Open' ELSE 'Closed' END AS "Check-In"
           FROM events e
           ORDER BY COALESCE(e.event_date, date(e.created_at)) DESC, e.id DESC"""
    )


# ----- team accounts and activity log -------------------------------------------
def hash_password(password: str) -> str:
    salt = os.urandom(16)
    iterations = 200_000
    digest = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, iterations)
    return f"pbkdf2_sha256${iterations}${salt.hex()}${digest.hex()}"


def verify_password(password: str, stored: str) -> bool:
    try:
        _, iterations, salt, expected = stored.split("$")
        digest = hashlib.pbkdf2_hmac("sha256", password.encode(), bytes.fromhex(salt), int(iterations))
        return hmac.compare_digest(digest.hex(), expected)
    except (ValueError, TypeError):
        return False


def owner_username() -> str:
    return get_setting("ADMIN_USERNAME", DEFAULT_ADMIN_USERNAME).strip() or DEFAULT_ADMIN_USERNAME


def authenticate(username: str, password: str):
    """Returns the signed-in person as a dict, or None."""
    username = (username or "").strip()
    if username.casefold() == owner_username().casefold():
        expected = get_setting("ADMIN_PASSWORD", DEFAULT_ADMIN_PASSWORD)
        if hmac.compare_digest(password.encode(), expected.encode()):
            return {"id": 0, "name": "Main administrator", "username": owner_username(), "role": "Owner"}
        return None
    with closing(get_conn()) as conn, conn:
        row = conn.execute("SELECT * FROM admins WHERE username = ? AND is_active = 1", (username,)).fetchone()
        if row and verify_password(password, row["password_hash"]):
            conn.execute("UPDATE admins SET last_login = CURRENT_TIMESTAMP WHERE id = ?", (row["id"],))
            return {"id": row["id"], "name": row["full_name"], "username": row["username"], "role": row["role"]}
    return None


def create_admin(full_name: str, username: str, password: str, role: str, created_by: str) -> int:
    """Raises sqlite3.IntegrityError if the username is taken."""
    with closing(get_conn()) as conn, conn:
        cur = conn.execute(
            """INSERT INTO admins (full_name, username, password_hash, role, created_by)
               VALUES (?, ?, ?, ?, ?)""",
            (full_name.strip(), username.strip(), hash_password(password), role, created_by),
        )
        return cur.lastrowid


def get_admin(admin_id: int):
    with closing(get_conn()) as conn:
        row = conn.execute("SELECT * FROM admins WHERE id = ?", (admin_id,)).fetchone()
        return dict(row) if row else None


def list_admins() -> pd.DataFrame:
    return query_df(
        """SELECT id, full_name AS "Name", username AS "Username", role AS "Role",
                  CASE WHEN is_active THEN 'Active' ELSE 'Suspended' END AS "Status",
                  COALESCE(created_by, '') AS "Added By",
                  strftime('%Y-%m-%d', created_at) AS "Added",
                  COALESCE(strftime('%Y-%m-%d %H:%M', last_login), 'Never') AS "Last Sign-In"
           FROM admins ORDER BY is_active DESC, full_name COLLATE NOCASE"""
    )


def update_admin(admin_id: int, **fields):
    allowed = {"role", "is_active", "password_hash", "full_name"}
    cols = [c for c in fields if c in allowed]
    with closing(get_conn()) as conn, conn:
        conn.execute(f"UPDATE admins SET {', '.join(c + ' = ?' for c in cols)} WHERE id = ?",
                     [fields[c] for c in cols] + [admin_id])


def delete_admin(admin_id: int):
    with closing(get_conn()) as conn, conn:
        conn.execute("DELETE FROM admins WHERE id = ?", (admin_id,))


def log_action(action: str, actor: str | None = None):
    """Write one line to the activity log. Never lets a logging problem
    interrupt the admin's work."""
    if actor is None:
        user = st.session_state.get("admin_user") or {}
        actor = f"{user.get('name', 'Unknown')} ({user.get('role', '?')})"
    try:
        with closing(get_conn()) as conn, conn:
            conn.execute("INSERT INTO activity_log (actor, action) VALUES (?, ?)", (actor, action))
    except sqlite3.Error:
        pass


def activity_log(limit: int = 500) -> pd.DataFrame:
    return query_df(
        """SELECT strftime('%Y-%m-%d %H:%M:%S', at) AS "When (GMT)", actor AS "Who", action AS "What"
           FROM activity_log ORDER BY id DESC LIMIT ?""",
        (limit,),
    )


# ----- backup / restore --------------------------------------------------------
def make_backup_bytes() -> bytes:
    """A consistent copy of the whole database, safe to take while people check in."""
    fd, path = tempfile.mkstemp(suffix=".db")
    os.close(fd)
    try:
        with closing(get_conn()) as src, closing(sqlite3.connect(path)) as dst:
            src.backup(dst)
        with open(path, "rb") as fh:
            return fh.read()
    finally:
        os.remove(path)


def restore_backup(uploaded_bytes: bytes) -> dict:
    """Check the uploaded file is a genuine Ignite backup, then copy it over
    the live database. Raises ValueError with a friendly message if not."""
    fd, path = tempfile.mkstemp(suffix=".db")
    os.close(fd)
    try:
        with open(path, "wb") as fh:
            fh.write(uploaded_bytes)
        try:
            with closing(sqlite3.connect(path)) as src:
                if src.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
                    raise ValueError("The backup file is damaged.")
                tables = {r[0] for r in src.execute("SELECT name FROM sqlite_master WHERE type='table'")}
                if not {"events", "members", "attendance"} <= tables:
                    raise ValueError("That file isn't an Ignite check-in backup.")
                counts = {t: src.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
                          for t in ("events", "members", "attendance")}
                with closing(get_conn()) as dst:
                    src.backup(dst)
        except sqlite3.DatabaseError:
            raise ValueError("That file isn't a valid database backup.")
    finally:
        os.remove(path)

    with closing(get_conn()) as conn, conn:   # bring an older backup up to date
        migrate(conn)
    get_flyer.clear()
    return counts


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def normalize_phone(raw: str) -> str:
    """Digits only, and +233XXXXXXXXX becomes 0XXXXXXXXX, so one number is
    always stored the same way however it was typed."""
    digits = re.sub(r"\D", "", raw or "")
    if digits.startswith("233") and len(digits) == 12:
        digits = "0" + digits[3:]
    return digits


def is_valid_phone(phone: str) -> bool:
    return 9 <= len(phone) <= 15


def is_valid_email(email: str) -> bool:
    return bool(re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email.strip()))


def fmt_date(iso) -> str:
    if not iso:
        return ""
    try:
        return date.fromisoformat(iso).strftime("%a %d %b %Y")
    except ValueError:
        return str(iso)


def event_label(ev: dict) -> str:
    parts = [ev["event_name"], ev["event_type"]]
    if ev.get("event_date"):
        parts.append(fmt_date(ev["event_date"]))
    return " · ".join(parts)


def public_label(ev: dict) -> str:
    """Event name (and date) for members, without the internal event type."""
    return ev["event_name"] + (f" · {fmt_date(ev['event_date'])}" if ev.get("event_date") else "")


def safe_filename(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]+", "_", text).strip("_") or "event"


def esc(text) -> str:
    return html.escape(str(text or ""))


def to_excel_bytes(df: pd.DataFrame, sheet: str) -> bytes:
    sheet = (safe_filename(sheet) or "Sheet")[:31]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet)
        ws = writer.sheets[sheet]
        for col_cells in ws.columns:   # readable column widths
            width = max(len(str(c.value or "")) for c in col_cells) + 2
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(width, 8), 40)
    return buf.getvalue()


def prepare_flyer(uploaded_file) -> bytes:
    """Check it's really an image, straighten phone photos, flatten
    transparency onto white, shrink it for mobile and save as JPEG."""
    if uploaded_file.size > FLYER_MAX_UPLOAD_MB * 1024 * 1024:
        raise ValueError(f"That file is over {FLYER_MAX_UPLOAD_MB} MB. Please use a smaller image.")
    try:
        img = Image.open(uploaded_file)
        img.load()
    except (UnidentifiedImageError, OSError, Image.DecompressionBombError):
        raise ValueError("That file doesn't look like a valid JPEG or PNG image.")

    img = ImageOps.exif_transpose(img)
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGBA")
        background = Image.new("RGB", img.size, "white")
        background.paste(img, mask=img.getchannel("A"))
        img = background
    else:
        img = img.convert("RGB")
    if img.width > FLYER_MAX_WIDTH:
        img = img.resize((FLYER_MAX_WIDTH, round(img.height * FLYER_MAX_WIDTH / img.width)), Image.LANCZOS)

    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85, optimize=True, progressive=True)
    return buf.getvalue()


def detect_base_url() -> str:
    """Where links and QR codes should point: APP_BASE_URL from secrets if
    set, otherwise the address the admin is using right now."""
    configured = get_setting("APP_BASE_URL").strip().rstrip("/")
    if configured:
        return configured
    try:
        parts = urlsplit(st.context.url or "")
        if parts.scheme and parts.netloc and "localhost" not in parts.netloc:
            return f"{parts.scheme}://{parts.netloc}"
    except Exception:
        pass
    return ""


def checkin_link(base: str, event_id: int) -> str:
    return f"{base}/?event={event_id}"


def register_link(base: str, event_id: int) -> str:
    return f"{base}/?event={event_id}&mode=register"


# ----- QR codes and printable posters -------------------------------------------
def make_qr_image(data: str) -> Image.Image:
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=10, border=4)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img = img.get_image() if hasattr(img, "get_image") else img
    return img.convert("RGB")


def to_png(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def load_font(size: int, bold: bool = False):
    names = (["DejaVuSans-Bold.ttf", "LiberationSans-Bold.ttf", "Arial Bold.ttf", "arialbd.ttf"]
             if bold else ["DejaVuSans.ttf", "LiberationSans-Regular.ttf", "Arial.ttf", "arial.ttf"])
    for name in names:
        for folder in ("", "/usr/share/fonts/truetype/dejavu/", "/usr/share/fonts/truetype/liberation/"):
            try:
                return ImageFont.truetype(folder + name, size)
            except OSError:
                continue
    try:
        return ImageFont.load_default(size=size)   # Pillow 10.1+
    except TypeError:
        return ImageFont.load_default()


def _wrap(draw, text, font, max_width):
    lines, current = [], ""
    for word in text.split():
        trial = f"{current} {word}".strip()
        if draw.textlength(trial, font=font) <= max_width:
            current = trial
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


def _centre(draw, width, y, text, font, fill):
    draw.text(((width - draw.textlength(text, font=font)) / 2, y), text, font=font, fill=fill)


@st.cache_data(show_spinner=False)
def make_poster(link: str, name: str, event_type: str, date_text: str, venue: str,
                heading: str, call_to_action: str) -> bytes:
    """A4-sized (150 dpi) poster with the event details and a big QR code."""
    W, H = 1240, 1754
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)

    d.rectangle([0, 0, W, 290], fill=BRAND)
    d.rectangle([0, 290, W, 306], fill=ACCENT)
    _centre(d, W, 80, "IGNITE PRAYER NETWORK", load_font(66, True), "white")
    _centre(d, W, 180, heading, load_font(40), "white")

    y = 370
    title_font = load_font(74, True)
    for line in _wrap(d, name, title_font, W - 160)[:3]:
        _centre(d, W, y, line, title_font, "#1d1d1f")
        y += 90
    meta = " · ".join(p for p in (event_type, date_text, venue) if p)
    meta_font = load_font(36)
    for line in _wrap(d, meta, meta_font, W - 200)[:2]:
        _centre(d, W, y + 6, line, meta_font, "#555555")
        y += 50

    size = 740
    y += 60
    x = (W - size) // 2
    d.rounded_rectangle([x - 26, y - 26, x + size + 26, y + size + 26], radius=30, outline=BRAND, width=6)
    img.paste(make_qr_image(link).resize((size, size), Image.NEAREST), (x, y))
    y += size + 80

    _centre(d, W, y, call_to_action, load_font(44, True), BRAND_DARK)
    _centre(d, W, y + 70, "Point your camera at the code, then tap the link that appears.",
            load_font(28), "#666666")

    d.rectangle([0, H - 100, W, H], fill="#1d1d1f")
    _centre(d, W, H - 66, urlsplit(link).netloc, load_font(26), "#ffffff")
    return to_png(img)


# ---------------------------------------------------------------------------
# Look and feel
# ---------------------------------------------------------------------------
def inject_css(public: bool):
    width = "720px" if public else "1200px"
    st.markdown(
        f"""
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&display=swap');
          .stApp, .stMarkdown p, label, input, textarea, button p, .stTabs button p {{
              font-family: 'Poppins', system-ui, -apple-system, 'Segoe UI', sans-serif; }}
          #MainMenu, footer, [data-testid="stToolbar"], [data-testid="stDecoration"] {{visibility:hidden;}}
          header[data-testid="stHeader"] {{background:transparent;}}
          .block-container {{max-width:{width}; padding-top:1.2rem; padding-bottom:3rem;}}

          /* Brand header */
          .ig-hero {{background:linear-gradient(135deg, {BRAND} 0%, #EF7A2E 55%, {ACCENT} 100%);
                     color:#fff; border-radius:20px; padding:1.4rem 1.2rem; text-align:center;
                     box-shadow:0 8px 24px rgba(228,87,46,.25); margin-bottom:1rem;}}
          .ig-hero .ig-flame {{font-size:2.1rem; line-height:1;}}
          .ig-hero .ig-title {{font-size:1.6rem; font-weight:700; letter-spacing:.2px; margin-top:.3rem;}}
          .ig-hero .ig-sub {{opacity:.92; font-size:.95rem; margin-top:.15rem;}}
          .ig-hero.ig-compact {{display:flex; align-items:center; gap:.8rem; text-align:left; padding:1rem 1.3rem;}}
          .ig-hero.ig-compact .ig-title {{margin-top:0; font-size:1.35rem;}}

          /* Event card */
          .ig-event {{border:1px solid rgba(228,87,46,.28); background:rgba(228,87,46,.07);
                      border-radius:16px; padding:1rem 1.1rem; margin:.2rem 0 1rem;}}
          .ig-event-name {{font-size:1.25rem; font-weight:700; margin-top:.35rem;}}
          .ig-event-meta {{opacity:.8; font-size:.92rem; margin-top:.2rem;}}
          .ig-badge {{display:inline-block; color:#fff; font-size:.72rem; font-weight:600;
                      padding:.18rem .6rem; border-radius:999px; letter-spacing:.3px; text-transform:uppercase;}}

          /* Flyer banner: full width, never taller than about half the phone screen */
          .st-key-flyer_banner img {{width:100%; max-height:55vh; object-fit:contain;
              border-radius:16px; box-shadow:0 6px 20px rgba(0,0,0,.18);}}

          /* Welcome / notice screens */
          .ig-success, .ig-notice {{text-align:center; padding:3.2rem 1.2rem; border-radius:22px; color:#fff;
                                    margin-top:1.5rem;}}
          .ig-success {{background:linear-gradient(135deg,#1B998B,#14746A);}}
          .ig-notice  {{background:linear-gradient(135deg,{ACCENT},#D98E04); color:#1d1d1f;}}
          .ig-success .ig-check, .ig-notice .ig-check {{font-size:3.2rem; line-height:1;}}
          .ig-success h2, .ig-notice h2 {{color:inherit; font-size:1.7rem; margin:.6rem 0 .3rem;}}

          /* Buttons, forms, metrics */
          div.stButton > button, div.stFormSubmitButton > button, div.stDownloadButton > button,
          div.stLinkButton > a {{border-radius:12px; font-weight:600; min-height:2.9rem;}}
          [data-testid="stBaseButton-primary"], [data-testid="stBaseButton-primaryFormSubmit"] {{
              background:{BRAND}; border-color:{BRAND};}}
          [data-testid="stBaseButton-primary"]:hover, [data-testid="stBaseButton-primaryFormSubmit"]:hover {{
              background:{BRAND_DARK}; border-color:{BRAND_DARK};}}
          [data-testid="stForm"] {{border-radius:16px;}}
          [data-testid="stMetric"] {{border:1px solid rgba(128,128,128,.25); border-radius:14px; padding:.8rem 1rem;}}

          .st-key-admin_entry .stButton {{display:flex; justify-content:center;}}
          [data-baseweb="tab-highlight"] {{background-color:{BRAND};}}
          .stTabs [aria-selected="true"] p {{color:{BRAND};}}
          .ig-flame-svg {{width:2rem; height:2.4rem; color:#fff; filter:drop-shadow(0 2px 6px rgba(0,0,0,.2));}}
          .ig-compact .ig-flame-svg {{width:1.7rem; height:2rem;}}
          .ig-ico {{width:1rem; height:1rem; vertical-align:-2px; margin-right:.35rem; opacity:.85;}}
          .ig-event-meta span {{display:inline-block; margin-right:1.1rem; margin-top:.15rem;}}
          .ig-lead {{font-size:1.05rem; font-weight:600; margin:.2rem 0 .3rem;}}
          .ig-note {{font-size:.9rem; opacity:.8; margin-bottom:.6rem; line-height:1.5;}}
          .ig-form-note {{font-size:.88rem; opacity:.75; margin-bottom:.2rem; line-height:1.5;}}
          .ig-section {{font-size:.78rem; font-weight:700; letter-spacing:.08em; text-transform:uppercase;
                        color:{BRAND}; margin:.9rem 0 .1rem; padding-top:.6rem;
                        border-top:1px solid rgba(128,128,128,.2);}}

          /* Confirmation ("Submitted · thank you") screen */
          .ig-done {{text-align:center; padding:2.2rem 1.4rem 2rem; border-radius:24px; color:#fff; margin:.6rem 0 1rem;
                     background:linear-gradient(160deg,#1B998B 0%,#127A6F 100%); box-shadow:0 10px 30px rgba(18,122,111,.35);}}
          .ig-done-info {{background:linear-gradient(160deg,#3D5A80 0%,#2B4162 100%); box-shadow:0 10px 30px rgba(43,65,98,.35);}}
          .ig-done-brand {{display:flex; align-items:center; justify-content:center; gap:.45rem;
                           font-weight:600; font-size:.9rem; opacity:.9; margin-bottom:1.3rem;}}
          .ig-done-brand .ig-flame-svg {{width:1.1rem; height:1.35rem; filter:none;}}
          .ig-done-ico {{width:4.2rem; height:4.2rem; color:#fff; display:block; margin:0 auto .9rem;}}
          .ig-done-kicker {{font-size:.78rem; font-weight:700; letter-spacing:.14em; text-transform:uppercase; opacity:.85;}}
          .ig-done h2 {{color:#fff; font-size:1.75rem; margin:.35rem 0 .5rem; padding:0;}}
          .ig-done p {{font-size:1.02rem; line-height:1.55; margin:0 auto; max-width:30rem; opacity:.95;}}
          .ig-done-meta {{margin-top:1.3rem; font-size:.82rem; opacity:.8; padding-top:1rem;
                          border-top:1px solid rgba(255,255,255,.25);}}
          .st-key-ig_scroll {{display:none;}}
          .ig-footer {{text-align:center; opacity:.55; font-size:.8rem; margin-top:.4rem;}}
        </style>
        """,
        unsafe_allow_html=True,
    )


# ----- small inline icons (crisp on every phone, no emoji differences) ----------
ICON_FLAME = ('<svg class="ig-flame-svg" viewBox="0 0 24 30" aria-hidden="true">'
              '<path d="M12 1c.6 5.2-6.5 8.4-6.5 16.2A6.5 6.5 0 0 0 12 23.7a6.5 6.5 0 0 0 6.5-6.5'
              'c0-3.6-1.9-6-3.1-7.6-.2 2.6-1.3 4-2.8 4.6C13.4 10.7 13.3 5.6 12 1z" fill="currentColor"/></svg>')
ICON_CAL = ('<svg class="ig-ico" viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M7 2h2v2h6V2h2v2h2'
            'a2 2 0 0 1 2 2v13a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2V2zm12 8H5v9h14v-9z"/></svg>')
ICON_PIN = ('<svg class="ig-ico" viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M12 2a7 7 0 0 1 7 7'
            'c0 5.2-7 13-7 13S5 14.2 5 9a7 7 0 0 1 7-7zm0 4.5A2.5 2.5 0 1 0 12 11.5 2.5 2.5 0 0 0 12 6.5z"/></svg>')
ICON_CHECK = ('<svg class="ig-done-ico" viewBox="0 0 52 52" aria-hidden="true"><circle cx="26" cy="26" r="25" '
              'fill="none" stroke="currentColor" stroke-width="2.5"/><path d="M15 27l7 7 15-16" fill="none" '
              'stroke="currentColor" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/></svg>')
ICON_INFO = ('<svg class="ig-done-ico" viewBox="0 0 52 52" aria-hidden="true"><circle cx="26" cy="26" r="25" '
             'fill="none" stroke="currentColor" stroke-width="2.5"/><path d="M26 23v13M26 16v.5" fill="none" '
             'stroke="currentColor" stroke-width="4" stroke-linecap="round"/></svg>')


def brand_header(subtitle: str = "", compact: bool = False):
    cls = "ig-hero ig-compact" if compact else "ig-hero"
    inner = (f'<div><div class="ig-title">{APP_NAME}</div><div class="ig-sub">{esc(subtitle)}</div></div>'
             if compact else
             f'<div class="ig-title">{APP_NAME}</div><div class="ig-sub">{esc(subtitle)}</div>')
    st.markdown(f'<div class="{cls}"><div class="ig-flame">{ICON_FLAME}</div>{inner}</div>', unsafe_allow_html=True)


def event_card(ev: dict, public: bool = False):
    """Event name, programme, date and venue. Members see the friendly
    programme name (e.g. "Special Meeting" instead of "Impromptu")."""
    colour = TYPE_COLOURS.get(ev["event_type"], BRAND)
    type_name = PUBLIC_TYPE_NAMES.get(ev["event_type"], ev["event_type"]) if public else ev["event_type"]
    meta = "".join(
        p for p in (
            f'<span>{ICON_CAL}{esc(fmt_date(ev.get("event_date")))}</span>' if ev.get("event_date") else "",
            f'<span>{ICON_PIN}{esc(ev.get("venue"))}</span>' if ev.get("venue") else "",
        ) if p
    )
    st.markdown(
        f"""<div class="ig-event">
              <span class="ig-badge" style="background:{colour}">{esc(type_name)}</span>
              <div class="ig-event-name">{esc(ev['event_name'])}</div>
              {f'<div class="ig-event-meta">{meta}</div>' if meta else ''}
            </div>""",
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Public pages: shared pieces
# ---------------------------------------------------------------------------
def confirm_and_reset(kind: str, title: str, message: str, ev: dict):
    """Record what to show on the confirmation screen, throw away the form
    (a new form counter means brand-new, empty widgets) and rerun."""
    st.session_state.confirmation = {
        "kind": kind, "title": title, "message": message,
        "event": ev["event_name"], "time": datetime.now(timezone.utc).strftime("%H:%M"),
        "shown_at": None,
    }
    st.session_state.form_nonce += 1
    st.rerun()


def scroll_to_top():
    """Bring the page back to the top so the confirmation is the first thing seen."""
    script = """<script>
        const d = window.parent.document;
        for (const sel of ['[data-testid="stMain"]', '[data-testid="stAppViewContainer"]', 'section.main']) {
            const el = d.querySelector(sel); if (el) el.scrollTo({top: 0, behavior: 'instant'});
        }
        window.parent.scrollTo(0, 0);
        </script>"""
    with st.container(key="ig_scroll"):          # hidden by CSS; the script still runs
        if hasattr(st, "iframe"):
            st.iframe(script, height=1)
        else:                                    # older Streamlit versions
            st.components.v1.html(script, height=0)


def show_confirmation() -> bool:
    """Privacy screen: after a submission the whole page is replaced by this
    card. No personal details are shown. It resets itself after a few
    seconds (or when "Done" is tapped), ready for the next person."""
    c = st.session_state.get("confirmation")
    if not c:
        return False
    if c["shown_at"] is None:
        c["shown_at"] = time.time()

    css, icon = ("ig-done", ICON_CHECK) if c["kind"] == "success" else ("ig-done ig-done-info", ICON_INFO)
    st.markdown(
        f"""<div class="{css}">
              <div class="ig-done-brand">{ICON_FLAME}<span>{APP_NAME}</span></div>
              {icon}
              <div class="ig-done-kicker">Submitted · thank you</div>
              <h2>{esc(c['title'])}</h2>
              <p>{esc(c['message'])}</p>
              <div class="ig-done-meta">{esc(c['event'])} &nbsp;·&nbsp; recorded at {c['time']} GMT</div>
            </div>""",
        unsafe_allow_html=True,
    )
    scroll_to_top()
    if st.button("Done", type="primary", width="stretch", key="confirm_done"):
        st.session_state.pop("confirmation", None)
        st.rerun()

    @st.fragment(run_every=1)
    def countdown():
        left = CONFIRM_SECONDS - int(time.time() - c["shown_at"])
        if left <= 0:
            st.session_state.pop("confirmation", None)
            st.rerun(scope="app")
        st.markdown(f'<div class="ig-footer">This screen will reset in {left} second{"s" if left != 1 else ""}.</div>',
                    unsafe_allow_html=True)

    countdown()
    return True


def show_errors(errors: list[str]):
    st.error("**Please check the following:**\n\n" + "\n".join(f"- {e}" for e in errors))


def phone_lookup_form(prefix: str, caption: str, button_label: str, not_found_hint: str):
    """Returns a member id after a valid submit, otherwise None."""
    nonce = st.session_state.form_nonce
    with st.form(key=f"{prefix}_quick_{nonce}"):
        st.markdown(f'<div class="ig-form-note">{esc(caption)}</div>', unsafe_allow_html=True)
        phone_raw = st.text_input("Phone number", placeholder="e.g. 024 123 4567",
                                  autocomplete="off", key=f"{prefix}_q_phone_{nonce}")
        submitted = st.form_submit_button(button_label, type="primary", width="stretch")
    if not submitted:
        return None
    phone = normalize_phone(phone_raw)
    if not is_valid_phone(phone):
        show_errors(["Please enter a valid phone number, e.g. 024 123 4567."])
        return None
    member_id = find_member_id(phone)
    if member_id is None:
        show_errors([f"We couldn't find {phone_raw.strip()} in our records. {not_found_hint}"])
    return member_id


def new_member_form(prefix: str, button_label: str):
    """Full registration form. Returns (member_id, is_new) after a valid
    submit, otherwise None. Existing numbers are reused, never overwritten."""
    nonce = st.session_state.form_nonce
    k = lambda name: f"{prefix}_{name}_{nonce}"
    with st.form(key=k("register")):
        st.markdown('<div class="ig-form-note">Fill this in once. Next time you\'ll only need your phone number. '
                    'Fields marked * are required.</div>', unsafe_allow_html=True)
        st.markdown('<div class="ig-section">Your details</div>', unsafe_allow_html=True)
        full_name = st.text_input("Full name *", placeholder="e.g. Ama Serwaa Mensah", autocomplete="off", key=k("name"))
        c1, c2 = st.columns(2)
        phone_raw = c1.text_input("Phone number *", placeholder="e.g. 024 123 4567", autocomplete="off", key=k("phone"))
        whatsapp_raw = c2.text_input("WhatsApp number", placeholder="Leave blank if the same",
                                     autocomplete="off", key=k("wa"))
        email = st.text_input("Email address", placeholder="Optional", autocomplete="off", key=k("email"))

        st.markdown('<div class="ig-section">Emergency contact</div>', unsafe_allow_html=True)
        c3, c4 = st.columns(2)
        ec_name = c3.text_input("Contact name *", placeholder="Who should we call?", autocomplete="off", key=k("ecn"))
        ec_phone_raw = c4.text_input("Contact phone *", placeholder="e.g. 020 123 4567",
                                     autocomplete="off", key=k("ecp"))

        st.markdown('<div class="ig-section">Under 18?</div>', unsafe_allow_html=True)
        is_minor = st.checkbox("I am under 18 years old", key=k("minor"))
        parent_raw = st.text_input("Parent/guardian phone", placeholder="Required if you are under 18",
                                   autocomplete="off", key=k("par"))

        st.markdown('<div class="ig-section">Consent</div>', unsafe_allow_html=True)
        consent = st.checkbox(
            f"I agree that {APP_NAME} may keep these details for attendance records and may contact "
            "my emergency contact or parent/guardian if needed. *",
            key=k("consent"),
        )
        submitted = st.form_submit_button(button_label, type="primary", width="stretch")
    if not submitted:
        return None

    phone = normalize_phone(phone_raw)
    ec_phone = normalize_phone(ec_phone_raw)
    parent = normalize_phone(parent_raw)
    whatsapp = normalize_phone(whatsapp_raw) or phone

    errors = []
    if len(full_name.strip()) < 2:
        errors.append("Enter your full name.")
    if not is_valid_phone(phone):
        errors.append("Enter a valid phone number.")
    if whatsapp_raw.strip() and not is_valid_phone(whatsapp):
        errors.append("The WhatsApp number doesn't look right.")
    if email.strip() and not is_valid_email(email):
        errors.append("The email address doesn't look right.")
    if not ec_name.strip():
        errors.append("Enter an emergency contact name.")
    if not is_valid_phone(ec_phone):
        errors.append("Enter a valid emergency contact phone number.")
    if is_minor and not parent:
        errors.append("A parent/guardian phone number is required for anyone under 18.")
    if parent and not is_valid_phone(parent):
        errors.append("The parent/guardian phone number doesn't look right.")
    if not consent:
        errors.append("Tick the consent box so we can keep your details.")
    if errors:
        show_errors(errors)
        return None

    member_id = find_member_id(phone)
    if member_id is not None:
        return member_id, False
    member_id = create_member({
        "full_name": full_name.strip(),
        "phone_number": phone,
        "whatsapp_number": whatsapp,
        "email": email.strip() or None,
        "emergency_contact_name": ec_name.strip(),
        "emergency_contact_phone": ec_phone,
        "parent_guardian_phone": parent or None,
        "is_minor": int(is_minor),
    })
    return member_id, True


def page_top(subtitle: str):
    """Flyer slot at the very top, then the brand header. Returns the slot."""
    banner = st.container(key="flyer_banner")
    brand_header(subtitle)
    return banner


def show_event_header(banner, ev: dict):
    flyer = get_flyer(ev["id"])
    if flyer:
        banner.image(flyer, width="stretch")
    event_card(ev, public=True)


def public_footer():
    # No admin button here on purpose: admins open the private link
    # https://<your-app>.streamlit.app/?view=admin (bookmark it).
    st.write("")
    st.markdown(f'<div class="ig-footer">© {datetime.now().year} {APP_NAME}</div>', unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Public page: day-of check-in
# ---------------------------------------------------------------------------
def resolve_checkin_event():
    """From ?event=<id>, or let the member choose among open events."""
    raw = st.query_params.get("event")
    if raw and str(raw).isdigit():
        ev = get_event(int(raw))
        if ev:
            return ev
        st.warning("That check-in link is no longer valid. Please choose your programme below.")
    events = get_events(open_only=True)
    if not events:
        st.info("There are no programmes open for check-in right now. Please check back later.")
        return None
    if len(events) == 1:
        return events[0]
    return st.selectbox("Which programme are you attending?", events, format_func=public_label)


def checkin_page():
    if show_confirmation():
        return
    banner = page_top("Arrival Check-In")
    ev = resolve_checkin_event()
    if ev is None:
        return
    show_event_header(banner, ev)

    if not ev["is_open"]:
        st.warning("Check-in for this programme is closed. Please speak to an usher if you need help.")
        return

    st.markdown('<div class="ig-lead">Welcome! Please confirm your arrival below.</div>', unsafe_allow_html=True)
    tab_returning, tab_new = st.tabs([
        ":material/how_to_reg: " + ("Registered / attended before" if ev["prereg_open"] else "Attended Ignite before"),
        ":material/person_add: First time with Ignite"])
    already = ("You're already checked in",
               "Your arrival for this programme was recorded earlier, so there's nothing more to do. Enjoy the programme!")
    with tab_returning:
        member_id = phone_lookup_form(
            "chk", ("Registered online for this programme, or attended Ignite before? "
                    if ev["prereg_open"] else "Been with Ignite before? ") + "Just enter your phone number.",
            "Check In",
            "If this is your first time with Ignite, please use the “First time with Ignite” tab.")
        if member_id:
            if record_check_in(member_id, ev["id"]):
                msg = ("You registered ahead of time, and your arrival is now confirmed. Enjoy the programme!"
                       if is_pre_registered(member_id, ev["id"])
                       else "Your arrival has been recorded. Enjoy the programme!")
                confirm_and_reset("success", "Welcome to Ignite Network!", msg, ev)
            else:
                confirm_and_reset("info", *already, ev)
    with tab_new:
        result = new_member_form("chk", "Submit & Check In")
        if result:
            member_id, is_new = result
            msg = ("Your details have been saved and your arrival is recorded. We're glad you're here!" if is_new
                   else "This phone number was already on our records, so we've simply checked you in.")
            if record_check_in(member_id, ev["id"]):
                confirm_and_reset("success", "Welcome to Ignite Network!", msg, ev)
            else:
                confirm_and_reset("info", *already, ev)


# ---------------------------------------------------------------------------
# Public page: pre-registration (shared before the day)
# ---------------------------------------------------------------------------
def register_page():
    if show_confirmation():
        return
    banner = page_top("Pre-Registration")

    raw = st.query_params.get("event")
    ev = get_event(int(raw)) if raw and str(raw).isdigit() else None
    if ev is None:
        st.error("This registration link isn't valid. Please ask the organisers for the correct link.")
        return
    show_event_header(banner, ev)

    if not ev["prereg_open"]:
        st.info("This programme doesn't need pre-registration. Just come along on the day and scan "
                "the QR code at the entrance to check in.", icon=":material/info:")
        return

    st.markdown(
        '<div class="ig-lead">Let us know you\'re coming so we can prepare for you.</div>'
        '<div class="ig-note">This reserves your place; it isn\'t your check-in. On the day, scan the QR code '
        'at the entrance and enter your phone number to confirm you\'ve arrived.</div>',
        unsafe_allow_html=True,
    )
    tab_known, tab_new = st.tabs([":material/how_to_reg: I've attended Ignite before",
                                  ":material/person_add: I'm new to Ignite"])
    done_title = "You're registered!"
    done_msg = "Your place is reserved. On the day, scan the QR code at the entrance and enter your phone number."
    already = ("You're already registered", "We already have your registration for this programme. See you there!")
    with tab_known:
        member_id = phone_lookup_form(
            "reg", "Enter the phone number you gave us before. We already have your other details.",
            "Register",
            "If you haven't been with us before, please use the “I'm new to Ignite” tab.")
        if member_id:
            if record_pre_registration(member_id, ev["id"]):
                confirm_and_reset("success", done_title, done_msg, ev)
            else:
                confirm_and_reset("info", *already, ev)
    with tab_new:
        result = new_member_form("reg", "Submit Registration")
        if result:
            member_id, _ = result
            if record_pre_registration(member_id, ev["id"]):
                confirm_and_reset("success", done_title, done_msg, ev)
            else:
                confirm_and_reset("info", *already, ev)


# ---------------------------------------------------------------------------
# Admin: access
# ---------------------------------------------------------------------------
def current_admin() -> dict:
    return st.session_state.get("admin_user") or {}


def end_admin_session():
    for key in ("admin_expires", "admin_user"):
        st.session_state.pop(key, None)


def admin_is_authenticated() -> bool:
    """Valid session, not idle too long, and (for team members) the account
    is still active. Role changes made by another admin apply straight away."""
    user = st.session_state.get("admin_user")
    if not user or st.session_state.get("admin_expires", 0) <= time.time():
        end_admin_session()
        return False
    if user["id"]:
        fresh = get_admin(user["id"])
        if not fresh or not fresh["is_active"]:
            end_admin_session()
            return False
        user["role"], user["name"] = fresh["role"], fresh["full_name"]
    st.session_state.admin_expires = time.time() + ADMIN_SESSION_MINUTES * 60
    return True


def leave_admin():
    if current_admin():
        log_action("Signed out")
    end_admin_session()
    st.session_state.show_admin = False
    if "view" in st.query_params:
        del st.query_params["view"]
    st.rerun()


def admin_login():
    brand_header("Admin Portal")
    locked_until = st.session_state.get("locked_until", 0)
    if locked_until > time.time():
        st.error(f"Too many attempts. Try again in {int(locked_until - time.time())} seconds.")
        return
    with st.form("admin_login", clear_on_submit=True):
        st.markdown("**Sign in to manage events and attendance**")
        username = st.text_input("Username", autocomplete="username")
        pw = st.text_input("Password", type="password", autocomplete="current-password")
        go = st.form_submit_button("Sign in", type="primary", width="stretch")
    if go:
        user = authenticate(username, pw)
        if user:
            st.session_state.admin_user = user
            st.session_state.admin_expires = time.time() + ADMIN_SESSION_MINUTES * 60
            st.session_state.failed_logins = 0
            log_action("Signed in")
            st.rerun()
        else:
            time.sleep(1)
            log_action(f"Failed sign-in attempt for username “{username.strip()[:40]}”", actor="Unknown")
            st.session_state.failed_logins = st.session_state.get("failed_logins", 0) + 1
            if st.session_state.failed_logins >= MAX_LOGIN_ATTEMPTS:
                st.session_state.locked_until = time.time() + LOCKOUT_SECONDS
                st.session_state.failed_logins = 0
            st.error("Incorrect username or password.")


def notify(message: str):
    """Show a confirmation at the top of the admin page after a rerun,
    and record the action in the activity log."""
    st.session_state.admin_notice = message
    log_action(message)


def event_picker(label: str, key: str, events: list[dict]) -> dict:
    """Select an event by id (ids survive edits and deletions cleanly)."""
    ids = [e["id"] for e in events]
    by_id = {e["id"]: e for e in events}
    if st.session_state.get(key) not in ids:
        st.session_state.pop(key, None)
    chosen = st.selectbox(label, ids, format_func=lambda i: event_label(by_id[i]), key=key)
    return by_id[chosen]


def download_pair(df: pd.DataFrame, stem: str, sheet: str, key: str):
    c1, c2 = st.columns(2)
    c1.download_button("Export CSV", df.to_csv(index=False).encode("utf-8"), file_name=f"{stem}.csv",
                       mime="text/csv", width="stretch", key=f"{key}_csv")
    c2.download_button("Export Excel", to_excel_bytes(df, sheet), file_name=f"{stem}.xlsx",
                       mime=XLSX_MIME, width="stretch", key=f"{key}_xlsx")


def filter_people(df: pd.DataFrame, search: str) -> pd.DataFrame:
    if not search:
        return df
    digits = re.sub(r"\D", "", search)
    mask = df["Full Name"].str.contains(search, case=False, na=False, regex=False)
    if digits:
        mask |= df["Phone"].str.contains(digits, na=False, regex=False)
    return df[mask]


# ---------------------------------------------------------------------------
# Admin: tabs
# ---------------------------------------------------------------------------
def overview_tab():
    s = overview_stats()
    c = st.columns(5)
    c[0].metric("Events", s["events"], help=f"{s['open_events']} currently open for check-in")
    c[1].metric("Members", s["members"])
    c[2].metric("Pre-registrations", s["preregs"])
    c[3].metric("Total check-ins", s["checkins"])
    c[4].metric("Check-ins today", s["today"])

    st.info("Free hosting can wipe the database when the app restarts or updates. "
            "Download a backup from the **Backup** tab after every event.", icon=":material/backup:")

    summary = events_summary()
    if summary.empty:
        st.caption("No events yet. Create your first one in the **Events** tab.")
        return
    st.subheader("Registrations and attendance by event")
    chart = summary.head(10).copy()
    chart["Label"] = chart["Event"] + " (#" + chart["id"].astype(str) + ")"
    st.bar_chart(chart.set_index("Label")[["Pre-Registered", "Checked In"]], horizontal=True,
                 stack=False, color=[ACCENT, BRAND], x_label="People", y_label="")
    summary["Date"] = summary["Date"].map(fmt_date)
    st.dataframe(summary.drop(columns=["id"]), hide_index=True, width="stretch")


def create_event_section(has_events: bool):
    with st.expander(":material/add_circle: Create a new event", expanded=not has_events):
        with st.form("new_event", clear_on_submit=True):
            c1, c2 = st.columns(2)
            ev_type = c1.selectbox("Event type", EVENT_TYPES)
            ev_name = c2.text_input("Event name", placeholder=f"e.g. Asteri {datetime.now().year + 1}")
            c3, c4 = st.columns(2)
            ev_date = c3.date_input("Date (optional)", value=None, format="DD/MM/YYYY")
            venue = c4.text_input("Venue (optional)")
            flyer_file = st.file_uploader("Programme flyer (optional, JPEG or PNG)", type=["jpg", "jpeg", "png"])
            c5, c6 = st.columns(2)
            prereg_open = c5.toggle("Use pre-registration", value=True,
                                    help="Turn off for attendance-only events. Members then simply check in on the day.")
            is_open = c6.toggle("Open day-of check-in", value=True)
            create = st.form_submit_button("Create event", type="primary")
        if create:
            name = ev_name.strip() or f"{ev_type} {datetime.now():%Y-%m-%d}"
            try:
                flyer = prepare_flyer(flyer_file) if flyer_file is not None else None
            except ValueError as err:
                st.error(f"{err} The event was not created.")
                return
            new_id = create_event(name, ev_type, ev_date.isoformat() if ev_date else None,
                                  venue, flyer, is_open, prereg_open)
            st.session_state.manage_event = new_id
            notify(f"Created “{name}”. " + ("Its registration link and check-in QR code are ready below."
                                             if prereg_open else "It's attendance-only; its check-in QR code is ready below."))
            st.rerun()


def base_url_input() -> str | None:
    if not st.session_state.get("base_url_input"):
        st.session_state.base_url_input = detect_base_url()
    base = st.text_input(
        "Public app URL", key="base_url_input",
        help="Filled in automatically with this app's address. Only change it if the app moves.",
    ).strip().rstrip("/")
    if not re.fullmatch(r"https?://[^\s/]+\.[^\s/]+", base) or "your-app" in base:
        st.error("Enter this app's real web address above (for example "
                 "https://ignite-checkin-xxxx.streamlit.app) so links and QR codes work.")
        return None
    return base


def poster_block(ev: dict, link: str, heading: str, cta: str, stem: str, key: str):
    poster = make_poster(link, ev["event_name"], ev["event_type"], fmt_date(ev["event_date"]),
                         ev["venue"] or "", heading, cta)
    st.image(poster, caption="Printable poster (A4)", width="stretch")
    st.download_button("Download poster (A4 PNG)", poster, file_name=f"{stem}.png",
                       mime="image/png", width="stretch", key=f"{key}_poster")
    st.download_button("Download QR code only", to_png(make_qr_image(link)),
                       file_name=f"{stem}_QR.png", mime="image/png", width="stretch", key=f"{key}_qr")


def prereg_section(ev: dict, base):
    c1, c2 = st.columns([3, 1])
    if ev["prereg_open"]:
        c1.markdown(f":green[●] **Pre-registration is on** · {ev['preregs']} registered so far")
    else:
        c1.markdown(":gray[●] **Pre-registration is off.** This is an attendance-only event: "
                    "members simply check in on the day.")
    if c2.button("Turn off pre-registration" if ev["prereg_open"] else "Turn on pre-registration",
                 key=f"toggle_prereg_{ev['id']}", width="stretch"):
        set_event_flag(ev["id"], "prereg_open", not ev["prereg_open"])
        notify(f"Pre-registration turned {'off' if ev['prereg_open'] else 'on'} for “{ev['event_name']}”.")
        st.rerun()
    if not ev["prereg_open"]:
        if ev["preregs"]:
            st.caption(f"The {ev['preregs']} registration(s) made earlier are kept in the Pre-registrations tab.")
        return
    if not base:
        st.warning("Set the app address above to get the registration link.")
        return

    link = register_link(base, ev["id"])
    when = f" on {fmt_date(ev['event_date'])}" if ev["event_date"] else ""
    where = f" at {ev['venue']}" if ev["venue"] else ""
    message = (f"🔥 {ev['event_name']}{when}{where}\n\n"
               f"Register ahead of time here: {link}\n\n"
               f"On the day, just scan the QR code at the entrance and enter your phone number. "
               f"See you there! — {APP_NAME}")

    left, right = st.columns([3, 2])
    with left:
        st.markdown("**Registration link** (share this before the day)")
        st.code(link, language=None)
        st.markdown("**Ready-made message** (tap the copy icon, then paste into WhatsApp)")
        st.code(message, language=None, wrap_lines=True)
        st.link_button("Share on WhatsApp", f"https://wa.me/?text={quote(message)}",
                       type="primary", width="stretch")
    with right:
        poster_block(ev, link, "Register Now", "Scan to register for this event",
                     f"Register_{safe_filename(ev['event_name'])}", f"reg_{ev['id']}")


def checkin_qr_section(ev: dict, base):
    status = ":green[●] **Check-in is open**" if ev["is_open"] else ":red[●] **Check-in is closed**"
    c1, c2 = st.columns([3, 1])
    c1.markdown(f"{status} · {ev['attendees']} checked in")
    if c2.button("Close check-in" if ev["is_open"] else "Reopen check-in",
                 key=f"toggle_open_{ev['id']}", width="stretch"):
        set_event_flag(ev["id"], "is_open", not ev["is_open"])
        notify(f"Check-in {'closed' if ev['is_open'] else 'reopened'} for “{ev['event_name']}”.")
        st.rerun()
    if not base:
        st.warning("Set the app address above to get the check-in QR code.")
        return

    link = checkin_link(base, ev["id"])
    left, right = st.columns([3, 2])
    with left:
        st.markdown("**Check-in link** (for the entrance on the day)")
        st.code(link, language=None)
        st.caption("Print the poster and place it at the entrance. Scan it with your own phone "
                   "before printing to be sure it opens this event.")
    with right:
        poster_block(ev, link, "Event Check-In", "Scan with your phone camera to check in",
                     f"CheckIn_{safe_filename(ev['event_name'])}", f"chk_{ev['id']}")


def details_section(ev: dict):
    with st.form(f"edit_event_{ev['id']}"):
        c1, c2 = st.columns(2)
        name = c1.text_input("Event name", value=ev["event_name"])
        ev_type = c2.selectbox("Event type", EVENT_TYPES, index=EVENT_TYPES.index(ev["event_type"]))
        c3, c4 = st.columns(2)
        current_date = date.fromisoformat(ev["event_date"]) if ev["event_date"] else None
        ev_date = c3.date_input("Date", value=current_date, format="DD/MM/YYYY")
        venue = c4.text_input("Venue", value=ev["venue"] or "")
        c5, c6 = st.columns(2)
        prereg_open = c5.toggle("Use pre-registration", value=bool(ev["prereg_open"]),
                                help="Turn off for attendance-only events.")
        is_open = c6.toggle("Day-of check-in open", value=bool(ev["is_open"]),
                            help="Turn this off after the event so no one can check in late.")
        save = st.form_submit_button("Save changes", type="primary")
    if save:
        if not name.strip():
            st.error("The event needs a name.")
        else:
            update_event(ev["id"], name, ev_type, ev_date.isoformat() if ev_date else None,
                         venue, is_open, prereg_open)
            notify(f"Saved changes to “{name.strip()}”.")
            st.rerun()


def flyer_section(ev: dict):
    current = get_flyer(ev["id"])
    if current:
        st.image(current, caption="Current flyer, as members see it", width=280)
    else:
        st.caption("This event has no flyer yet.")
    with st.form(f"flyer_form_{ev['id']}", clear_on_submit=True):
        new_file = st.file_uploader("Upload a new flyer" if current else "Upload a flyer",
                                    type=["jpg", "jpeg", "png"])
        save = st.form_submit_button("Save flyer", type="primary")
    if save:
        if new_file is None:
            st.warning("Choose an image first.")
        else:
            try:
                set_event_flyer(ev["id"], prepare_flyer(new_file))
                notify(f"Flyer updated for “{ev['event_name']}”.")
                st.rerun()
            except ValueError as err:
                st.error(str(err))
    if current and st.button("Remove flyer", key=f"remove_flyer_{ev['id']}"):
        set_event_flyer(ev["id"], None)
        notify(f"Flyer removed from “{ev['event_name']}”.")
        st.rerun()


def delete_section(ev: dict):
    st.markdown(
        f"Deleting **{esc(ev['event_name'])}** permanently removes the event, its "
        f"**{ev['attendees']} check-in record(s)** and its **{ev['preregs']} pre-registration(s)**. "
        "Members stay in the directory because they may belong to other events. This can't be undone."
    )
    stem = safe_filename(ev["event_name"])
    c1, c2 = st.columns(2)
    if ev["attendees"]:
        c1.download_button("Export check-ins first (Excel)", to_excel_bytes(get_ledger(ev["id"]), "Check-ins"),
                           file_name=f"{stem}_checkins.xlsx", mime=XLSX_MIME,
                           key=f"pre_delete_chk_{ev['id']}", width="stretch")
    if ev["preregs"]:
        c2.download_button("Export pre-registrations first (Excel)",
                           to_excel_bytes(get_prereg_ledger(ev["id"]), "Pre-registrations"),
                           file_name=f"{stem}_preregistrations.xlsx", mime=XLSX_MIME,
                           key=f"pre_delete_reg_{ev['id']}", width="stretch")
    typed = st.text_input(f"Type the event name to confirm: {ev['event_name']}",
                          key=f"confirm_delete_{ev['id']}")
    confirmed = typed.strip().casefold() == ev["event_name"].strip().casefold()
    if st.button("Delete this event", type="primary", disabled=not confirmed,
                 key=f"delete_event_{ev['id']}"):
        checkins, preregs = delete_event(ev["id"])
        st.session_state.pop("manage_event", None)
        st.session_state.pop(f"confirm_delete_{ev['id']}", None)
        notify(f"Deleted “{ev['event_name']}”, {checkins} check-in(s) and {preregs} pre-registration(s).")
        st.rerun()


def events_tab():
    events = get_events()
    create_event_section(bool(events))
    if not events:
        return

    st.subheader("Manage an event")
    ev = event_picker("Event", "manage_event", events)
    event_card(ev)
    current = st.session_state.get("base_url_input") or detect_base_url()
    looks_ok = bool(re.fullmatch(r"https?://[^\s/]+\.[^\s/]+", current)) and "your-app" not in current
    with st.expander("App address used in links and QR codes", expanded=not looks_ok):
        base = base_url_input()

    t_reg, t_chk, t_details, t_flyer, t_delete = st.tabs(
        [":material/link: Pre-registration link", ":material/qr_code_2: Check-in QR & poster",
         ":material/edit: Edit details", ":material/image: Flyer", ":material/delete: Delete"])
    with t_reg:
        prereg_section(ev, base)
    with t_chk:
        checkin_qr_section(ev, base)
    with t_details:
        details_section(ev)
    with t_flyer:
        flyer_section(ev)
    with t_delete:
        delete_section(ev)


def prereg_tab():
    events = get_events()
    if not events:
        st.info("No events yet.")
        return
    ev = event_picker("Event", "prereg_event", events)
    df = get_prereg_ledger(ev["id"])
    if not ev["prereg_open"] and df.empty:
        st.info("Pre-registration is turned off for this event, so it only has day-of check-ins. "
                "You can turn it on in Events → Pre-registration link.", icon=":material/info:")
        return
    arrived = int((df["Arrived At"] != "").sum())
    c = st.columns(3)
    c[0].metric("Pre-registered", len(df))
    c[1].metric("Arrived on the day", arrived)
    c[2].metric("Not yet arrived", len(df) - arrived)

    c1, c2 = st.columns([2, 1])
    search = c1.text_input("Search by name or phone", key="prereg_search")
    show = c2.radio("Show", ["Everyone", "Arrived", "Not yet arrived"], horizontal=True, key="prereg_show")
    view = filter_people(df, search)
    if show == "Arrived":
        view = view[view["Arrived At"] != ""]
    elif show == "Not yet arrived":
        view = view[view["Arrived At"] == ""]
    st.dataframe(view, width="stretch", hide_index=True)
    st.caption("This list is kept separately from the day-of check-in ledger. "
               "“Arrived At” fills in when a pre-registered person checks in on the day.")
    download_pair(view, f"{safe_filename(ev['event_name'])}_preregistrations", "Pre-registrations",
                  f"prereg_{ev['id']}")


def attendance_tab(limited: bool = False):
    """limited=True is the Usher view: live list and manual check-in, but no
    contact details, parents' numbers or exports."""
    events = get_events()
    if not events:
        st.info("No events yet.")
        return
    ev = event_picker("Event", "ledger_event", events)

    with st.expander(":material/edit_note: Manual check-in (for someone without a phone)"):
        with st.form(f"manual_checkin_{ev['id']}", clear_on_submit=True):
            phone_raw = st.text_input("Member's registered phone number")
            go = st.form_submit_button("Check them in", type="primary")
        if go:
            phone = normalize_phone(phone_raw)
            member_id = find_member_id(phone) if is_valid_phone(phone) else None
            if member_id is None:
                st.error("No member has that number. New members need to register on the check-in page.")
            elif record_check_in(member_id, ev["id"]):
                who = get_member(member_id)["full_name"]
                log_action(f"Manual check-in: {who} for “{ev['event_name']}”")
                st.success(f"{who} is checked in.")
            else:
                st.info("They're already checked in for this event.")

    auto = st.toggle(f"Auto-refresh every {LEDGER_REFRESH_SECONDS} seconds", value=True)

    def render_ledger():
        df = get_ledger(ev["id"])
        c = st.columns(4)
        c[0].metric("Checked in", len(df))
        c[1].metric("Pre-registered", int((df["Pre-Registered"] == "Yes").sum()))
        c[2].metric("First-timers", int((df["First Visit"] == "Yes").sum()))
        c[3].metric("Under 18", int((df["Under 18"] == "Yes").sum()))
        view = filter_people(df, st.text_input("Search by name or phone", key="ledger_search"))
        if limited:
            view = view[["#", "Check-In Time", "Full Name", "Pre-Registered", "First Visit", "Under 18"]]
        st.dataframe(view, width="stretch", hide_index=True)
        st.caption(f"Arrival times are recorded by the server in GMT (Accra time). "
                   f"Last refreshed {datetime.now(timezone.utc):%Y-%m-%d %H:%M:%S}.")
        if not limited:
            download_pair(view, f"{safe_filename(ev['event_name'])}_checkins", "Check-ins", f"ledger_{ev['id']}")

    st.fragment(run_every=LEDGER_REFRESH_SECONDS if auto else None)(render_ledger)()


def member_detail(member_id: int):
    m = get_member(member_id)
    if not m:
        return
    history = member_history(member_id)
    st.markdown(f"#### {esc(m['full_name'])}")
    attended = int((history["Checked In"] != "").sum()) if not history.empty else 0
    st.caption(f"Registered {m['created_at'][:10]} · attended {attended} event(s)"
               + (" · under 18" if m["is_minor"] else ""))

    t_hist, t_edit, t_delete = st.tabs(["History", "Edit details", "Remove member"])
    with t_hist:
        if history.empty:
            st.caption("No registrations or check-ins yet.")
        else:
            st.dataframe(history, hide_index=True, width="stretch")

    with t_edit:
        with st.form(f"edit_member_{member_id}"):
            full_name = st.text_input("Full name", value=m["full_name"])
            c1, c2 = st.columns(2)
            phone = c1.text_input("Phone", value=m["phone_number"])
            whatsapp = c2.text_input("WhatsApp", value=m["whatsapp_number"] or "")
            email = st.text_input("Email", value=m["email"] or "")
            c3, c4 = st.columns(2)
            ec_name = c3.text_input("Emergency contact name", value=m["emergency_contact_name"] or "")
            ec_phone = c4.text_input("Emergency contact phone", value=m["emergency_contact_phone"] or "")
            is_minor = st.checkbox("Under 18", value=bool(m["is_minor"]))
            parent = st.text_input("Parent/guardian phone", value=m["parent_guardian_phone"] or "")
            save = st.form_submit_button("Save changes", type="primary")
        if save:
            data = {
                "full_name": full_name.strip(),
                "phone_number": normalize_phone(phone),
                "whatsapp_number": normalize_phone(whatsapp) or None,
                "email": email.strip() or None,
                "emergency_contact_name": ec_name.strip() or None,
                "emergency_contact_phone": normalize_phone(ec_phone) or None,
                "parent_guardian_phone": normalize_phone(parent) or None,
                "is_minor": int(is_minor),
            }
            if len(data["full_name"]) < 2 or not is_valid_phone(data["phone_number"]):
                st.error("A name and a valid phone number are required.")
            elif data["email"] and not is_valid_email(data["email"]):
                st.error("The email address doesn't look right.")
            elif is_minor and not data["parent_guardian_phone"]:
                st.error("Members under 18 need a parent/guardian phone number.")
            else:
                try:
                    update_member(member_id, data)
                    notify("Member details saved.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Another member already uses that phone number.")

    with t_delete:
        st.markdown("Removes this person with all their check-ins and pre-registrations from every event. "
                    "Use this when someone asks for their data to be deleted.")
        typed = st.text_input(f"Type their phone number to confirm: {m['phone_number']}",
                              key=f"confirm_member_{member_id}")
        if st.button("Remove member", type="primary", key=f"delete_member_{member_id}",
                     disabled=normalize_phone(typed) != m["phone_number"]):
            removed = delete_member(member_id)
            st.session_state.pop("member_pick", None)
            notify(f"Removed {m['full_name']} and {removed} record(s).")
            st.rerun()


def members_tab():
    c1, c2 = st.columns([3, 1])
    term = c1.text_input("Search members by name or phone", key="member_search")
    everyone = all_members_export()
    c2.write("")
    c2.download_button("Export all (Excel)", to_excel_bytes(everyone, "Members"),
                       file_name="Ignite_members.xlsx", width="stretch", mime=XLSX_MIME,
                       disabled=everyone.empty)

    df = search_members(term)
    if df.empty:
        st.caption("No members found." if term else "No members have registered yet.")
        return
    st.caption(f"{len(df)} member(s)")
    st.dataframe(df.drop(columns=["id"]), hide_index=True, width="stretch", height=300)

    labels = dict(zip(df["id"], df["Full Name"] + " · " + df["Phone"]))
    ids = list(labels)
    if st.session_state.get("member_pick") not in ids:
        st.session_state.pop("member_pick", None)
    picked = st.selectbox("Open a member's record", ids, format_func=labels.get,
                          index=None, placeholder="Choose a member…", key="member_pick")
    if picked:
        member_detail(int(picked))


def backup_tab():
    st.markdown("#### Download a backup")
    st.markdown("Saves everything: events, flyers, members, pre-registrations and every check-in. "
                "Keep it on your computer or Google Drive.")
    if st.button("Prepare backup file", key="prep_backup"):
        st.session_state.backup_bytes = make_backup_bytes()
        log_action("Downloaded a full backup")
        st.session_state.backup_name = f"ignite_backup_{datetime.now():%Y-%m-%d_%H%M}.db"
    if st.session_state.get("backup_bytes"):
        st.download_button("Download backup", st.session_state.backup_bytes,
                           file_name=st.session_state.backup_name,
                           mime="application/octet-stream", type="primary")

    st.divider()
    st.markdown("#### Restore from a backup")
    st.markdown("Use this if the app restarted and your data is gone. "
                "**Everything currently in the app is replaced** by the backup.")
    upload = st.file_uploader("Backup file (.db)", type=["db", "sqlite", "sqlite3"], key="restore_file")
    sure = st.checkbox("I understand the current data will be replaced", key="restore_sure")
    if st.button("Restore backup", type="primary", disabled=not (upload and sure), key="restore_go"):
        try:
            counts = restore_backup(upload.getvalue())
            for key in ("manage_event", "ledger_event", "prereg_event", "member_pick", "team_pick",
                        "restore_sure", "backup_bytes"):
                st.session_state.pop(key, None)
            notify(f"Backup restored: {counts['events']} events, {counts['members']} members, "
                   f"{counts['attendance']} check-ins.")
            st.rerun()
        except ValueError as err:
            st.error(str(err))


def valid_username(u: str) -> bool:
    return bool(re.fullmatch(r"[A-Za-z0-9._-]{3,30}", u))


def account_section():
    user = current_admin()
    st.markdown("#### Your account")
    st.caption(f"Signed in as **{user['name']}** (username “{user['username']}”) · {user['role']}")
    if user["role"] == "Owner":
        st.caption("This is the main administrator account. Its password is set in Streamlit "
                   "Secrets (ADMIN_PASSWORD), so change it there.")
        return
    with st.form("change_own_password", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        current = c1.text_input("Current password", type="password")
        new = c2.text_input("New password", type="password", help="At least 8 characters")
        again = c3.text_input("Repeat new password", type="password")
        go = st.form_submit_button("Change my password")
    if go:
        me = get_admin(user["id"])
        if not me or not verify_password(current, me["password_hash"]):
            st.error("Your current password is incorrect.")
        elif len(new) < 8:
            st.error("The new password must be at least 8 characters.")
        elif new != again:
            st.error("The two new passwords don't match.")
        else:
            update_admin(user["id"], password_hash=hash_password(new))
            notify("Changed their own password.")
            st.rerun()


def team_tab():
    user = current_admin()
    st.markdown("#### Team members")
    st.caption("Give trusted people their own sign-in instead of sharing the main password. "
               "Suspend or remove them at any time, and see everything they do in the activity log.")
    for role, desc in TEAM_ROLES.items():
        st.markdown(f"- **{role}:** {desc}")

    df = list_admins()
    with st.expander(":material/person_add: Add a team member", expanded=df.empty):
        with st.form("add_admin", clear_on_submit=True):
            c1, c2 = st.columns(2)
            full_name = c1.text_input("Full name")
            username = c2.text_input("Username", help="3–30 letters, numbers, dots, dashes or underscores")
            c3, c4 = st.columns(2)
            password = c3.text_input("Temporary password", type="password",
                                     help="At least 8 characters. Tell them to change it after signing in.")
            role = c4.selectbox("Role", list(TEAM_ROLES), format_func=lambda r: f"{r}: {TEAM_ROLES[r]}")
            add = st.form_submit_button("Add team member", type="primary")
        if add:
            problems = []
            if len(full_name.strip()) < 2:
                problems.append("Enter their full name.")
            if not valid_username(username.strip()):
                problems.append("Usernames need 3–30 letters, numbers, dots, dashes or underscores (no spaces).")
            elif username.strip().casefold() == owner_username().casefold():
                problems.append("That username is reserved for the main administrator.")
            if len(password) < 8:
                problems.append("The temporary password must be at least 8 characters.")
            if problems:
                show_errors(problems)
            else:
                try:
                    create_admin(full_name, username, password, role, user["name"])
                    notify(f"Added {full_name.strip()} (“{username.strip()}”) to the team as {role}.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("That username is already taken.")

    if df.empty:
        st.caption("No team members yet.")
    else:
        st.dataframe(df.drop(columns=["id"]), hide_index=True, width="stretch")
        labels = dict(zip(df["id"], df["Name"] + " (" + df["Username"] + ") · " + df["Role"] + " · " + df["Status"]))
        ids = list(labels)
        if st.session_state.get("team_pick") not in ids:
            st.session_state.pop("team_pick", None)
        picked = st.selectbox("Manage a team member", ids, format_func=labels.get, index=None,
                              placeholder="Choose a team member…", key="team_pick")
        if picked:
            member = get_admin(int(picked))
            is_me = member["id"] == user["id"]
            if is_me:
                st.info("This is your own account. Another admin can change your role or access.")
            else:
                c1, c2, c3 = st.columns(3)
                with c1:
                    new_role = st.selectbox("Role", list(TEAM_ROLES), index=list(TEAM_ROLES).index(member["role"]),
                                            key=f"role_{member['id']}")
                    if st.button("Save role", key=f"save_role_{member['id']}", width="stretch",
                                 disabled=new_role == member["role"]):
                        update_admin(member["id"], role=new_role)
                        notify(f"Changed {member['full_name']}'s role to {new_role}.")
                        st.rerun()
                with c2:
                    st.write("")
                    st.write("")
                    label = "Suspend access" if member["is_active"] else "Restore access"
                    if st.button(label, key=f"active_{member['id']}", width="stretch"):
                        update_admin(member["id"], is_active=0 if member["is_active"] else 1)
                        notify(f"{'Suspended' if member['is_active'] else 'Restored'} access for {member['full_name']}.")
                        st.rerun()
                with c3:
                    with st.form(f"reset_pw_{member['id']}", clear_on_submit=True):
                        temp = st.text_input("New temporary password", type="password")
                        if st.form_submit_button("Reset password", width="stretch"):
                            if len(temp) < 8:
                                st.error("At least 8 characters.")
                            else:
                                update_admin(member["id"], password_hash=hash_password(temp))
                                notify(f"Reset the password for {member['full_name']}.")
                                st.rerun()
                with st.expander(f"Remove {member['full_name']} from the team"):
                    typed = st.text_input(f"Type their username to confirm: {member['username']}",
                                          key=f"confirm_admin_{member['id']}")
                    if st.button("Remove from team", type="primary", key=f"remove_admin_{member['id']}",
                                 disabled=typed.strip().casefold() != member["username"].casefold()):
                        delete_admin(member["id"])
                        st.session_state.pop("team_pick", None)
                        notify(f"Removed {member['full_name']} (“{member['username']}”) from the team.")
                        st.rerun()

    st.divider()
    account_section()

    st.divider()
    st.markdown("#### Activity log")
    st.caption("Every change made in the admin portal, newest first. Entries can't be edited or deleted here.")
    log = activity_log()
    if log.empty:
        st.caption("Nothing recorded yet.")
    else:
        who = st.selectbox("Filter by person", ["Everyone"] + sorted(log["Who"].unique()), key="log_who")
        view = log if who == "Everyone" else log[log["Who"] == who]
        st.dataframe(view, hide_index=True, width="stretch", height=320)
        download_pair(view, "Ignite_activity_log", "Activity log", "activity_log")


def admin_page():
    if not admin_is_authenticated():
        admin_login()
        return

    user = current_admin()
    c1, c2 = st.columns([5, 1])
    with c1:
        brand_header(f"Admin Portal · {user['name']} ({user['role']})", compact=True)
    with c2:
        st.write("")
        if st.button("Log out", width="stretch"):
            leave_admin()

    notice = st.session_state.pop("admin_notice", None)
    if notice:
        st.success(notice)

    if user["role"] == "Usher":
        tabs = st.tabs([":material/fact_check: Check-ins", ":material/person: My account"])
        with tabs[0]:
            attendance_tab(limited=True)
        with tabs[1]:
            account_section()
        return

    base = st.session_state.get("base_url_input") or detect_base_url()
    if base:
        st.caption(f":material/bookmark: Bookmark your private admin link: {base}/?view=admin "
                   "(it isn't shown on member pages).")

    tabs = st.tabs([":material/dashboard: Overview", ":material/event: Events", ":material/how_to_reg: Pre-registrations",
                    ":material/fact_check: Check-ins", ":material/group: Members",
                    ":material/admin_panel_settings: Team", ":material/backup: Backup"])
    with tabs[0]:
        overview_tab()
    with tabs[1]:
        events_tab()
    with tabs[2]:
        prereg_tab()
    with tabs[3]:
        attendance_tab()
    with tabs[4]:
        members_tab()
    with tabs[5]:
        team_tab()
    with tabs[6]:
        backup_tab()


# ---------------------------------------------------------------------------
# Router
# ---------------------------------------------------------------------------
def main():
    st.session_state.setdefault("form_nonce", 0)
    st.session_state.setdefault("show_admin", False)
    logged_in = bool(st.session_state.get("admin_user")) and st.session_state.get("admin_expires", 0) > time.time()
    if st.session_state.show_admin or st.query_params.get("view") == "admin" or logged_in:
        mode = "admin"
    elif st.query_params.get("mode") == "register":
        mode = "register"
    else:
        mode = "checkin"

    titles = {"admin": "Admin", "register": "Register", "checkin": "Check-In"}
    st.set_page_config(
        page_title=f"{APP_NAME} | {titles[mode]}",
        page_icon="🔥",
        layout="wide" if mode == "admin" and logged_in else "centered",
        initial_sidebar_state="collapsed",
    )
    init_db()
    inject_css(public=not (mode == "admin" and logged_in))

    if mode == "admin":
        admin_page()
    elif mode == "register":
        register_page()
        public_footer()
    else:
        checkin_page()
        public_footer()


if __name__ == "__main__":
    main()
