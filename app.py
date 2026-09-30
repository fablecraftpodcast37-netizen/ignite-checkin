"""
Ignite Prayer Network: Membership, Event Registration & Check-In
================================================================

One Streamlit app with four public pages and a private admin portal.

  * Arrival check-in (each event's QR code)   /?event=<id>
  * Pre-registration (shared before the day)  /?event=<id>&mode=register
  * Member registration (share on WhatsApp)   /?mode=join
  * Admin portal                              tap the (c) line at the foot of any
                                              member page, or open /?view=admin

How the member journey works
  Everyone starts by typing their phone number. If we already know the number,
  they're done in one tap. If we don't, a short details form appears straight
  away and they're registered and checked in (or pre-registered) in one go.
  Nobody has to decide whether they are "new" or "existing".

Records are kept apart
  * Every event has its own event_id; records are always read and written with it.
  * Pre-registrations and day-of check-ins live in separate tables.
  * Phone numbers are stored in international format (+233..., +44..., +1...),
    so members in Ghana, the UK, the USA, the UAE and elsewhere all work.

Settings (Streamlit Cloud: Settings -> Secrets)
    ADMIN_USERNAME = "admin"
    ADMIN_PASSWORD = "your-strong-password"
    APP_BASE_URL   = "https://your-app-name.streamlit.app"

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
DEFAULT_ADMIN_USERNAME = "admin"
DEFAULT_ADMIN_PASSWORD = "IgniteAdmin2026"     # change before going live (use Secrets)
TEAM_ROLES = {
    "Admin": "Full access: events, members, messages, exports, backups and the team",
    "Usher": "Event-day helper: live check-in list and manual check-in only",
}
EVENT_TYPES = ["Asteri", "Shekinah Glory", "Impromptu"]
# What members see for each programme type. Change these freely.
PUBLIC_TYPE_NAMES = {"Asteri": "Asteri", "Shekinah Glory": "Shekinah Glory", "Impromptu": "Special Meeting"}

# Countries offered in the phone-number picker (name, dialling code, flag).
# The first entry is the default. Add more lines as your network grows.
COUNTRIES = [
    ("Ghana", "+233", "🇬🇭"),
    ("United Kingdom", "+44", "🇬🇧"),
    ("United States", "+1", "🇺🇸"),
    ("Canada", "+1", "🇨🇦"),
    ("United Arab Emirates", "+971", "🇦🇪"),
    ("Nigeria", "+234", "🇳🇬"),
    ("Germany", "+49", "🇩🇪"),
    ("Netherlands", "+31", "🇳🇱"),
    ("Italy", "+39", "🇮🇹"),
    ("Ireland", "+353", "🇮🇪"),
    ("South Africa", "+27", "🇿🇦"),
    ("Qatar", "+974", "🇶🇦"),
    ("Saudi Arabia", "+966", "🇸🇦"),
    ("Other country", "", "🌍"),
]
COUNTRY_NAMES = [c[0] for c in COUNTRIES]
DIAL_CODES = {c[0]: c[1] for c in COUNTRIES}
FLAGS = {c[0]: c[2] for c in COUNTRIES}

CONFIRM_SECONDS = 7          # how long the "Submitted, thank you" screen stays up
ADMIN_SESSION_MINUTES = 30
MAX_LOGIN_ATTEMPTS = 5
LOCKOUT_SECONDS = 60
FLYER_MAX_UPLOAD_MB = 10
FLYER_MAX_WIDTH = 1200
LEDGER_REFRESH_SECONDS = 15

# Palette: midnight, royal purple and glory gold, with a small flame accent.
INK = "#1F1A33"        # deep midnight (text, dark surfaces)
ROYAL = "#4B2E83"      # royal purple (buttons, highlights)
ROYAL_DARK = "#35205F"
GOLD = "#C9A24B"       # glory gold (accents)
GOLD_TEXT = "#8C6A1E"  # gold that stays readable on light backgrounds
FLAME = "#E0662A"      # the Ignite flame
PAPER = "#FBF8F3"      # ivory page
SAND = "#F3EEE6"       # field backgrounds
LINE = "#E6DFD3"
MUTED = "#6E6780"
EMBER = ROYAL          # accent used by posters and older code
EMBER_DARK = ROYAL_DARK
TYPE_COLOURS = {"Asteri": "#2E4A8B", "Shekinah Glory": GOLD_TEXT, "Impromptu": "#2F6B5A"}
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
    prereg_open INTEGER NOT NULL DEFAULT 1,   -- 1 = event uses pre-registration
    created_at  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    flyer_bytes BLOB                          -- programme flyer, stored as JPEG
);

CREATE TABLE IF NOT EXISTS members (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    full_name               TEXT NOT NULL,
    phone_number            TEXT NOT NULL UNIQUE,   -- international format, e.g. +233241234567
    whatsapp_number         TEXT,
    email                   TEXT,
    country                 TEXT,                    -- where they live
    emergency_contact_name  TEXT,
    emergency_contact_phone TEXT,
    parent_guardian_phone   TEXT,
    is_minor                INTEGER NOT NULL DEFAULT 0,
    sms_opt_in              INTEGER NOT NULL DEFAULT 0,  -- legacy column, no longer used
    source                  TEXT,                    -- how they joined: check-in, join link, import...
    consent_at              DATETIME,
    created_at              DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
);

-- Day-of check-ins. Timestamps come from the database clock, never from the
-- member's phone. One check-in per person per event.
CREATE TABLE IF NOT EXISTS attendance (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    member_id           INTEGER NOT NULL REFERENCES members(id),
    event_id            INTEGER NOT NULL REFERENCES events(id),
    check_in_timestamp  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (member_id, event_id)
);

-- Registrations made before the day through the shared link.
CREATE TABLE IF NOT EXISTS pre_registrations (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    member_id      INTEGER NOT NULL REFERENCES members(id),
    event_id       INTEGER NOT NULL REFERENCES events(id),
    registered_at  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (member_id, event_id)
);

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

CREATE TABLE IF NOT EXISTS activity_log (
    id      INTEGER PRIMARY KEY AUTOINCREMENT,
    at      DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    actor   TEXT NOT NULL,
    action  TEXT NOT NULL
);

-- Simple key/value settings editable from the admin portal
-- (e.g. the WhatsApp community link).
CREATE TABLE IF NOT EXISTS settings (
    key    TEXT PRIMARY KEY,
    value  TEXT
);

CREATE INDEX IF NOT EXISTS idx_attendance_event  ON attendance(event_id);
CREATE INDEX IF NOT EXISTS idx_attendance_member ON attendance(member_id);
CREATE INDEX IF NOT EXISTS idx_prereg_event      ON pre_registrations(event_id);
CREATE INDEX IF NOT EXISTS idx_prereg_member     ON pre_registrations(member_id);

-- Arrival times can't be edited.
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
    ("members", "country", "TEXT"),
    ("members", "sms_opt_in", "INTEGER NOT NULL DEFAULT 0"),
    ("members", "source", "TEXT"),
]
DATA_VERSION = "intl-phones-1"   # bump when a one-off data upgrade is added


def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, timeout=15, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _legacy_to_intl(value):
    """Numbers saved by earlier versions were Ghana-local (0XXXXXXXXX).
    Convert them to international format; leave anything else untouched."""
    if not value or str(value).startswith("+"):
        return value
    digits = re.sub(r"\D", "", str(value))
    if len(digits) == 10 and digits.startswith("0"):
        return "+233" + digits[1:]
    if len(digits) == 12 and digits.startswith("233"):
        return "+" + digits
    if len(digits) >= 11:
        return "+" + digits
    return value


def _upgrade_phone_numbers(conn: sqlite3.Connection):
    rows = conn.execute(
        """SELECT id, phone_number, whatsapp_number, emergency_contact_phone, parent_guardian_phone
           FROM members
           WHERE phone_number NOT LIKE '+%' OR whatsapp_number NOT LIKE '+%'
              OR emergency_contact_phone NOT LIKE '+%' OR parent_guardian_phone NOT LIKE '+%'"""
    ).fetchall()
    for r in rows:
        for col in ("whatsapp_number", "emergency_contact_phone", "parent_guardian_phone"):
            new = _legacy_to_intl(r[col])
            if new != r[col]:
                conn.execute(f"UPDATE members SET {col} = ? WHERE id = ?", (new, r["id"]))
        new_phone = _legacy_to_intl(r["phone_number"])
        if new_phone != r["phone_number"]:
            try:
                conn.execute("UPDATE members SET phone_number = ? WHERE id = ?", (new_phone, r["id"]))
            except sqlite3.IntegrityError:
                pass   # a duplicate already exists in the new format; keep the old value
    conn.execute("UPDATE members SET country = 'Ghana' WHERE country IS NULL AND phone_number LIKE '+233%'")


def migrate(conn: sqlite3.Connection):
    """Add missing columns, create missing tables, and run one-off data
    upgrades. Existing records are never removed."""
    existing = {r["name"] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    for table, column, decl in MIGRATIONS:
        if table in existing:
            cols = {r["name"] for r in conn.execute(f"PRAGMA table_info({table})")}
            if column not in cols:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {decl}")
    conn.executescript(SCHEMA)
    done = conn.execute("SELECT value FROM settings WHERE key = 'data_version'").fetchone()
    if not done or done["value"] != DATA_VERSION:
        _upgrade_phone_numbers(conn)
        conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('data_version', ?)", (DATA_VERSION,))


# The fingerprint changes whenever the layout changes, so upgrades run after a
# code update even though Streamlit Cloud doesn't restart the app.
SCHEMA_VERSION = hashlib.sha1((SCHEMA + repr(MIGRATIONS) + DATA_VERSION).encode()).hexdigest()[:12]


@st.cache_resource
def _init_db(version: str) -> bool:
    with closing(get_conn()) as conn, conn:
        migrate(conn)
    return True


def init_db() -> bool:
    return _init_db(SCHEMA_VERSION)


def query_df(sql: str, params=()) -> pd.DataFrame:
    with closing(get_conn()) as conn:
        return pd.read_sql_query(sql, conn, params=params)


def get_app_setting(key: str, default: str = "") -> str:
    with closing(get_conn()) as conn:
        row = conn.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
        return row["value"] if row and row["value"] is not None else default


def set_app_setting(key: str, value: str):
    with closing(get_conn()) as conn, conn:
        conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, value))


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
            """INSERT INTO events (event_name, event_type, event_date, venue, flyer_bytes, is_open, prereg_open)
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
    with closing(get_conn()) as conn:
        row = conn.execute("SELECT flyer_bytes FROM events WHERE id = ?", (event_id,)).fetchone()
        return bytes(row["flyer_bytes"]) if row and row["flyer_bytes"] else None


# ----- members ---------------------------------------------------------------
MEMBER_FIELDS = ("full_name", "phone_number", "whatsapp_number", "email", "country",
                 "emergency_contact_name", "emergency_contact_phone", "parent_guardian_phone",
                 "is_minor", "source")


def find_member_id(phone: str):
    if not phone:
        return None
    with closing(get_conn()) as conn:
        row = conn.execute("SELECT id FROM members WHERE phone_number = ?", (phone,)).fetchone()
        return row["id"] if row else None


def create_member(data: dict) -> int:
    row = {f: data.get(f) for f in MEMBER_FIELDS}
    row["is_minor"] = int(row["is_minor"] or 0)
    with closing(get_conn()) as conn, conn:
        cur = conn.execute(
            f"""INSERT INTO members ({', '.join(MEMBER_FIELDS)}, consent_at)
                VALUES ({', '.join(':' + f for f in MEMBER_FIELDS)}, CURRENT_TIMESTAMP)""",
            row,
        )
        return cur.lastrowid


def get_member(member_id: int):
    with closing(get_conn()) as conn:
        row = conn.execute("SELECT * FROM members WHERE id = ?", (member_id,)).fetchone()
        return dict(row) if row else None


def update_member(member_id: int, data: dict):
    """Raises sqlite3.IntegrityError if the new phone number belongs to someone else."""
    cols = [f for f in data if f in MEMBER_FIELDS]
    with closing(get_conn()) as conn, conn:
        conn.execute(f"UPDATE members SET {', '.join(c + ' = :' + c for c in cols)} WHERE id = :id",
                     {**{c: data[c] for c in cols}, "id": member_id})


def delete_member(member_id: int) -> int:
    with closing(get_conn()) as conn, conn:
        removed = conn.execute("DELETE FROM attendance WHERE member_id = ?", (member_id,)).rowcount
        removed += conn.execute("DELETE FROM pre_registrations WHERE member_id = ?", (member_id,)).rowcount
        conn.execute("DELETE FROM members WHERE id = ?", (member_id,))
    return removed


def search_members(term: str = "", country: str = "All countries") -> pd.DataFrame:
    term = (term or "").strip()
    digits = re.sub(r"\D", "", term).lstrip("0")
    where, params = ["(? = '' OR m.full_name LIKE ? OR m.phone_number LIKE ?)"], [term, f"%{term}%", f"%{digits or term}%"]
    if country != "All countries":
        where.append("COALESCE(m.country, 'Not recorded') = ?")
        params.append(country)
    return query_df(
        f"""SELECT m.id,
                  m.full_name AS "Full Name",
                  m.phone_number AS "Phone",
                  COALESCE(m.country, '') AS "Country",
                  CASE WHEN m.is_minor THEN 'Yes' ELSE '' END AS "Under 18",
                  (SELECT COUNT(*) FROM attendance a WHERE a.member_id = m.id) AS "Events Attended",
                  (SELECT strftime('%Y-%m-%d', MAX(a.check_in_timestamp))
                     FROM attendance a WHERE a.member_id = m.id) AS "Last Seen",
                  strftime('%Y-%m-%d', m.created_at) AS "Joined",
                  COALESCE(m.source, '') AS "Joined Via"
           FROM members m
           WHERE {' AND '.join(where)}
           ORDER BY m.full_name COLLATE NOCASE""",
        params,
    )


def member_countries() -> list[str]:
    with closing(get_conn()) as conn:
        return [r[0] for r in conn.execute(
            "SELECT DISTINCT COALESCE(country, 'Not recorded') FROM members ORDER BY 1").fetchall()]


def all_members_export() -> pd.DataFrame:
    return query_df(
        """SELECT m.full_name AS "Full Name", m.phone_number AS "Phone",
                  m.whatsapp_number AS "WhatsApp", m.email AS "Email", m.country AS "Country",
                  CASE WHEN m.is_minor THEN 'Yes' ELSE '' END AS "Under 18",
                  m.parent_guardian_phone AS "Parent/Guardian Phone",
                  m.emergency_contact_name AS "Emergency Contact",
                  m.emergency_contact_phone AS "Emergency Phone",
                  (SELECT COUNT(*) FROM attendance a WHERE a.member_id = m.id) AS "Events Attended",
                  strftime('%Y-%m-%d', m.created_at) AS "Joined", m.source AS "Joined Via"
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
    try:
        with closing(get_conn()) as conn, conn:
            conn.execute("INSERT INTO attendance (member_id, event_id) VALUES (?, ?)", (member_id, event_id))
        return True
    except sqlite3.IntegrityError:
        return False


def record_pre_registration(member_id: int, event_id: int) -> bool:
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


PEOPLE_COLUMNS = """m.full_name AS "Full Name", m.phone_number AS "Phone",
                    m.whatsapp_number AS "WhatsApp", m.email AS "Email", COALESCE(m.country, '') AS "Country",
                    CASE WHEN m.is_minor THEN 'Yes' ELSE '' END AS "Under 18",
                    m.parent_guardian_phone AS "Parent/Guardian Phone",
                    m.emergency_contact_name AS "Emergency Contact",
                    m.emergency_contact_phone AS "Emergency Phone" """


def get_ledger(event_id: int) -> pd.DataFrame:
    df = query_df(
        f"""SELECT strftime('%Y-%m-%d %H:%M:%S', a.check_in_timestamp) AS "Check-In Time",
                   {PEOPLE_COLUMNS},
                   CASE WHEN p.id IS NOT NULL THEN 'Yes' ELSE '' END AS "Pre-Registered",
                   CASE WHEN a.id = (SELECT MIN(a2.id) FROM attendance a2 WHERE a2.member_id = a.member_id)
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
    df = query_df(
        f"""SELECT strftime('%Y-%m-%d %H:%M:%S', p.registered_at) AS "Registered At",
                   {PEOPLE_COLUMNS},
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
                      (SELECT COUNT(*) FROM members WHERE country IS NOT NULL AND country <> 'Ghana') AS abroad,
                      (SELECT COUNT(*) FROM pre_registrations) AS preregs,
                      (SELECT COUNT(*) FROM attendance) AS checkins,
                      (SELECT COUNT(*) FROM attendance WHERE date(check_in_timestamp) = date('now')) AS today"""
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
    with closing(get_conn()) as conn, conn:
        cur = conn.execute(
            "INSERT INTO admins (full_name, username, password_hash, role, created_by) VALUES (?, ?, ?, ?, ?)",
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
                    raise ValueError("That file isn't an Ignite backup.")
                counts = {t: src.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
                          for t in ("events", "members", "attendance")}
                with closing(get_conn()) as dst:
                    src.backup(dst)
        except sqlite3.DatabaseError:
            raise ValueError("That file isn't a valid database backup.")
    finally:
        os.remove(path)
    with closing(get_conn()) as conn, conn:
        migrate(conn)
    get_flyer.clear()
    return counts


# ---------------------------------------------------------------------------
# Phone numbers (stored in international format: +<country code><number>)
# ---------------------------------------------------------------------------
# Expected length of the national part for common countries, used to catch typos.
NATIONAL_LENGTHS = {"233": (9,), "44": (10,), "1": (10,), "971": (8, 9), "234": (10,),
                    "49": (10, 11), "31": (9,), "353": (9,), "27": (9,), "974": (8,), "966": (9,)}
COUNTRY_HINTS = {"233": "Ghana numbers have 10 digits, e.g. 024 123 4567",
                 "44": "UK mobile numbers look like 07700 900123",
                 "1": "US and Canadian numbers have 10 digits, e.g. 202 555 0147",
                 "971": "UAE mobile numbers look like 050 123 4567"}


def to_intl(raw: str, country: str = COUNTRY_NAMES[0]) -> str:
    """Turn whatever someone typed into +<code><number>. Numbers already
    starting with + or 00 are kept as typed; local numbers get the chosen
    country's code (and lose their leading 0)."""
    raw = (raw or "").strip()
    digits = re.sub(r"\D", "", raw)
    if not digits:
        return ""
    if raw.startswith("+"):
        return "+" + digits
    if raw.startswith("00"):
        return "+" + digits[2:]
    code = re.sub(r"\D", "", DIAL_CODES.get(country, "") or "")
    if not code:
        return "+" + digits
    if digits.startswith(code) and len(digits) >= len(code) + 8:
        return "+" + digits
    return f"+{code}{digits.lstrip('0')}"


def _matching_code(intl: str):
    for code in sorted(NATIONAL_LENGTHS, key=len, reverse=True):
        if intl[1:].startswith(code):
            return code
    return None


def phone_problem(intl: str):
    """None if the number looks right, otherwise a friendly explanation."""
    if not re.fullmatch(r"\+[1-9]\d{7,14}", intl or ""):
        return "That doesn't look like a complete phone number."
    code = _matching_code(intl)
    if code:
        national = len(intl) - 1 - len(code)
        if national not in NATIONAL_LENGTHS[code]:
            hint = COUNTRY_HINTS.get(code, "Please check the number and the country")
            return f"That number looks too {'short' if national < min(NATIONAL_LENGTHS[code]) else 'long'}. {hint}."
    return None


def country_from_number(intl: str) -> str:
    best = None
    for name, code, _ in COUNTRIES:
        c = code.lstrip("+")
        if c and intl[1:].startswith(c) and (best is None or len(c) > len(DIAL_CODES[best].lstrip("+"))):
            best = name
    return best or "Other country"


def fmt_phone(intl: str) -> str:
    """+233241234567 -> +233 24 123 4567 (for display only)."""
    if not intl or not intl.startswith("+"):
        return intl or ""
    code = _matching_code(intl) or intl[1:4]
    rest = intl[1 + len(code):]
    if code == "44" and len(rest) == 10:
        parts = [rest[:4], rest[4:]]
    elif len(rest) >= 9:
        parts = [rest[:-7], rest[-7:-4], rest[-4:]]
    else:
        parts = re.findall(r".{1,3}", rest)
    return "+" + code + " " + " ".join(p for p in parts if p)


def country_label(name: str) -> str:
    code = DIAL_CODES.get(name, "")
    return f"{FLAGS.get(name, '')}  {name}" + (f"  ({code})" if code else "")


def is_valid_email(email: str) -> bool:
    return bool(re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email.strip()))


# ---------------------------------------------------------------------------
# General helpers
# ---------------------------------------------------------------------------
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
    return ev["event_name"] + (f" · {fmt_date(ev['event_date'])}" if ev.get("event_date") else "")


def safe_filename(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]+", "_", text).strip("_") or "file"


def esc(text) -> str:
    return html.escape(str(text or ""))


def first_name(full_name: str) -> str:
    return (full_name or "").strip().split(" ")[0] if full_name else ""


def to_excel_bytes(df: pd.DataFrame, sheet: str) -> bytes:
    sheet = (safe_filename(sheet) or "Sheet")[:31]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet)
        ws = writer.sheets[sheet]
        for col_cells in ws.columns:
            width = max(len(str(c.value or "")) for c in col_cells) + 2
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(width, 8), 42)
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
    configured = get_setting("APP_BASE_URL").strip().rstrip("/") or get_app_setting("app_base_url").strip().rstrip("/")
    if configured:
        return configured
    try:
        parts = urlsplit(st.context.url or "")
        if parts.scheme and parts.netloc and "localhost" not in parts.netloc:
            return f"{parts.scheme}://{parts.netloc}"
    except Exception:
        pass
    return ""


def current_base_url() -> str:
    return (st.session_state.get("base_url_input") or detect_base_url() or "").strip().rstrip("/")


def base_url_ok(base: str) -> bool:
    return bool(re.fullmatch(r"https?://[^\s/]+\.[^\s/]+", base or "")) and "your-app" not in base


def checkin_link(base: str, event_id: int) -> str:
    return f"{base}/?event={event_id}"


def register_link(base: str, event_id: int) -> str:
    return f"{base}/?event={event_id}&mode=register"


def join_link(base: str) -> str:
    return f"{base}/?mode=join"


# ---------------------------------------------------------------------------
# QR codes and printable posters
# ---------------------------------------------------------------------------
def make_qr_image(data: str) -> Image.Image:
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=10, border=4)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color=INK, back_color="white")
    img = img.get_image() if hasattr(img, "get_image") else img
    return img.convert("RGB")


def to_png(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def load_font(size: int, bold: bool = False, serif: bool = False):
    if serif:
        names = ["DejaVuSerif-Bold.ttf", "LiberationSerif-Bold.ttf", "Georgia Bold.ttf", "georgiab.ttf"]
    elif bold:
        names = ["DejaVuSans-Bold.ttf", "LiberationSans-Bold.ttf", "Arial Bold.ttf", "arialbd.ttf"]
    else:
        names = ["DejaVuSans.ttf", "LiberationSans-Regular.ttf", "Arial.ttf", "arial.ttf"]
    for name in names:
        for folder in ("", "/usr/share/fonts/truetype/dejavu/", "/usr/share/fonts/truetype/liberation/"):
            try:
                return ImageFont.truetype(folder + name, size)
            except OSError:
                continue
    try:
        return ImageFont.load_default(size=size)
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


def _spaced(text: str) -> str:
    return " ".join(text.upper())


@st.cache_data(show_spinner=False)
def make_poster(link: str, name: str, event_type: str, date_text: str, venue: str,
                heading: str, call_to_action: str) -> bytes:
    """A4 (150 dpi) poster: ink header, serif title, large QR code."""
    W, H = 1240, 1754
    img = Image.new("RGB", (W, H), PAPER)
    d = ImageDraw.Draw(img)

    d.rectangle([0, 0, W, 250], fill=INK)
    d.rectangle([0, 250, W, 258], fill=GOLD)
    _centre(d, W, 78, _spaced("Ignite"), load_font(58, serif=True), PAPER)
    _centre(d, W, 158, _spaced("Prayer Network"), load_font(24), GOLD)

    _centre(d, W, 320, _spaced(heading), load_font(26, bold=True), ROYAL)
    y = 380
    title_font = load_font(76, serif=True)
    for line in _wrap(d, name, title_font, W - 180)[:3]:
        _centre(d, W, y, line, title_font, INK)
        y += 92
    meta = "  ·  ".join(p for p in (event_type, date_text, venue) if p)
    meta_font = load_font(32)
    for line in _wrap(d, meta, meta_font, W - 220)[:2]:
        _centre(d, W, y + 8, line, meta_font, MUTED)
        y += 46

    size = 700
    y += 70
    x = (W - size) // 2
    d.rounded_rectangle([x - 34, y - 34, x + size + 34, y + size + 34], radius=26, fill="white", outline=LINE, width=3)
    img.paste(make_qr_image(link).resize((size, size), Image.NEAREST), (x, y))
    y += size + 90

    _centre(d, W, y, call_to_action, load_font(40, bold=True), INK)
    _centre(d, W, y + 62, "Open your phone camera, point it at the code and tap the link.", load_font(27), MUTED)

    d.rectangle([0, H - 92, W, H], fill=INK)
    _centre(d, W, H - 60, urlsplit(link).netloc, load_font(24), "#CFC6E3")
    return to_png(img)



# ---------------------------------------------------------------------------
# Look and feel
# ---------------------------------------------------------------------------
ICON_FLAME = ('<svg class="ig-flame" viewBox="0 0 24 30" aria-hidden="true">'
              '<path d="M12 1c.6 5.2-6.5 8.4-6.5 16.2A6.5 6.5 0 0 0 12 23.7a6.5 6.5 0 0 0 6.5-6.5'
              'c0-3.6-1.9-6-3.1-7.6-.2 2.6-1.3 4-2.8 4.6C13.4 10.7 13.3 5.6 12 1z" fill="currentColor"/></svg>')
ICON_PIN = ('<svg class="ig-ico" viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M12 2a7 7 0 0 1 7 7'
            'c0 5.2-7 13-7 13S5 14.2 5 9a7 7 0 0 1 7-7zm0 4.5A2.5 2.5 0 1 0 12 11.5 2.5 2.5 0 0 0 12 6.5z"/></svg>')
ICON_PHONE = ('<svg class="ig-ico" viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M7 2h10a2 2 0 0 1 2 2'
              'v16a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2zm0 3v13h10V5H7zm5 14.2a1 1 0 1 0 0 .01z"/></svg>')
ICON_CHECK = ('<svg viewBox="0 0 52 52" aria-hidden="true"><path d="M15 27l7 7 15-16" fill="none" '
              'stroke="currentColor" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/></svg>')
ICON_INFO = ('<svg viewBox="0 0 52 52" aria-hidden="true"><path d="M26 23v13M26 16v.5" fill="none" '
             'stroke="currentColor" stroke-width="4.5" stroke-linecap="round"/></svg>')


def inject_css(public: bool):
    width = "560px" if public else "1240px"
    st.markdown(
        f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@500;600;700&family=Manrope:wght@400;500;600;700&display=swap');

:root {{ --ink:{INK}; --royal:{ROYAL}; --royal-dark:{ROYAL_DARK}; --gold:{GOLD}; --gold-text:{GOLD_TEXT};
        --flame:{FLAME}; --paper:{PAPER}; --sand:{SAND}; --line:{LINE}; --muted:{MUTED}; }}
.stApp {{ background: var(--paper); color: var(--ink); }}
.stApp, .stMarkdown, .stMarkdown p, .stMarkdown li, label, input, textarea, button p,
[data-baseweb="select"] div, .stTabs button p, [data-testid="stMetricLabel"] p, .stCaption, small {{
    font-family: 'Manrope', system-ui, -apple-system, 'Segoe UI', sans-serif; }}
.stApp p, .stApp label, .stApp input, .stApp textarea, .stApp button, .stApp li, .stApp td, .stApp th,
.stApp [data-baseweb="select"] span:not([data-testid="stIconMaterial"]) {{
    font-family: 'Manrope', system-ui, -apple-system, 'Segoe UI', sans-serif !important; }}
.stApp .stMarkdown div, .stApp .stMarkdown span, .stApp [data-testid="stCaptionContainer"] {{
    font-family: 'Manrope', system-ui, -apple-system, 'Segoe UI', sans-serif !important; }}
.stApp .stMarkdown .ig-word, .stApp .stMarkdown .ig-h1, .stApp .stMarkdown .ig-ticket-date .d,
.stApp .stMarkdown .ig-ticket-name, .stApp .stMarkdown .ig-formtitle, .stApp .stMarkdown .ig-serif,
.stApp .stMarkdown .ig-serif *, .stApp .stMarkdown h1 *, .stApp .stMarkdown h2 *, .stApp .stMarkdown h3 *,
.stApp .stMarkdown h4 * {{ font-family: 'Cormorant Garamond', Georgia, serif !important; }}
.stApp [data-testid="stIconMaterial"] {{ font-family: 'Material Symbols Rounded' !important; }}
.ig-word, .ig-h1, .ig-ticket-date .d, .ig-ticket-name, .ig-formtitle, [data-testid="stMetricValue"],
h1, h2, h3, h4, .ig-serif {{ font-variant-numeric: lining-nums; }}
h1, h2, h3, h4, .ig-serif {{ font-family: 'Cormorant Garamond', Georgia, 'Times New Roman', serif !important;
    font-weight: 600 !important; letter-spacing: 0; color: var(--ink); }}
#MainMenu, footer, [data-testid="stToolbar"], [data-testid="stDecoration"] {{ visibility: hidden; }}
header[data-testid="stHeader"] {{ background: transparent; }}
.block-container {{ max-width: {width}; padding-top: 1.4rem; padding-bottom: 3rem; }}

/* ---- brand ---- */
.ig-brand {{ display:flex; align-items:center; gap:.75rem; margin:.4rem 0 1.5rem; }}
.ig-mark {{ width:42px; height:42px; border-radius:50%; flex:none; display:flex; align-items:center; justify-content:center;
           background: radial-gradient(circle at 50% 35%, #3A2E5C 0%, var(--ink) 70%); color:var(--flame);
           box-shadow: 0 0 0 3px rgba(201,162,75,.28); }}
.ig-flame {{ width:15px; height:19px; }}
.ig-word {{ font-family:'Cormorant Garamond', Georgia, serif; font-size:1.55rem; font-weight:700; line-height:1; color:var(--ink); }}
.ig-word-sub {{ font-size:.64rem; font-weight:600; letter-spacing:.26em; text-transform:uppercase; color:var(--gold-text); margin-top:.3rem; }}
.ig-center {{ justify-content:center; }}

.ig-kicker {{ font-size:.72rem; font-weight:700; letter-spacing:.2em; text-transform:uppercase; color:var(--royal); }}
.ig-h1 {{ font-family:'Cormorant Garamond', Georgia, serif; font-size:2.35rem; line-height:1.08; font-weight:600;
         color:var(--ink); margin:.3rem 0 .5rem; }}
.ig-lead {{ color:var(--muted); font-size:.98rem; line-height:1.6; margin-bottom:1.1rem; }}

/* ---- flyer ---- */
.st-key-flyer_banner img {{ width:100%; max-height:56vh; object-fit:contain; border-radius:16px;
    box-shadow:0 20px 44px -24px rgba(31,26,51,.6); }}

/* ---- event ticket ---- */
.ig-ticket {{ display:grid; grid-template-columns:86px 1fr; background:#fff; border:1px solid var(--line);
             border-radius:16px; overflow:hidden; margin:.2rem 0 1.3rem; box-shadow:0 1px 0 rgba(31,26,51,.04); }}
.ig-ticket-date {{ background:linear-gradient(170deg, #2C2450 0%, var(--ink) 100%); color:var(--paper);
                  display:flex; flex-direction:column; align-items:center; justify-content:center; padding:.8rem .3rem; }}
.ig-ticket-date .m {{ font-size:.64rem; font-weight:600; letter-spacing:.22em; color:#CFC6E3; }}
.ig-ticket-date .d {{ font-family:'Cormorant Garamond', Georgia, serif; font-size:2.3rem; font-weight:700; line-height:1; color:#fff; }}
.ig-ticket-date .w {{ font-size:.64rem; font-weight:600; letter-spacing:.22em; color:var(--gold); }}
.ig-ticket-date .ig-flame {{ width:20px; height:26px; color:var(--flame); }}
.ig-ticket-body {{ padding:.95rem 1.1rem; border-left:2px dashed var(--line); }}
.ig-ticket-type {{ font-size:.64rem; font-weight:700; letter-spacing:.2em; text-transform:uppercase; }}
.ig-ticket-name {{ font-family:'Cormorant Garamond', Georgia, serif; font-size:1.5rem; font-weight:700; line-height:1.15;
                  margin:.2rem 0 .3rem; color:var(--ink); }}
.ig-ticket-venue {{ color:var(--muted); font-size:.88rem; }}
.ig-ico {{ width:.95rem; height:.95rem; vertical-align:-2px; margin-right:.3rem; }}

/* ---- forms ---- */
[data-testid="stForm"] {{ background:#fff; border:1px solid var(--line) !important; border-radius:18px;
                          padding:1.3rem 1.2rem 1.1rem; box-shadow:0 12px 30px -26px rgba(31,26,51,.45); }}
.ig-step {{ font-size:.7rem; font-weight:700; letter-spacing:.18em; text-transform:uppercase; color:var(--gold-text); margin-bottom:.4rem; }}
.ig-formtitle {{ font-family:'Cormorant Garamond', Georgia, serif; font-size:1.55rem; font-weight:700; color:var(--ink); margin-bottom:.15rem; }}
.ig-formnote {{ color:var(--muted); font-size:.9rem; line-height:1.55; margin-bottom:.4rem; }}
.ig-section {{ font-size:.68rem; font-weight:700; letter-spacing:.18em; text-transform:uppercase; color:var(--royal);
              margin:1rem 0 .1rem; padding-top:.85rem; border-top:1px solid var(--line); }}
.ig-chip {{ display:inline-flex; align-items:center; gap:.2rem; background:#fff; border:1px solid var(--line);
           border-radius:999px; padding:.4rem .85rem; font-size:.92rem; font-weight:600; color:var(--ink); margin:.1rem 0 .5rem; }}
.ig-chip .ig-ico {{ color:var(--royal); }}
.ig-note {{ background:#fff; border:1px solid var(--line); border-left:3px solid var(--gold); border-radius:0 12px 12px 0;
           padding:.75rem .95rem; font-size:.9rem; color:var(--ink); margin:.4rem 0 1rem; line-height:1.55; }}
[data-baseweb="input"] input, [data-baseweb="select"] > div {{ font-size:.98rem; }}

/* ---- buttons ---- */
div.stButton > button, div.stFormSubmitButton > button, div.stDownloadButton > button, div.stLinkButton > a {{
    border-radius:12px; font-weight:700; min-height:3rem; letter-spacing:.01em; }}
[data-testid="stBaseButton-primary"], [data-testid="stBaseButton-primaryFormSubmit"] {{
    background:var(--royal) !important; border-color:var(--royal) !important; color:#fff !important;
    box-shadow:0 10px 22px -14px rgba(75,46,131,.9); }}
[data-testid="stBaseButton-primary"]:hover, [data-testid="stBaseButton-primaryFormSubmit"]:hover {{
    background:var(--royal-dark) !important; border-color:var(--royal-dark) !important; }}
[data-testid="stBaseButton-secondary"], [data-testid="stBaseButton-secondaryFormSubmit"] {{
    background:#fff; border-color:var(--line); color:var(--ink); }}
[data-testid="stBaseButton-tertiary"] p {{ color:var(--royal); font-weight:600; }}

/* ---- confirmation ---- */
.ig-done {{ background: radial-gradient(120% 90% at 50% 0%, #3A2E66 0%, var(--ink) 62%); color:var(--paper);
           border-radius:24px; padding:2.4rem 1.6rem 1.9rem; text-align:center; margin:.4rem 0 1rem;
           box-shadow:0 26px 54px -30px rgba(31,26,51,.9); }}
.ig-done-seal {{ width:76px; height:76px; margin:0 auto 1.1rem; border-radius:50%; background:var(--gold);
                color:var(--ink); display:flex; align-items:center; justify-content:center;
                box-shadow:0 0 0 8px rgba(201,162,75,.18); }}
.ig-done-info .ig-done-seal {{ background:#CFC6E3; }}
.ig-done-seal svg {{ width:40px; height:40px; }}
.ig-done-kicker {{ font-size:.7rem; letter-spacing:.24em; text-transform:uppercase; color:var(--gold); font-weight:700; }}
.ig-done h2 {{ color:#fff !important; font-size:2.1rem; margin:.4rem 0 .6rem; padding:0; }}
.ig-done p {{ color:#D9D3E6; font-size:.98rem; line-height:1.65; max-width:26rem; margin:0 auto; }}
.ig-done-meta {{ margin-top:1.4rem; padding-top:1rem; border-top:1px solid rgba(255,255,255,.12);
                font-size:.78rem; color:#A79FBD; }}
.ig-muted {{ text-align:center; color:var(--muted); font-size:.82rem; }}

/* ---- footer (doubles as the quiet admin entrance) ---- */
.st-key-ig_footer_btn, .st-key-ig_footer_btn [data-testid="stButton"], .st-key-ig_footer_btn .stButton {{
    display:flex; justify-content:center; width:100%; }}
.st-key-ig_footer_btn button {{ min-height:auto !important; padding:.2rem .5rem; background:transparent !important;
                               border:none !important; cursor:default; box-shadow:none !important; }}
.st-key-ig_footer_btn button p {{ font-size:.76rem !important; font-weight:500 !important; letter-spacing:.06em;
                                 color:var(--muted) !important; }}
.st-key-ig_scroll {{ display:none; }}

/* ---- admin ---- */
.ig-pill {{ font-family:'Manrope', sans-serif; font-size:.6rem; font-weight:700; letter-spacing:.18em; text-transform:uppercase;
           border:1px solid var(--gold); color:var(--gold-text); border-radius:999px; padding:.2rem .55rem;
           margin-left:.5rem; vertical-align:middle; }}
[data-testid="stMetric"] {{ background:#fff; border:1px solid var(--line); border-radius:14px; padding:.85rem 1rem; }}
[data-testid="stMetricValue"] {{ font-family:'Cormorant Garamond', Georgia, serif; font-weight:700; color:var(--ink); }}
[data-baseweb="tab-highlight"] {{ background-color:var(--royal); }}
.stTabs [aria-selected="true"] p {{ color:var(--royal); font-weight:700; }}
[data-testid="stExpander"] details {{ background:#fff; border-color:var(--line); border-radius:14px; }}
.ig-badge {{ display:inline-block; color:#fff; font-size:.64rem; font-weight:700; letter-spacing:.14em;
            text-transform:uppercase; padding:.2rem .55rem; border-radius:999px; }}
</style>
        """,
        unsafe_allow_html=True,
    )


def brand_row(center: bool = False):
    st.markdown(
        f"""<div class="ig-brand{' ig-center' if center else ''}">
              <div class="ig-mark">{ICON_FLAME}</div>
              <div><div class="ig-word">Ignite</div><div class="ig-word-sub">Prayer Network</div></div>
            </div>""",
        unsafe_allow_html=True,
    )


def page_heading(kicker: str, title: str, lead: str = ""):
    st.markdown(
        f"""<div class="ig-kicker">{esc(kicker)}</div>
            <div class="ig-h1">{esc(title)}</div>
            {f'<div class="ig-lead">{esc(lead)}</div>' if lead else ''}""",
        unsafe_allow_html=True,
    )


def event_ticket(ev: dict, public: bool = True):
    colour = TYPE_COLOURS.get(ev["event_type"], EMBER)
    type_name = PUBLIC_TYPE_NAMES.get(ev["event_type"], ev["event_type"]) if public else ev["event_type"]
    if ev.get("event_date"):
        try:
            d = date.fromisoformat(ev["event_date"])
            month, weekday = d.strftime("%b").upper(), d.strftime("%a").upper()
            date_block = f'<span class="m">{month}</span><span class="d">{d.day}</span><span class="w">{weekday}</span>'
        except ValueError:
            date_block = ICON_FLAME
    else:
        date_block = ICON_FLAME
    venue = f'<div class="ig-ticket-venue">{ICON_PIN}{esc(ev["venue"])}</div>' if ev.get("venue") else ""
    st.markdown(
        f"""<div class="ig-ticket">
              <div class="ig-ticket-date">{date_block}</div>
              <div class="ig-ticket-body">
                <div class="ig-ticket-type" style="color:{colour}">{esc(type_name)}</div>
                <div class="ig-ticket-name">{esc(ev['event_name'])}</div>{venue}
              </div>
            </div>""",
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Public pages: shared pieces
# ---------------------------------------------------------------------------
FLOW_PREFIXES = ("chk", "reg", "join")


def reset_flow_state():
    for p in FLOW_PREFIXES:
        st.session_state.pop(f"{p}_pending", None)
        st.session_state.pop(f"{p}_known", None)


def confirm_and_reset(kind: str, title: str, message: str, context: str,
                      auto_reset: bool = True, link: tuple | None = None):
    """Show the confirmation screen, throw away the form and anything the
    person typed (a new form counter means brand-new, empty widgets)."""
    st.session_state.confirmation = {
        "kind": kind, "title": title, "message": message, "context": context,
        "time": datetime.now(timezone.utc).strftime("%H:%M"), "shown_at": None,
        "auto_reset": auto_reset, "link": link,
    }
    reset_flow_state()
    st.session_state.form_nonce += 1
    st.rerun()


def scroll_to_top():
    script = """<script>
        const d = window.parent.document;
        for (const sel of ['[data-testid="stMain"]', '[data-testid="stAppViewContainer"]', 'section.main']) {
            const el = d.querySelector(sel); if (el) el.scrollTo({top: 0, behavior: 'instant'});
        }
        window.parent.scrollTo(0, 0);
        </script>"""
    with st.container(key="ig_scroll"):
        if hasattr(st, "iframe"):
            st.iframe(script, height=1)
        else:
            st.components.v1.html(script, height=0)


def show_confirmation() -> bool:
    """Privacy screen: the page is replaced by this card. It shows no personal
    details, and on the shared check-in page it resets itself for the next person."""
    c = st.session_state.get("confirmation")
    if not c:
        return False
    if c["shown_at"] is None:
        c["shown_at"] = time.time()
    brand_row(center=True)
    seal, css = (ICON_CHECK, "ig-done") if c["kind"] == "success" else (ICON_INFO, "ig-done ig-done-info")
    st.markdown(
        f"""<div class="{css}">
              <div class="ig-done-seal">{seal}</div>
              <div class="ig-done-kicker">Submitted · thank you</div>
              <h2>{esc(c['title'])}</h2>
              <p>{esc(c['message'])}</p>
              <div class="ig-done-meta">{esc(c['context'])} &nbsp;·&nbsp; recorded {c['time']} GMT</div>
            </div>""",
        unsafe_allow_html=True,
    )
    scroll_to_top()
    if c.get("link"):
        st.link_button(c["link"][0], c["link"][1], width="stretch", type="primary")
    if st.button("Done", width="stretch", key="confirm_done",
                 type="secondary" if c.get("link") else "primary"):
        st.session_state.pop("confirmation", None)
        st.rerun()

    if c.get("auto_reset"):
        @st.fragment(run_every=1)
        def countdown():
            left = CONFIRM_SECONDS - int(time.time() - c["shown_at"])
            if left <= 0:
                st.session_state.pop("confirmation", None)
                st.rerun(scope="app")
            st.markdown(f'<div class="ig-muted">Ready for the next person in {left} second{"s" if left != 1 else ""}.</div>',
                        unsafe_allow_html=True)
        countdown()
    return True


def show_errors(errors: list[str]):
    st.error("**Please check the following:**\n\n" + "\n".join(f"- {e}" for e in errors))


def phone_step(prefix: str, title: str, note: str, button: str):
    """Step 1: the phone number. Returns {"phone", "country"} or None."""
    nonce = st.session_state.form_nonce
    with st.form(key=f"{prefix}_phone_form_{nonce}"):
        st.markdown(f'<div class="ig-formtitle">{esc(title)}</div><div class="ig-formnote">{esc(note)}</div>',
                    unsafe_allow_html=True)
        c1, c2 = st.columns([1, 1.5])
        country = c1.selectbox("Country", COUNTRY_NAMES, format_func=country_label, key=f"{prefix}_cc_{nonce}")
        raw = c2.text_input("Phone number", placeholder="e.g. 024 123 4567", autocomplete="tel",
                            key=f"{prefix}_ph_{nonce}")
        go = st.form_submit_button(button, type="primary", width="stretch")
    if not go:
        return None
    phone = to_intl(raw, country)
    problem = phone_problem(phone) if raw.strip() else "Please enter your phone number."
    if country == "Other country" and raw.strip() and not raw.strip().startswith(("+", "00")):
        problem = "For other countries, please start with + and the country code, e.g. +49 151 2345 6789."
    if problem:
        show_errors([problem])
        return None
    return {"phone": phone, "country": country if country != "Other country" else country_from_number(phone)}


def details_step(prefix: str, pending: dict, title: str, note: str, button: str, source: str):
    """Step 2, only for numbers we don't know yet. Returns the new member id or None."""
    nonce = st.session_state.form_nonce
    st.markdown(f'<div class="ig-step">Step 2 of 2</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([3, 2])
    c1.markdown(f'<span class="ig-chip">{ICON_PHONE}{esc(fmt_phone(pending["phone"]))}</span>', unsafe_allow_html=True)
    if c2.button("Use a different number", key=f"{prefix}_change_{nonce}", type="tertiary"):
        st.session_state.pop(f"{prefix}_pending", None)
        st.rerun()

    k = lambda name: f"{prefix}_{name}_{nonce}"
    home = pending["country"] if pending["country"] in COUNTRY_NAMES else "Other country"
    with st.form(key=k("details")):
        st.markdown(f'<div class="ig-formtitle">{esc(title)}</div><div class="ig-formnote">{esc(note)}</div>',
                    unsafe_allow_html=True)
        st.markdown('<div class="ig-section">About you</div>', unsafe_allow_html=True)
        full_name = st.text_input("Full name *", placeholder="e.g. Ama Serwaa Mensah", key=k("name"))
        c1, c2 = st.columns(2)
        email = c1.text_input("Email", placeholder="Optional", key=k("email"))
        lives_in = c2.selectbox("Where do you live?", COUNTRY_NAMES, index=COUNTRY_NAMES.index(home),
                                format_func=lambda n: f"{FLAGS.get(n, '')}  {n}", key=k("home"))
        whatsapp_raw = st.text_input("WhatsApp number", placeholder="Leave blank if it's the number above",
                                     key=k("wa"))

        st.markdown('<div class="ig-section">Emergency contact</div>', unsafe_allow_html=True)
        c3, c4 = st.columns(2)
        ec_name = c3.text_input("Name *", placeholder="Who should we call?", key=k("ecn"))
        ec_raw = c4.text_input("Phone *", placeholder="Their phone number", key=k("ecp"))

        st.markdown('<div class="ig-section">Under 18?</div>', unsafe_allow_html=True)
        is_minor = st.checkbox("I am under 18 years old", key=k("minor"))
        parent_raw = st.text_input("Parent or guardian's phone", placeholder="Required if you are under 18",
                                   key=k("par"))

        st.markdown('<div class="ig-section">Consent</div>', unsafe_allow_html=True)
        consent = st.checkbox(
            f"I agree that {APP_NAME} may keep these details for attendance records and may contact my "
            "emergency contact or parent/guardian if needed. *", key=k("consent"))
        st.caption("For numbers in another country, start with + and the country code.")
        go = st.form_submit_button(button, type="primary", width="stretch")
    if not go:
        return None

    num_country = pending["country"]
    whatsapp = to_intl(whatsapp_raw, num_country) if whatsapp_raw.strip() else pending["phone"]
    ec_phone = to_intl(ec_raw, num_country)
    parent = to_intl(parent_raw, num_country) if parent_raw.strip() else ""

    errors = []
    if len(full_name.strip()) < 2:
        errors.append("Enter your full name.")
    if email.strip() and not is_valid_email(email):
        errors.append("The email address doesn't look right.")
    if whatsapp_raw.strip() and phone_problem(whatsapp):
        errors.append("The WhatsApp number doesn't look right.")
    if not ec_name.strip():
        errors.append("Enter the name of your emergency contact.")
    if not ec_raw.strip() or phone_problem(ec_phone):
        errors.append("Enter a valid phone number for your emergency contact.")
    if is_minor and not parent:
        errors.append("A parent or guardian's phone number is required for anyone under 18.")
    if parent and phone_problem(parent):
        errors.append("The parent or guardian's phone number doesn't look right.")
    if not consent:
        errors.append("Tick the consent box so we can keep your details.")
    if errors:
        show_errors(errors)
        return None

    existing = find_member_id(pending["phone"])   # someone may have registered meanwhile
    if existing:
        return existing
    return create_member({
        "full_name": full_name.strip(), "phone_number": pending["phone"], "whatsapp_number": whatsapp,
        "email": email.strip() or None, "country": lives_in if lives_in != "Other country" else None,
        "emergency_contact_name": ec_name.strip(), "emergency_contact_phone": ec_phone,
        "parent_guardian_phone": parent or None, "is_minor": int(is_minor),
        "source": source,
    })


def page_top():
    """Flyer slot at the very top, then the brand. Returns the flyer slot."""
    banner = st.container(key="flyer_banner")
    brand_row()
    return banner


def show_flyer(banner, ev: dict):
    flyer = get_flyer(ev["id"])
    if flyer:
        banner.image(flyer, width="stretch")


def public_footer():
    """The (c) line looks like plain text; tapping it opens the admin sign-in."""
    st.write("")
    with st.container(key="ig_footer_btn"):
        if st.button(f"© {datetime.now().year} {APP_NAME}", key="footer_admin", type="tertiary"):
            st.session_state.show_admin = True
            st.rerun()


def whatsapp_community_link():
    url = get_app_setting("whatsapp_group_url").strip()
    return ("Join our WhatsApp community", url) if url.startswith("http") else None


# ---------------------------------------------------------------------------
# Public page: arrival check-in
# ---------------------------------------------------------------------------
def resolve_checkin_event():
    raw = st.query_params.get("event")
    if raw and str(raw).isdigit():
        ev = get_event(int(raw))
        if ev:
            return ev
        st.warning("That check-in link is no longer valid. Please choose your programme below.")
    events = get_events(open_only=True)
    if not events:
        page_heading("Check-in", "No programme is open right now",
                     "Check-in opens shortly before each programme. Please check back soon.")
        return None
    if len(events) == 1:
        return events[0]
    ids = [e["id"] for e in events]
    by_id = {e["id"]: e for e in events}
    return by_id[st.selectbox("Which programme are you attending?", ids, format_func=lambda i: public_label(by_id[i]))]


def complete_checkin(member_id: int, ev: dict, is_new: bool):
    context = ev["event_name"]
    if record_check_in(member_id, ev["id"]):
        if is_new:
            msg = "Your details are saved and your arrival is recorded. Next time, your phone number is all you need."
        elif is_pre_registered(member_id, ev["id"]):
            msg = "You registered ahead of time, and your arrival is now confirmed. Enjoy the programme."
        else:
            msg = "Your arrival has been recorded. Enjoy the programme."
        confirm_and_reset("success", "Welcome. You're checked in.", msg, context)
    else:
        confirm_and_reset("info", "You're already checked in",
                          "Your arrival was recorded earlier, so there's nothing more to do. Enjoy the programme.", context)


def checkin_page():
    if show_confirmation():
        return
    banner = page_top()
    ev = resolve_checkin_event()
    if ev is None:
        return
    show_flyer(banner, ev)
    page_heading("Arrival check-in", "Welcome. Let's get you checked in.")
    event_ticket(ev)
    if not ev["is_open"]:
        st.markdown('<div class="ig-note">Check-in for this programme is closed. '
                    'Please speak to an usher if you need help.</div>', unsafe_allow_html=True)
        return

    pending = st.session_state.get("chk_pending")
    if pending is None:
        found = phone_step("chk", "Your phone number",
                           "Enter the number you use with Ignite. If it's your first time, we'll ask for a few details next.",
                           "Continue")
        if found:
            member_id = find_member_id(found["phone"])
            if member_id:
                complete_checkin(member_id, ev, is_new=False)
            else:
                st.session_state.chk_pending = found
                st.rerun()
    else:
        member_id = details_step(
            "chk", pending, "A few details, just once",
            "We don't have this number on our list yet. That's normal if you haven't used this check-in before. "
            "Fill this in once and next time your number is all you need.",
            "Save & check me in", source="Check-in")
        if member_id:
            complete_checkin(member_id, ev, is_new=True)


# ---------------------------------------------------------------------------
# Public page: pre-registration
# ---------------------------------------------------------------------------
def complete_prereg(member_id: int, ev: dict):
    if record_pre_registration(member_id, ev["id"]):
        confirm_and_reset("success", "Your place is reserved",
                          "On the day, scan the QR code at the entrance and enter your phone number to check in.",
                          ev["event_name"], auto_reset=False, link=whatsapp_community_link())
    else:
        confirm_and_reset("info", "You're already registered",
                          "We already have your registration for this programme. We look forward to seeing you.",
                          ev["event_name"], auto_reset=False)


def register_page():
    if show_confirmation():
        return
    banner = page_top()
    raw = st.query_params.get("event")
    ev = get_event(int(raw)) if raw and str(raw).isdigit() else None
    if ev is None:
        page_heading("Registration", "This link isn't valid",
                     "Please ask the organisers for the correct registration link.")
        return
    show_flyer(banner, ev)
    page_heading("Pre-registration", "Reserve your place", "Let us know you're coming so we can prepare for you.")
    event_ticket(ev)
    if not ev["prereg_open"]:
        st.markdown('<div class="ig-note">This programme doesn\'t need pre-registration. Just come along on the day '
                    'and scan the QR code at the entrance to check in.</div>', unsafe_allow_html=True)
        return

    pending = st.session_state.get("reg_pending")
    if pending is None:
        found = phone_step("reg", "Your phone number",
                           "Enter the number you use with Ignite. If we don't have it yet, we'll ask for a few details next.",
                           "Continue")
        if found:
            member_id = find_member_id(found["phone"])
            if member_id:
                complete_prereg(member_id, ev)
            else:
                st.session_state.reg_pending = found
                st.rerun()
    else:
        member_id = details_step("reg", pending, "A few details, just once",
                                 "We don't have this number yet. Fill this in once and it's saved for every programme.",
                                 "Reserve my place", source="Pre-registration")
        if member_id:
            complete_prereg(member_id, ev)
    st.markdown('<div class="ig-note">This reserves your place. It isn\'t your check-in: on the day, '
                'scan the QR code at the entrance to confirm you\'ve arrived.</div>', unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Public page: member registration (shared in the WhatsApp group)
# ---------------------------------------------------------------------------
def join_page():
    if show_confirmation():
        return
    brand_row()
    page_heading("Membership", "Join the Ignite family",
                 "Register once so we know you're part of the family, and check in faster whenever you join us.")
    link = whatsapp_community_link()
    known = st.session_state.get("join_known")
    pending = st.session_state.get("join_pending")

    if known:
        # Number already registered. Don't reveal whose it is.
        st.session_state.pop("join_known", None)
        confirm_and_reset("info", "You're already registered",
                          "This number is already on our member list, so there's nothing more to do. Thank you.",
                          "Membership", auto_reset=False, link=link)

    if pending is None:
        found = phone_step("join", "Start with your phone number",
                           "Members abroad: choose your country first. We'll ask for a few details next.", "Continue")
        if found:
            member_id = find_member_id(found["phone"])
            if member_id:
                st.session_state.join_known = member_id
            else:
                st.session_state.join_pending = found
            st.rerun()
    else:
        member_id = details_step("join", pending, "Your details",
                                 "This takes about a minute. Your details are kept private and only used by the Ignite team.",
                                 "Join Ignite", source="Join link")
        if member_id:
            confirm_and_reset("success", "Welcome to the family",
                              "You're now on the Ignite member list. We'll keep you posted on upcoming programmes.",
                              "Membership", auto_reset=False, link=link)


# ---------------------------------------------------------------------------
# Admin: access
# ---------------------------------------------------------------------------
def current_admin() -> dict:
    return st.session_state.get("admin_user") or {}


def end_admin_session():
    for key in ("admin_expires", "admin_user"):
        st.session_state.pop(key, None)


def admin_is_authenticated() -> bool:
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
    brand_row()
    page_heading("Admin", "Sign in", "For the Ignite team only.")
    locked_until = st.session_state.get("locked_until", 0)
    if locked_until > time.time():
        st.error(f"Too many attempts. Try again in {int(locked_until - time.time())} seconds.")
        return
    with st.form("admin_login", clear_on_submit=True):
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
    if st.button("← Back to member page", key="back_to_public", type="tertiary"):
        leave_admin()


def notify(message: str):
    st.session_state.admin_notice = message
    log_action(message)


def event_picker(label: str, key: str, events: list[dict]) -> dict:
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
    digits = re.sub(r"\D", "", search).lstrip("0")
    mask = df["Full Name"].str.contains(search, case=False, na=False, regex=False)
    if digits:
        mask |= df["Phone"].str.contains(digits, na=False, regex=False)
    return df[mask]


def base_url_input():
    if not st.session_state.get("base_url_input"):
        st.session_state.base_url_input = detect_base_url()
    base = st.text_input("Public app address", key="base_url_input",
                         help="Filled in automatically. Only change it if the app's address changes.").strip().rstrip("/")
    if not base_url_ok(base):
        st.error("Enter this app's real web address (for example https://ignite-checkin-xxxx.streamlit.app) "
                 "so links and QR codes work.")
        return None
    return base


def poster_block(link: str, name: str, type_text: str, date_text: str, venue: str,
                 heading: str, cta: str, stem: str, key: str):
    poster = make_poster(link, name, type_text, date_text, venue, heading, cta)
    st.image(poster, caption="Printable poster (A4)", width="stretch")
    st.download_button("Download poster (A4 PNG)", poster, file_name=f"{stem}.png",
                       mime="image/png", width="stretch", key=f"{key}_poster")
    st.download_button("Download QR code only", to_png(make_qr_image(link)),
                       file_name=f"{stem}_QR.png", mime="image/png", width="stretch", key=f"{key}_qr")


def admin_phone_input(key: str, label: str = "Phone number", default_country: str = COUNTRY_NAMES[0]):
    c1, c2 = st.columns([1, 1.6])
    country = c1.selectbox("Country", COUNTRY_NAMES, index=COUNTRY_NAMES.index(default_country),
                           format_func=country_label, key=f"{key}_cc")
    raw = c2.text_input(label, key=f"{key}_ph")
    return raw, country


# ---------------------------------------------------------------------------
# Admin: overview
# ---------------------------------------------------------------------------
def overview_tab():
    s = overview_stats()
    c = st.columns(6)
    c[0].metric("Members", s["members"])
    c[1].metric("Living abroad", s["abroad"])
    c[2].metric("Events", s["events"], help=f"{s['open_events']} open for check-in")
    c[3].metric("Pre-registrations", s["preregs"])
    c[4].metric("Check-ins", s["checkins"])
    c[5].metric("Checked in today", s["today"])
    st.info("Free hosting can clear the database when the app restarts. Download a backup from the "
            "**Backup** tab after every event.", icon=":material/backup:")
    summary = events_summary()
    if summary.empty:
        st.caption("No events yet. Create your first one in the **Events** tab.")
        return
    st.markdown("#### Registrations and attendance")
    chart = summary.head(10).copy()
    chart["Label"] = chart["Event"] + " (#" + chart["id"].astype(str) + ")"
    st.bar_chart(chart.set_index("Label")[["Pre-Registered", "Checked In"]], horizontal=True,
                 stack=False, color=[GOLD, ROYAL], x_label="People", y_label="")
    summary["Date"] = summary["Date"].map(fmt_date)
    st.dataframe(summary.drop(columns=["id"]), hide_index=True, width="stretch")


# ---------------------------------------------------------------------------
# Admin: events
# ---------------------------------------------------------------------------
def create_event_section(has_events: bool):
    with st.expander(":material/add_circle: Create a new event", expanded=not has_events):
        with st.form("new_event", clear_on_submit=True):
            c1, c2 = st.columns(2)
            ev_type = c1.selectbox("Event type", EVENT_TYPES,
                                   format_func=lambda t: t if t != "Impromptu" else "Impromptu (members see “Special Meeting”)")
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


def prereg_section(ev: dict, base):
    c1, c2 = st.columns([3, 1])
    if ev["prereg_open"]:
        c1.markdown(f":green[●] **Pre-registration is on** · {ev['preregs']} registered so far")
    else:
        c1.markdown(":gray[●] **Pre-registration is off.** Attendance-only: members simply check in on the day.")
    if c2.button("Turn off pre-registration" if ev["prereg_open"] else "Turn on pre-registration",
                 key=f"toggle_prereg_{ev['id']}", width="stretch"):
        set_event_flag(ev["id"], "prereg_open", not ev["prereg_open"])
        notify(f"Pre-registration turned {'off' if ev['prereg_open'] else 'on'} for “{ev['event_name']}”.")
        st.rerun()
    if not ev["prereg_open"]:
        return
    if not base:
        st.warning("Set the app address above to get the registration link.")
        return
    link = register_link(base, ev["id"])
    when = f" on {fmt_date(ev['event_date'])}" if ev["event_date"] else ""
    where = f" at {ev['venue']}" if ev["venue"] else ""
    message = (f"{ev['event_name']}{when}{where}\n\nReserve your place here: {link}\n\n"
               f"On the day, simply scan the QR code at the entrance and enter your phone number. "
               f"See you there.\n{APP_NAME}")
    left, right = st.columns([3, 2])
    with left:
        st.markdown("**Registration link**")
        st.code(link, language=None)
        st.markdown("**Ready-made WhatsApp message**")
        st.code(message, language=None, wrap_lines=True)
        st.link_button("Share on WhatsApp", f"https://wa.me/?text={quote(message)}", type="primary", width="stretch")
    with right:
        poster_block(link, ev["event_name"], PUBLIC_TYPE_NAMES.get(ev["event_type"], ""), fmt_date(ev["event_date"]),
                     ev["venue"] or "", "Pre-registration", "Scan to reserve your place",
                     f"Register_{safe_filename(ev['event_name'])}", f"reg_{ev['id']}")


def checkin_qr_section(ev: dict, base):
    status = ":green[●] **Check-in is open**" if ev["is_open"] else ":red[●] **Check-in is closed**"
    c1, c2 = st.columns([3, 1])
    c1.markdown(f"{status} · {ev['attendees']} checked in")
    if c2.button("Close check-in" if ev["is_open"] else "Reopen check-in", key=f"toggle_open_{ev['id']}", width="stretch"):
        set_event_flag(ev["id"], "is_open", not ev["is_open"])
        notify(f"Check-in {'closed' if ev['is_open'] else 'reopened'} for “{ev['event_name']}”.")
        st.rerun()
    if not base:
        st.warning("Set the app address above to get the check-in QR code.")
        return
    link = checkin_link(base, ev["id"])
    left, right = st.columns([3, 2])
    with left:
        st.markdown("**Check-in link**")
        st.code(link, language=None)
        st.caption("Print the poster for the entrance. Scan it with your own phone before printing.")
    with right:
        poster_block(link, ev["event_name"], PUBLIC_TYPE_NAMES.get(ev["event_type"], ""), fmt_date(ev["event_date"]),
                     ev["venue"] or "", "Arrival check-in", "Scan to check in",
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
        prereg_open = c5.toggle("Use pre-registration", value=bool(ev["prereg_open"]))
        is_open = c6.toggle("Day-of check-in open", value=bool(ev["is_open"]))
        save = st.form_submit_button("Save changes", type="primary")
    if save:
        if not name.strip():
            st.error("The event needs a name.")
        else:
            update_event(ev["id"], name, ev_type, ev_date.isoformat() if ev_date else None, venue, is_open, prereg_open)
            notify(f"Saved changes to “{name.strip()}”.")
            st.rerun()


def flyer_section(ev: dict):
    current = get_flyer(ev["id"])
    if current:
        st.image(current, caption="Current flyer, as members see it", width=280)
    else:
        st.caption("This event has no flyer yet.")
    with st.form(f"flyer_form_{ev['id']}", clear_on_submit=True):
        new_file = st.file_uploader("Upload a new flyer" if current else "Upload a flyer", type=["jpg", "jpeg", "png"])
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
    st.markdown(f"Deleting **{esc(ev['event_name'])}** permanently removes the event, its "
                f"**{ev['attendees']} check-in record(s)** and **{ev['preregs']} pre-registration(s)**. "
                "Members stay in the directory. This can't be undone.")
    stem = safe_filename(ev["event_name"])
    c1, c2 = st.columns(2)
    if ev["attendees"]:
        c1.download_button("Export check-ins first", to_excel_bytes(get_ledger(ev["id"]), "Check-ins"),
                           file_name=f"{stem}_checkins.xlsx", mime=XLSX_MIME, key=f"pre_del_chk_{ev['id']}", width="stretch")
    if ev["preregs"]:
        c2.download_button("Export pre-registrations first", to_excel_bytes(get_prereg_ledger(ev["id"]), "Pre-registrations"),
                           file_name=f"{stem}_preregistrations.xlsx", mime=XLSX_MIME, key=f"pre_del_reg_{ev['id']}",
                           width="stretch")
    typed = st.text_input(f"Type the event name to confirm: {ev['event_name']}", key=f"confirm_delete_{ev['id']}")
    if st.button("Delete this event", type="primary", key=f"delete_event_{ev['id']}",
                 disabled=typed.strip().casefold() != ev["event_name"].strip().casefold()):
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
    st.markdown("#### Manage an event")
    ev = event_picker("Event", "manage_event", events)
    event_ticket(ev, public=False)
    with st.expander("App address used in links and QR codes", expanded=not base_url_ok(current_base_url())):
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


# ---------------------------------------------------------------------------
# Admin: pre-registrations and check-ins
# ---------------------------------------------------------------------------
def prereg_tab():
    events = get_events()
    if not events:
        st.info("No events yet.")
        return
    ev = event_picker("Event", "prereg_event", events)
    df = get_prereg_ledger(ev["id"])
    if not ev["prereg_open"] and df.empty:
        st.info("Pre-registration is turned off for this event, so it only has day-of check-ins.", icon=":material/info:")
        return
    arrived = int((df["Arrived At"] != "").sum())
    c = st.columns(3)
    c[0].metric("Pre-registered", len(df))
    c[1].metric("Arrived", arrived)
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
    download_pair(view, f"{safe_filename(ev['event_name'])}_preregistrations", "Pre-registrations", f"prereg_{ev['id']}")


def attendance_tab(limited: bool = False):
    events = get_events()
    if not events:
        st.info("No events yet.")
        return
    ev = event_picker("Event", "ledger_event", events)
    with st.expander(":material/edit_note: Manual check-in (for someone without a phone)"):
        with st.form(f"manual_checkin_{ev['id']}", clear_on_submit=True):
            raw, country = admin_phone_input(f"manual_{ev['id']}", "Member's phone number")
            go = st.form_submit_button("Check them in", type="primary")
        if go:
            phone = to_intl(raw, country)
            member_id = find_member_id(phone) if not phone_problem(phone) else None
            if member_id is None:
                st.error("No member has that number. They can register on the check-in page in under a minute.")
            elif record_check_in(member_id, ev["id"]):
                who = get_member(member_id)["full_name"]
                log_action(f"Manual check-in: {who} for “{ev['event_name']}”")
                st.success(f"{who} is checked in.")
            else:
                st.info("They're already checked in for this event.")
    auto = st.toggle(f"Refresh every {LEDGER_REFRESH_SECONDS} seconds", value=True)

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
                   f"Updated {datetime.now(timezone.utc):%H:%M:%S}.")
        if not limited:
            download_pair(view, f"{safe_filename(ev['event_name'])}_checkins", "Check-ins", f"ledger_{ev['id']}")

    st.fragment(run_every=LEDGER_REFRESH_SECONDS if auto else None)(render_ledger)()


# ---------------------------------------------------------------------------
# Admin: members
# ---------------------------------------------------------------------------
def member_detail(member_id: int):
    m = get_member(member_id)
    if not m:
        return
    history = member_history(member_id)
    attended = int((history["Checked In"] != "").sum()) if not history.empty else 0
    st.markdown(f"#### {esc(m['full_name'])}")
    st.caption(f"{fmt_phone(m['phone_number'])} · {m.get('country') or 'Country not recorded'} · joined "
               f"{m['created_at'][:10]} · attended {attended} event(s)"
               + (" · under 18" if m["is_minor"] else ""))
    t_hist, t_edit, t_delete = st.tabs(["History", "Edit details", "Remove member"])
    with t_hist:
        if history.empty:
            st.caption("No registrations or check-ins yet.")
        else:
            st.dataframe(history, hide_index=True, width="stretch")
    with t_edit:
        home = m.get("country") if m.get("country") in COUNTRY_NAMES else "Other country"
        with st.form(f"edit_member_{member_id}"):
            full_name = st.text_input("Full name", value=m["full_name"])
            c1, c2 = st.columns(2)
            phone = c1.text_input("Phone (international format)", value=m["phone_number"])
            whatsapp = c2.text_input("WhatsApp", value=m["whatsapp_number"] or "")
            c3, c4 = st.columns(2)
            email = c3.text_input("Email", value=m["email"] or "")
            country = c4.selectbox("Lives in", COUNTRY_NAMES, index=COUNTRY_NAMES.index(home),
                                   format_func=lambda n: f"{FLAGS.get(n, '')}  {n}")
            c5, c6 = st.columns(2)
            ec_name = c5.text_input("Emergency contact name", value=m["emergency_contact_name"] or "")
            ec_phone = c6.text_input("Emergency contact phone", value=m["emergency_contact_phone"] or "")
            is_minor = st.checkbox("Under 18", value=bool(m["is_minor"]))
            parent = st.text_input("Parent/guardian phone", value=m["parent_guardian_phone"] or "")
            save = st.form_submit_button("Save changes", type="primary")
        if save:
            num_country = country_from_number(m["phone_number"])
            data = {
                "full_name": full_name.strip(), "phone_number": to_intl(phone, num_country),
                "whatsapp_number": to_intl(whatsapp, num_country) if whatsapp.strip() else None,
                "email": email.strip() or None, "country": country if country != "Other country" else None,
                "emergency_contact_name": ec_name.strip() or None,
                "emergency_contact_phone": to_intl(ec_phone, num_country) if ec_phone.strip() else None,
                "parent_guardian_phone": to_intl(parent, num_country) if parent.strip() else None,
                "is_minor": int(is_minor),
            }
            problem = phone_problem(data["phone_number"])
            if len(data["full_name"]) < 2 or problem:
                st.error(problem or "A name is required.")
            elif data["email"] and not is_valid_email(data["email"]):
                st.error("The email address doesn't look right.")
            elif is_minor and not data["parent_guardian_phone"]:
                st.error("Members under 18 need a parent/guardian phone number.")
            else:
                try:
                    update_member(member_id, data)
                    notify(f"Updated {data['full_name']}'s details.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Another member already uses that phone number.")
    with t_delete:
        st.markdown("Removes this person, their check-ins, pre-registrations and text history. "
                    "Use this when someone asks for their data to be deleted.")
        typed = st.text_input(f"Type their phone number to confirm: {m['phone_number']}", key=f"confirm_member_{member_id}")
        if st.button("Remove member", type="primary", key=f"delete_member_{member_id}",
                     disabled=to_intl(typed, country_from_number(m["phone_number"])) != m["phone_number"]):
            removed = delete_member(member_id)
            st.session_state.pop("member_pick", None)
            notify(f"Removed {m['full_name']} and {removed} record(s).")
            st.rerun()


def add_member_section():
    with st.expander(":material/person_add: Add a member yourself"):
        with st.form("admin_add_member", clear_on_submit=True):
            name = st.text_input("Full name")
            raw, country = admin_phone_input("add_member")
            email = st.text_input("Email (optional)")
            go = st.form_submit_button("Add member", type="primary")
        if go:
            phone = to_intl(raw, country)
            problem = phone_problem(phone)
            if len(name.strip()) < 2 or problem:
                st.error(problem or "Enter their full name.")
            elif find_member_id(phone):
                st.error("A member with that number already exists.")
            else:
                create_member({"full_name": name.strip(), "phone_number": phone, "whatsapp_number": phone,
                               "email": email.strip() or None, "country": country if country != "Other country" else None,
                               "source": "Added by admin"})
                notify(f"Added {name.strip()} to the member list.")
                st.rerun()


def directory_section():
    add_member_section()
    c1, c2 = st.columns([2, 1])
    term = c1.text_input("Search by name or phone", key="member_search")
    country = c2.selectbox("Country", ["All countries"] + member_countries(), key="member_country")
    df = search_members(term, country)
    if df.empty:
        st.caption("No members match." if (term or country != "All countries")
                   else "No members yet. Share the membership link or import your existing list.")
        return
    st.caption(f"{len(df)} member(s)")
    st.dataframe(df.drop(columns=["id"]), hide_index=True, width="stretch", height=320)
    everyone = all_members_export()
    st.download_button("Export the full member list (Excel)", to_excel_bytes(everyone, "Members"),
                       file_name="Ignite_members.xlsx", mime=XLSX_MIME, key="export_members")
    labels = dict(zip(df["id"], df["Full Name"] + " · " + df["Phone"]))
    ids = list(labels)
    if st.session_state.get("member_pick") not in ids:
        st.session_state.pop("member_pick", None)
    picked = st.selectbox("Open a member's record", ids, format_func=labels.get, index=None,
                          placeholder="Choose a member…", key="member_pick")
    if picked:
        member_detail(int(picked))


def membership_link_section():
    base = current_base_url()
    if not base_url_ok(base):
        st.warning("The app's web address couldn't be detected. Enter it once and it will be remembered.")
        with st.form("save_base_url"):
            entered = st.text_input("App address", placeholder="https://ignite-checkin-xxxx.streamlit.app")
            if st.form_submit_button("Save address", type="primary"):
                entered = entered.strip().rstrip("/")
                if base_url_ok(entered):
                    set_app_setting("app_base_url", entered)
                    notify(f"Saved the app address: {entered}")
                    st.rerun()
                else:
                    st.error("Enter the full address, starting with https://")
        return
    link = join_link(base)
    message = (f"Dear Ignite family,\n\nPlease take a minute to register on our member list, "
               f"wherever you are in the world:\n{link}\n\n"
               f"You'll also check in faster at our gatherings.\n{APP_NAME}")
    left, right = st.columns([3, 2])
    with left:
        st.markdown("**Membership link.** Post it in your WhatsApp group so every member can register at any time.")
        st.code(link, language=None)
        st.markdown("**Ready-made WhatsApp message**")
        st.code(message, language=None, wrap_lines=True)
        st.link_button("Share on WhatsApp", f"https://wa.me/?text={quote(message)}", type="primary", width="stretch")
        st.divider()
        st.markdown("**WhatsApp community link (optional).** When set, new members see a “Join our WhatsApp "
                    "community” button after registering.")
        with st.form("wa_link_form"):
            url = st.text_input("WhatsApp group or community invite link", value=get_app_setting("whatsapp_group_url"),
                                placeholder="https://chat.whatsapp.com/…")
            if st.form_submit_button("Save link"):
                if url.strip() and not url.strip().startswith("https://"):
                    st.error("Paste the full invite link, starting with https://")
                else:
                    set_app_setting("whatsapp_group_url", url.strip())
                    notify("Updated the WhatsApp community link." if url.strip() else "Removed the WhatsApp community link.")
                    st.rerun()
    with right:
        poster_block(link, "Join the Ignite family", "", "", "", "Member registration", "Scan to join the member list",
                     "Ignite_membership", "join")


IMPORT_NONE = "— not in my file —"


def _guess(columns, words):
    for i, c in enumerate(columns):
        if any(w in str(c).lower() for w in words):
            return i
    return 0


def import_section():
    st.markdown("Bring in the members you already have (for example from an Excel sheet or a WhatsApp export), "
                "so they're recognised the first time they check in.")
    upload = st.file_uploader("Excel or CSV file", type=["xlsx", "csv"], key="import_file")
    if not upload:
        st.caption("Your file needs at least a name column and a phone number column. Other columns are optional.")
        return
    try:
        raw = pd.read_csv(upload, dtype=str) if upload.name.lower().endswith(".csv") else pd.read_excel(upload, dtype=str)
    except Exception:
        st.error("That file couldn't be read. Save it as .xlsx or .csv and try again.")
        return
    raw = raw.dropna(how="all")
    cols = list(raw.columns)
    if not cols:
        st.error("The file looks empty.")
        return
    c1, c2, c3 = st.columns(3)
    name_col = c1.selectbox("Column with names", cols, index=_guess(cols, ["name"]))
    phone_col = c2.selectbox("Column with phone numbers", cols, index=_guess(cols, ["phone", "mobile", "number", "tel"]))
    email_col = c3.selectbox("Column with emails", [IMPORT_NONE] + cols,
                             index=(_guess(cols, ["mail"]) + 1) if any("mail" in str(c).lower() for c in cols) else 0)
    default_country = st.selectbox("Numbers without a + are from", COUNTRY_NAMES[:-1], format_func=country_label)

    rows, seen = [], set()
    for _, r in raw.iterrows():
        name = str(r.get(name_col) or "").strip()
        phone = to_intl(str(r.get(phone_col) or ""), default_country)
        email = str(r.get(email_col) or "").strip() if email_col != IMPORT_NONE else ""
        if not name or name.lower() == "nan":
            status = "Skipped: no name"
        elif phone_problem(phone):
            status = "Skipped: phone number looks wrong"
        elif phone in seen or find_member_id(phone):
            status = "Already on the list"
        else:
            status = "Will be added"
        seen.add(phone)
        rows.append({"Name": name, "Phone": phone, "Email": email if email.lower() != "nan" else "", "Result": status})
    preview = pd.DataFrame(rows)
    new = preview[preview["Result"] == "Will be added"]
    st.dataframe(preview, hide_index=True, width="stretch", height=280)
    st.caption(f"{len(new)} new · {int((preview['Result'] == 'Already on the list').sum())} already on the list · "
               f"{int(preview['Result'].str.startswith('Skipped').sum())} skipped")
    if st.button(f"Import {len(new)} member(s)", type="primary", disabled=new.empty, key="do_import"):
        added = 0
        for _, r in new.iterrows():
            try:
                create_member({"full_name": r["Name"], "phone_number": r["Phone"], "whatsapp_number": r["Phone"],
                               "email": r["Email"] or None, "country": country_from_number(r["Phone"]),
                               "source": "Imported"})
                added += 1
            except sqlite3.IntegrityError:
                pass
        st.session_state.pop("import_file", None)
        notify(f"Imported {added} member(s) from {upload.name}.")
        st.rerun()


def members_tab():
    t_dir, t_link, t_import = st.tabs([":material/groups: Directory", ":material/share: Membership link",
                                       ":material/upload_file: Import existing members"])
    with t_dir:
        directory_section()
    with t_link:
        membership_link_section()
    with t_import:
        import_section()


# ---------------------------------------------------------------------------
# Admin: team and backup
# ---------------------------------------------------------------------------
def valid_username(u: str) -> bool:
    return bool(re.fullmatch(r"[A-Za-z0-9._-]{3,30}", u))


def account_section():
    user = current_admin()
    st.markdown("#### Your account")
    st.caption(f"Signed in as **{user['name']}** (username “{user['username']}”) · {user['role']}")
    if user["role"] == "Owner":
        st.caption("This is the main administrator account. Its password is set in Streamlit Secrets (ADMIN_PASSWORD).")
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
    st.caption("Give trusted people their own sign-in instead of sharing the main password.")
    for role, desc in TEAM_ROLES.items():
        st.markdown(f"- **{role}:** {desc}")
    df = list_admins()
    with st.expander(":material/person_add: Add a team member", expanded=df.empty):
        with st.form("add_admin", clear_on_submit=True):
            c1, c2 = st.columns(2)
            full_name = c1.text_input("Full name")
            username = c2.text_input("Username", help="3–30 letters, numbers, dots, dashes or underscores")
            c3, c4 = st.columns(2)
            password = c3.text_input("Temporary password", type="password", help="At least 8 characters.")
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
            if member["id"] == user["id"]:
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
                    if st.button("Suspend access" if member["is_active"] else "Restore access",
                                 key=f"active_{member['id']}", width="stretch"):
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
    st.caption("Every change made in the admin portal, newest first.")
    log = activity_log()
    if log.empty:
        st.caption("Nothing recorded yet.")
    else:
        who = st.selectbox("Filter by person", ["Everyone"] + sorted(log["Who"].unique()), key="log_who")
        view = log if who == "Everyone" else log[log["Who"] == who]
        st.dataframe(view, hide_index=True, width="stretch", height=320)
        download_pair(view, "Ignite_activity_log", "Activity log", "activity_log")


def backup_tab():
    st.markdown("#### Download a backup")
    st.markdown("Saves everything: events, flyers, members, registrations, check-ins, and team accounts.")
    if st.button("Prepare backup file", key="prep_backup"):
        st.session_state.backup_bytes = make_backup_bytes()
        log_action("Downloaded a full backup")
        st.session_state.backup_name = f"ignite_backup_{datetime.now():%Y-%m-%d_%H%M}.db"
    if st.session_state.get("backup_bytes"):
        st.download_button("Download backup", st.session_state.backup_bytes, file_name=st.session_state.backup_name,
                           mime="application/octet-stream", type="primary")
    st.divider()
    st.markdown("#### Restore from a backup")
    st.markdown("Use this if the app restarted and your data is gone. **Everything currently in the app is replaced.**")
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


def admin_page():
    if not admin_is_authenticated():
        admin_login()
        return
    user = current_admin()
    c1, c2 = st.columns([4, 1])
    with c1:
        st.markdown(f"""<div class="ig-brand" style="margin:0">
              <div class="ig-mark">{ICON_FLAME}</div>
              <div><div class="ig-word">Ignite <span class="ig-pill">Admin</span></div>
              <div class="ig-word-sub">{esc(user['name'])} · {esc(user['role'])}</div></div></div>""",
                    unsafe_allow_html=True)
    with c2:
        if st.button("Sign out", width="stretch"):
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
    elif st.query_params.get("mode") == "join":
        mode = "join"
    else:
        mode = "checkin"
    titles = {"admin": "Admin", "register": "Reserve your place", "join": "Join", "checkin": "Check-in"}
    st.set_page_config(page_title=f"{titles[mode]} · {APP_NAME}", page_icon="🔥",
                       layout="wide" if mode == "admin" and logged_in else "centered",
                       initial_sidebar_state="collapsed")
    init_db()
    inject_css(public=not (mode == "admin" and logged_in))
    if mode == "admin":
        admin_page()
    elif mode == "register":
        register_page()
        public_footer()
    elif mode == "join":
        join_page()
        public_footer()
    else:
        checkin_page()
        public_footer()


if __name__ == "__main__":
    main()
