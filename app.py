"""
Ignite Prayer Network: Membership, Event Registration & Check-In
================================================================

One Streamlit app with seven public pages and a private admin portal.

  * Arrival check-in (each event's QR code)   /?event=<id>
  * Pre-registration (shared before the day)  /?event=<id>&mode=register
  * Member registration (share on WhatsApp)   /?mode=join
  * Prophet Masterclass enrolment             /?mode=masterclass
  * Join Live Masterclass Room (the gate)     /?mode=live
  * Daily Google Meet prayer room             /?mode=meet
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

Where the data lives
  On Render, attach a persistent disk mounted at /data. The database is then
  kept at /data/ignite_network.db and survives restarts and redeploys.
  Without a /data folder (your own computer) it sits next to this file.
  IGNITE_DB_PATH overrides both.

Settings (Render: Environment variables. Streamlit Cloud: Secrets)
    ADMIN_USERNAME = "admin"
    ADMIN_PASSWORD = "IgniteAdmin2026"      # the master password; change it here
    APP_BASE_URL   = "https://your-app.onrender.com"

"""

import base64
import calendar
import hashlib
import hmac
import html
import io
import json
import os
import re
import secrets
import sqlite3
import tempfile
import time
from contextlib import closing
from datetime import date, datetime, timedelta, timezone
from urllib.parse import quote, urlsplit

import pandas as pd
import qrcode
import streamlit as st
from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageOps, UnidentifiedImageError

# ---------------------------------------------------------------------------
# Settings (change here, or override with Streamlit secrets)
# ---------------------------------------------------------------------------
APP_NAME = "Ignite Prayer Network"
DB_FILENAME = "ignite_network.db"
PERSISTENT_DIR = "/data"            # Render persistent disk mount path
DEFAULT_ADMIN_USERNAME = "admin"
DEFAULT_ADMIN_PASSWORD = "IgniteAdmin2026"     # change before going live (use Secrets)
TEAM_ROLES = {
    "Admin": "Full access: events, members, masterclass, Google Meet, exports, backups and the team",
    "Usher": "Event-day helper: live check-in list and manual check-in only",
    "PA": "Masterclass payments: sees masterclass registrations and approves payments only",
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
CHECKIN_CONFIRM_SECONDS = 2  # the door check-in clears itself faster, for the next person in line
MC_TYPES = ["Free", "Paid"]
MC_FREE, MC_PENDING, MC_PAID = "Free Approved", "Pending Verification", "Paid Approved"
MC_APPROVED = (MC_FREE, MC_PAID)
DEFAULT_PA_WHATSAPP = "+17813309525"   # used until a different number is saved in the admin portal
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
GREEN = "#1E7B47"      # success text on light backgrounds
AMBER = GOLD_TEXT      # waiting / pending
TYPE_COLOURS = {"Asteri": "#2E4A8B", "Shekinah Glory": GOLD_TEXT, "Impromptu": "#2F6B5A"}
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def get_setting(key: str, default: str = "") -> str:
    """Read a setting from an environment variable (Render), then Streamlit
    secrets (Streamlit Cloud), falling back to the default."""
    value = os.environ.get(key)
    if value not in (None, ""):
        return value
    try:
        return str(st.secrets[key])
    except Exception:
        return default


def resolve_db_path() -> str:
    """Use Render's persistent disk when it's there, so nothing is lost on a
    restart or redeploy. Fall back to the app folder for local testing."""
    override = os.environ.get("IGNITE_DB_PATH", "").strip()
    if override:
        folder = os.path.dirname(os.path.abspath(override))
        os.makedirs(folder, exist_ok=True)
        return override
    if os.path.isdir(PERSISTENT_DIR) and os.access(PERSISTENT_DIR, os.W_OK):
        return os.path.join(PERSISTENT_DIR, DB_FILENAME)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), DB_FILENAME)


DB_PATH = resolve_db_path()
DB_IS_PERSISTENT = os.path.abspath(DB_PATH).startswith(PERSISTENT_DIR + os.sep) or bool(
    os.environ.get("IGNITE_DB_PATH", "").strip())


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
    prereg_deadline TEXT,                     -- 'YYYY-MM-DD HH:MM' Ghana time; pre-registration closes then
    uses_checkin INTEGER NOT NULL DEFAULT 1,  -- 0 = pre-registration only, no check-in at the venue
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
    birth_day               INTEGER,                 -- 1-31, no year kept
    birth_month             INTEGER,                 -- 1-12
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
    role           TEXT NOT NULL CHECK (role IN ('Admin', 'Usher', 'PA')),
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

-- Prophet Masterclass ---------------------------------------------------------
CREATE TABLE IF NOT EXISTS masterclass_sessions (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    session_name   TEXT NOT NULL,
    session_type   TEXT NOT NULL CHECK (session_type IN ('Free', 'Paid')),
    active_status  INTEGER NOT NULL DEFAULT 1,    -- 1 = open for enrolment and the live room
    streaming_url  TEXT,                          -- the secret stream link; never shown on a page
    session_date   TEXT,                          -- YYYY-MM-DD, optional
    price_note     TEXT,                          -- e.g. "GHS 150 by MoMo to 024 000 0000"
    created_at     DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS masterclass_registrations (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    reg_id          TEXT NOT NULL UNIQUE,         -- the member's token, e.g. IGNITE-8921
    session_id      INTEGER NOT NULL REFERENCES masterclass_sessions(id),
    full_name       TEXT NOT NULL,
    email           TEXT,
    phone           TEXT NOT NULL,                -- international format
    payment_status  TEXT NOT NULL
                    CHECK (payment_status IN ('Free Approved', 'Pending Verification', 'Paid Approved')),
    approved_by     TEXT,
    approved_at     DATETIME,
    created_at      DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (session_id, phone)
);

-- Every time someone is let into a live room.
CREATE TABLE IF NOT EXISTS masterclass_entries (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    reg_id      TEXT NOT NULL,
    session_id  INTEGER NOT NULL REFERENCES masterclass_sessions(id),
    entered_at  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
);

-- Daily Google Meet prayer room ----------------------------------------------------
CREATE TABLE IF NOT EXISTS google_meet_tracker (
    id                   INTEGER PRIMARY KEY AUTOINCREMENT,
    member_display_name  TEXT NOT NULL,
    phone_number         TEXT NOT NULL,
    tracking_date        DATE NOT NULL,               -- the day in Ghana time
    join_time            TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
);

-- Who stayed to the end of the daily prayer. Filled by the closing code members
-- type at the end, by Google's own attendance report, or by an admin.
CREATE TABLE IF NOT EXISTS meet_stays (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    tracking_date  DATE NOT NULL,
    phone_number   TEXT,                          -- NULL when a report name couldn't be matched
    display_name   TEXT NOT NULL,
    email          TEXT,
    method         TEXT NOT NULL CHECK (method IN ('code', 'report', 'manual')),
    minutes        REAL,                          -- time in the call (report only)
    stayed         INTEGER NOT NULL DEFAULT 1,
    detail         TEXT,
    created_at     DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE UNIQUE INDEX IF NOT EXISTS ux_meet_stay
    ON meet_stays(tracking_date, method, COALESCE(phone_number, lower(display_name)));
CREATE INDEX IF NOT EXISTS idx_mc_reg_session   ON masterclass_registrations(session_id);
CREATE INDEX IF NOT EXISTS idx_mc_reg_phone     ON masterclass_registrations(phone);
CREATE INDEX IF NOT EXISTS idx_mc_entry_session ON masterclass_entries(session_id);
CREATE INDEX IF NOT EXISTS idx_meet_date        ON google_meet_tracker(tracking_date);
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
    ("members", "birth_day", "INTEGER"),
    ("events", "prereg_deadline", "TEXT"),
    ("events", "uses_checkin", "INTEGER NOT NULL DEFAULT 1"),
    ("members", "birth_month", "INTEGER"),
    ("masterclass_sessions", "session_date", "TEXT"),
    ("masterclass_sessions", "price_note", "TEXT"),
    ("masterclass_registrations", "approved_by", "TEXT"),
    ("masterclass_registrations", "approved_at", "DATETIME"),
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


def _allow_pa_role(conn: sqlite3.Connection):
    """Older databases only allowed Admin and Usher. Rebuild the team table
    (keeping every account) so the PA role can be used."""
    row = conn.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='admins'").fetchone()
    if not row or "'PA'" in (row["sql"] or ""):
        return
    cols = [r["name"] for r in conn.execute("PRAGMA table_info(admins)")]
    conn.execute("ALTER TABLE admins RENAME TO admins_old")
    conn.execute("CREATE TABLE admins" + SCHEMA.split("CREATE TABLE IF NOT EXISTS admins", 1)[1].split(");", 1)[0] + ")")
    keep = ", ".join(c for c in cols if c in
                     {"id", "full_name", "username", "password_hash", "role", "is_active", "created_by",
                      "created_at", "last_login"})
    conn.execute(f"INSERT INTO admins ({keep}) SELECT {keep} FROM admins_old")
    conn.execute("DROP TABLE admins_old")


def migrate(conn: sqlite3.Connection):
    """Add missing columns, create missing tables, and run one-off data
    upgrades. Existing records are never removed."""
    existing = {r["name"] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    if "admins" in existing:
        _allow_pa_role(conn)
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
    folder = os.path.dirname(os.path.abspath(DB_PATH))
    os.makedirs(folder, exist_ok=True)
    with closing(get_conn()) as conn:
        conn.execute("PRAGMA journal_mode = WAL")   # safer writes on a persistent disk
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
                   e.prereg_open, e.prereg_deadline, e.uses_checkin, e.created_at, e.flyer_bytes IS NOT NULL AS has_flyer,
                   (SELECT COUNT(*) FROM attendance a WHERE a.event_id = e.id) AS attendees,
                   (SELECT COUNT(*) FROM pre_registrations p WHERE p.event_id = e.id) AS preregs"""


def get_events(open_only: bool = False) -> list[dict]:
    where = "WHERE e.is_open = 1 AND e.uses_checkin = 1" if open_only else ""
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
                 is_open=True, prereg_open=True, prereg_deadline=None, uses_checkin=True) -> int:
    with closing(get_conn()) as conn, conn:
        cur = conn.execute(
            """INSERT INTO events (event_name, event_type, event_date, venue, flyer_bytes, is_open, prereg_open,
                                   prereg_deadline, uses_checkin)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (name.strip(), event_type, event_date, (venue or "").strip() or None, flyer,
             int(is_open), int(prereg_open), prereg_deadline, int(uses_checkin)),
        )
    get_flyer.clear()
    return cur.lastrowid


def update_event(event_id, name, event_type, event_date, venue, is_open, prereg_open, prereg_deadline=None,
                 uses_checkin=None):
    """uses_checkin=None leaves that setting as it is."""
    with closing(get_conn()) as conn, conn:
        conn.execute(
            """UPDATE events SET event_name = ?, event_type = ?, event_date = ?, venue = ?,
                                 is_open = ?, prereg_open = ?, prereg_deadline = ?,
                                 uses_checkin = COALESCE(?, uses_checkin)
               WHERE id = ?""",
            (name.strip(), event_type, event_date, (venue or "").strip() or None,
             int(is_open), int(prereg_open), prereg_deadline,
             None if uses_checkin is None else int(uses_checkin), event_id),
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


# ----- Prophet Masterclass -------------------------------------------------------
MC_SESSION_COLUMNS = """s.id, s.session_name, s.session_type, s.active_status, s.streaming_url, s.session_date,
                        s.price_note, s.created_at,
                        (SELECT COUNT(*) FROM masterclass_registrations r WHERE r.session_id = s.id) AS registered,
                        (SELECT COUNT(*) FROM masterclass_registrations r WHERE r.session_id = s.id
                           AND r.payment_status = 'Pending Verification') AS pending,
                        (SELECT COUNT(*) FROM masterclass_registrations r WHERE r.session_id = s.id
                           AND r.payment_status IN ('Free Approved', 'Paid Approved')) AS approved,
                        (SELECT COUNT(DISTINCT x.reg_id) FROM masterclass_entries x WHERE x.session_id = s.id)
                           AS joined"""


def mc_sessions(active_only: bool = False) -> list[dict]:
    where = "WHERE s.active_status = 1" if active_only else ""
    with closing(get_conn()) as conn:
        rows = conn.execute(f"""SELECT {MC_SESSION_COLUMNS} FROM masterclass_sessions s {where}
                                ORDER BY s.active_status DESC, COALESCE(s.session_date, date(s.created_at)) DESC,
                                         s.id DESC""").fetchall()
        return [dict(r) for r in rows]


def mc_session(session_id: int):
    with closing(get_conn()) as conn:
        row = conn.execute(f"SELECT {MC_SESSION_COLUMNS} FROM masterclass_sessions s WHERE s.id = ?",
                           (session_id,)).fetchone()
        return dict(row) if row else None


def mc_create_session(name, session_type, streaming_url, session_date=None, price_note=None,
                      active=True) -> int:
    with closing(get_conn()) as conn, conn:
        cur = conn.execute(
            """INSERT INTO masterclass_sessions (session_name, session_type, active_status, streaming_url,
                                                 session_date, price_note)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (name.strip(), session_type, int(active), (streaming_url or "").strip() or None, session_date,
             (price_note or "").strip() or None))
        return cur.lastrowid


def mc_update_session(session_id, name, session_type, streaming_url, session_date, price_note, active):
    with closing(get_conn()) as conn, conn:
        conn.execute(
            """UPDATE masterclass_sessions SET session_name = ?, session_type = ?, streaming_url = ?,
                      session_date = ?, price_note = ?, active_status = ? WHERE id = ?""",
            (name.strip(), session_type, (streaming_url or "").strip() or None, session_date,
             (price_note or "").strip() or None, int(active), session_id))


def mc_set_active(session_id: int, active: bool):
    with closing(get_conn()) as conn, conn:
        conn.execute("UPDATE masterclass_sessions SET active_status = ? WHERE id = ?", (int(active), session_id))


def mc_delete_session(session_id: int) -> int:
    with closing(get_conn()) as conn, conn:
        n = conn.execute("DELETE FROM masterclass_registrations WHERE session_id = ?", (session_id,)).rowcount
        conn.execute("DELETE FROM masterclass_entries WHERE session_id = ?", (session_id,))
        conn.execute("DELETE FROM masterclass_sessions WHERE id = ?", (session_id,))
        return n


def new_reg_token(conn: sqlite3.Connection) -> str:
    """IGNITE-8921 style. Four digits while there's room, then six."""
    for attempt in range(60):
        digits = 4 if attempt < 40 else 6
        token = f"IGNITE-{secrets.randbelow(9 * 10 ** (digits - 1)) + 10 ** (digits - 1)}"
        if not conn.execute("SELECT 1 FROM masterclass_registrations WHERE reg_id = ?", (token,)).fetchone():
            return token
    raise RuntimeError("Couldn't create a unique registration token.")


def normalise_token(text: str) -> str:
    t = re.sub(r"\s+", "", (text or "").upper())
    if re.fullmatch(r"\d{4,6}", t):
        t = "IGNITE-" + t
    if re.fullmatch(r"IGNITE\d{4,6}", t):
        t = "IGNITE-" + t[6:]
    return t


def mc_register(session: dict, full_name: str, email: str, phone: str) -> tuple[dict, bool]:
    """Enrol someone. Returns (registration, is_new). The same phone number on the
    same session gets its existing token back instead of a second one."""
    status = MC_FREE if session["session_type"] == "Free" else MC_PENDING
    with closing(get_conn()) as conn, conn:
        row = conn.execute("SELECT * FROM masterclass_registrations WHERE session_id = ? AND phone = ?",
                           (session["id"], phone)).fetchone()
        if row:
            return dict(row), False
        token = new_reg_token(conn)
        conn.execute(
            """INSERT INTO masterclass_registrations (reg_id, session_id, full_name, email, phone, payment_status)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (token, session["id"], full_name.strip(), (email or "").strip() or None, phone, status))
        row = conn.execute("SELECT * FROM masterclass_registrations WHERE reg_id = ?", (token,)).fetchone()
        return dict(row), True


def mc_lookup(entry: str, country: str = "Ghana") -> list[dict]:
    """Find registrations on active sessions by token or phone number."""
    token = normalise_token(entry)
    sql = """SELECT r.*, s.session_name, s.session_type, s.streaming_url, s.session_date
             FROM masterclass_registrations r JOIN masterclass_sessions s ON s.id = r.session_id
             WHERE s.active_status = 1 AND {cond} ORDER BY r.created_at DESC"""
    with closing(get_conn()) as conn:
        if token.startswith("IGNITE-"):
            rows = conn.execute(sql.format(cond="r.reg_id = ?"), (token,)).fetchall()
        else:
            phone = to_intl(entry, country)
            if phone_problem(phone):
                return []
            rows = conn.execute(sql.format(cond="r.phone = ?"), (phone,)).fetchall()
        return [dict(r) for r in rows]


def mc_log_entry(reg_id: str, session_id: int):
    with closing(get_conn()) as conn, conn:
        conn.execute("INSERT INTO masterclass_entries (reg_id, session_id) VALUES (?, ?)", (reg_id, session_id))


def mc_set_status(reg_ids: list[str], status: str, actor: str) -> int:
    assert status in (MC_FREE, MC_PENDING, MC_PAID)
    with closing(get_conn()) as conn, conn:
        n = 0
        for rid in reg_ids:
            n += conn.execute(
                """UPDATE masterclass_registrations
                   SET payment_status = ?, approved_by = CASE WHEN ? = 'Paid Approved' THEN ? ELSE NULL END,
                       approved_at = CASE WHEN ? = 'Paid Approved' THEN CURRENT_TIMESTAMP ELSE NULL END
                   WHERE reg_id = ?""", (status, status, actor, status, rid)).rowcount
        return n


def mc_delete_registration(reg_id: str):
    with closing(get_conn()) as conn, conn:
        conn.execute("DELETE FROM masterclass_entries WHERE reg_id = ?", (reg_id,))
        conn.execute("DELETE FROM masterclass_registrations WHERE reg_id = ?", (reg_id,))


def mc_registrations(session_id=None, status=None) -> pd.DataFrame:
    cond, params = [], []
    if session_id:
        cond.append("r.session_id = ?")
        params.append(session_id)
    if status:
        cond.append("r.payment_status = ?")
        params.append(status)
    where = ("WHERE " + " AND ".join(cond)) if cond else ""
    return query_df(
        f"""SELECT r.reg_id AS "Token", r.full_name AS "Full Name", r.phone AS "Phone",
                   COALESCE(r.email, '') AS "Email", s.session_name AS "Session", s.session_type AS "Type",
                   r.payment_status AS "Status",
                   strftime('%Y-%m-%d %H:%M', r.created_at) AS "Registered (GMT)",
                   COALESCE(r.approved_by, '') AS "Approved By",
                   COALESCE(strftime('%Y-%m-%d %H:%M', r.approved_at), '') AS "Approved (GMT)",
                   (SELECT COUNT(*) FROM masterclass_entries x WHERE x.reg_id = r.reg_id) AS "Room Entries",
                   COALESCE((SELECT strftime('%Y-%m-%d %H:%M', MAX(x.entered_at)) FROM masterclass_entries x
                             WHERE x.reg_id = r.reg_id), '') AS "Last Entry (GMT)"
            FROM masterclass_registrations r JOIN masterclass_sessions s ON s.id = r.session_id
            {where} ORDER BY r.created_at DESC""", params)


# ----- Daily Google Meet prayer room -----------------------------------------------
def meet_log(display_name: str, phone: str):
    with closing(get_conn()) as conn, conn:
        conn.execute(
            """INSERT INTO google_meet_tracker (member_display_name, phone_number, tracking_date)
               VALUES (?, ?, date('now'))""", (display_name.strip(), phone))


def meet_attendance(day: str) -> pd.DataFrame:
    return query_df(
        """SELECT t.member_display_name AS "Meet Display Name", t.phone_number AS "Phone",
                  COALESCE(m.full_name, '') AS "Member On File",
                  strftime('%H:%M:%S', MIN(t.join_time)) AS "First Joined (GMT)",
                  COUNT(*) AS "Times Joined"
           FROM google_meet_tracker t LEFT JOIN members m ON m.phone_number = t.phone_number
           WHERE t.tracking_date = ?
           GROUP BY t.phone_number ORDER BY MIN(t.join_time)""", (day,))


MEET_DEFAULTS = {"meet_code_minutes": "10", "meet_grace": "5", "meet_min_pct": "75",
                 "meet_start": "", "meet_end": ""}


def meet_setting(key: str) -> str:
    return get_app_setting(key, MEET_DEFAULTS.get(key, ""))


def meet_code_state() -> dict:
    """The closing code members type at the end of prayer to show they stayed."""
    code = get_app_setting("meet_code")
    day = get_app_setting("meet_code_date")
    try:
        until = float(get_app_setting("meet_code_until", "0") or 0)
    except ValueError:
        until = 0.0
    active = bool(code) and day == today_local().isoformat() and until > time.time()
    return {"code": code, "date": day, "until": until, "active": active}


def meet_open_code(minutes: int) -> str:
    code = f"{secrets.randbelow(9000) + 1000}"
    set_app_setting("meet_code", code)
    set_app_setting("meet_code_date", today_local().isoformat())
    set_app_setting("meet_code_until", str(time.time() + int(minutes) * 60))
    return code


def meet_close_code():
    set_app_setting("meet_code_until", "0")


def meet_joined_name(phone: str, day: str):
    with closing(get_conn()) as conn:
        row = conn.execute("""SELECT member_display_name FROM google_meet_tracker
                              WHERE phone_number = ? AND tracking_date = ? ORDER BY join_time LIMIT 1""",
                           (phone, day)).fetchone()
        return row["member_display_name"] if row else None


def meet_record_stay(day: str, phone, name: str, method: str, stayed: bool = True, minutes=None,
                     email=None, detail=None):
    with closing(get_conn()) as conn, conn:
        conn.execute("""DELETE FROM meet_stays WHERE tracking_date = ? AND method = ?
                        AND COALESCE(phone_number, lower(display_name)) = COALESCE(?, lower(?))""",
                     (day, method, phone, name))
        conn.execute("""INSERT INTO meet_stays (tracking_date, phone_number, display_name, email, method, minutes,
                                                stayed, detail) VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                     (day, phone, name, email, method, minutes, int(stayed), detail))


def meet_remove_stay(day: str, phone: str, method: str):
    with closing(get_conn()) as conn, conn:
        conn.execute("DELETE FROM meet_stays WHERE tracking_date = ? AND method = ? AND phone_number = ?",
                     (day, method, phone))


def _norm_name(text) -> str:
    text = re.sub(r"[^a-z0-9 ]", " ", str(text or "").casefold())
    return " ".join(text.split())


def parse_duration_minutes(value):
    """Turn '1 hr 5 min', '1:05:30', '00:45', '45 min', '65' into minutes."""
    if value is None:
        return None
    if isinstance(value, (int, float)) and not pd.isna(value):
        return float(value)
    if isinstance(value, timedelta):
        return value.total_seconds() / 60
    text = str(value).strip().lower()
    if not text or text == "nan":
        return None
    m = re.fullmatch(r"(\d+):(\d{1,2}):(\d{1,2})(\.\d+)?", text)
    if m:
        return int(m.group(1)) * 60 + int(m.group(2)) + int(m.group(3)) / 60
    m = re.fullmatch(r"(\d+):(\d{1,2})", text)
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    if re.fullmatch(r"\d+(\.\d+)?", text):
        return float(text)
    total, found = 0.0, False
    for num, unit in re.findall(r"(\d+(?:\.\d+)?)\s*(h(?:ou)?rs?|h|min(?:ute)?s?|m|sec(?:ond)?s?|s)\b", text):
        found = True
        n = float(num)
        total += n * 60 if unit.startswith("h") else n / 60 if unit.startswith("s") else n
    return total if found else None


def _read_report_table(upload) -> pd.DataFrame:
    """Google's sheet can have a few title lines above the table; find the header row."""
    name = upload.name.lower()
    raw = (pd.read_csv(upload, header=None, dtype=str) if name.endswith(".csv")
           else pd.read_excel(upload, header=None, dtype=object))
    raw = raw.dropna(how="all")
    header_at = None
    for i, (_, row) in enumerate(raw.iterrows()):
        cells = [str(c).strip().lower() for c in row.tolist()]
        if any("duration" in c for c in cells) or (any("name" in c for c in cells) and any("email" in c for c in cells)):
            header_at = i
            break
    if header_at is None:
        raise ValueError("Couldn't find the table in that file. It needs a name column and a Duration column.")
    table = raw.iloc[header_at + 1:].copy()
    table.columns = [str(c).strip() for c in raw.iloc[header_at].tolist()]
    return table.dropna(how="all")


def parse_attendance_report(table: pd.DataFrame, grace: int, min_pct: int, meeting_minutes=None) -> pd.DataFrame:
    cols = {c.lower(): c for c in table.columns}
    pick = lambda *words: next((cols[c] for c in cols if any(w in c for w in words)), None)
    first, last = pick("first"), pick("surname", "last name", "last")
    full = pick("full name", "participant", "name") if not first else None
    email_c, dur_c = pick("email", "e-mail"), pick("duration")
    join_c, exit_c = pick("joined", "join time", "join"), pick("exited", "exit", "left", "leave")
    if not (first or full) or not dur_c:
        raise ValueError("The report needs a name column and a Duration column.")
    rows = []
    for _, r in table.iterrows():
        clean = lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()
        name = (f"{clean(r.get(first))} {clean(r.get(last))}" if first else clean(r.get(full))).strip()
        if not name:
            continue
        rows.append({"Name": " ".join(name.split()),
                     "Email": clean(r.get(email_c)) if email_c else "",
                     "Minutes": parse_duration_minutes(r.get(dur_c)),
                     "_join": pd.to_datetime(str(r.get(join_c)), errors="coerce") if join_c else pd.NaT,
                     "_exit": pd.to_datetime(str(r.get(exit_c)), errors="coerce") if exit_c else pd.NaT})
    df = pd.DataFrame(rows)
    if df.empty:
        raise ValueError("No participants found in that file.")
    exits = df["_exit"].dropna()
    joins = df["_join"].dropna()
    if meeting_minutes:
        span = float(meeting_minutes)
    elif len(exits) and len(joins):
        span = (exits.max() - joins.min()).total_seconds() / 60
    else:
        span = float(df["Minutes"].max() or 0)
    end = exits.max() if len(exits) else None
    need = span * min_pct / 100
    verdicts = []
    for _, r in df.iterrows():
        mins = r["Minutes"]
        long_enough = mins is not None and not pd.isna(mins) and mins >= need
        to_end = end is None or pd.isna(r["_exit"]) or (end - r["_exit"]).total_seconds() / 60 <= grace
        stayed = bool(long_enough and to_end)
        if stayed:
            why = "Stayed to the end"
        elif not long_enough:
            why = f"Left early (in the call {mins:.0f} of {span:.0f} min)" if mins is not None and not pd.isna(mins) else "No time recorded"
        else:
            why = "Left before the end"
        verdicts.append((stayed, why))
    df["Stayed"] = [v[0] for v in verdicts]
    df["Result"] = [v[1] for v in verdicts]
    df.attrs["span"] = span
    return df


def meet_save_report(day: str, df: pd.DataFrame) -> tuple[int, int]:
    """Match report names/emails to the people who joined through the app that day."""
    joined = query_df("""SELECT t.phone_number, t.member_display_name, COALESCE(m.full_name, '') AS full_name,
                                COALESCE(m.email, '') AS email
                         FROM google_meet_tracker t LEFT JOIN members m ON m.phone_number = t.phone_number
                         WHERE t.tracking_date = ?""", (day,))
    by_name, by_email = {}, {}
    for _, j in joined.iterrows():
        for n in (j["member_display_name"], j["full_name"]):
            if _norm_name(n):
                by_name.setdefault(_norm_name(n), j["phone_number"])
        if j["email"]:
            by_email.setdefault(j["email"].strip().casefold(), j["phone_number"])
    matched = 0
    with closing(get_conn()) as conn, conn:
        conn.execute("DELETE FROM meet_stays WHERE tracking_date = ? AND method = 'report'", (day,))
        for _, r in df.iterrows():
            phone = by_email.get((r["Email"] or "").casefold()) or by_name.get(_norm_name(r["Name"]))
            matched += bool(phone)
            mins = None if r["Minutes"] is None or pd.isna(r["Minutes"]) else float(r["Minutes"])
            conn.execute("""INSERT OR REPLACE INTO meet_stays (tracking_date, phone_number, display_name, email, method,
                                                               minutes, stayed, detail)
                            VALUES (?, ?, ?, ?, 'report', ?, ?, ?)""",
                         (day, phone, r["Name"], r["Email"] or None, mins, int(r["Stayed"]), r["Result"]))
    return matched, len(df)


STATUS_STAYED = "Stayed to the end"
STATUS_EARLY = "Left early"
STATUS_UNCONFIRMED = "Joined, not confirmed"


def meet_day_summary(day: str) -> pd.DataFrame:
    """One row per person for the day, with the best evidence of whether they stayed.
    Order of trust: an admin's mark, then Google's report, then the closing code."""
    people = query_df(
        """SELECT t.phone_number AS phone, MIN(t.member_display_name) AS display,
                  COALESCE(m.full_name, '') AS member, strftime('%H:%M', MIN(t.join_time)) AS joined
           FROM google_meet_tracker t LEFT JOIN members m ON m.phone_number = t.phone_number
           WHERE t.tracking_date = ? GROUP BY t.phone_number ORDER BY MIN(t.join_time)""", (day,))
    stays = query_df("SELECT * FROM meet_stays WHERE tracking_date = ?", (day,))
    rank = {"manual": 0, "report": 1, "code": 2}
    best = {}
    for _, x in stays.iterrows():
        key = x["phone_number"] if isinstance(x["phone_number"], str) and x["phone_number"] else "name:" + _norm_name(x["display_name"])
        if key not in best or rank[x["method"]] < rank[best[key]["method"]]:
            best[key] = x
    out = []

    def status_of(x):
        if x is None:
            return STATUS_UNCONFIRMED, ""
        how = {"manual": "marked by admin", "report": "Google Meet report", "code": "closing code"}[x["method"]]
        mins = "" if x["minutes"] is None or pd.isna(x["minutes"]) else f", {float(x['minutes']):.0f} min"
        return (STATUS_STAYED if x["stayed"] else STATUS_EARLY), f"{how}{mins}"

    for _, p in people.iterrows():
        st_, how = status_of(best.pop(p["phone"], None))
        out.append({"Meet Display Name": p["display"], "Phone": p["phone"], "Member On File": p["member"],
                    "Joined (GMT)": p["joined"], "Status": st_, "How We Know": how})
    for key, x in best.items():
        phone = x["phone_number"] if isinstance(x["phone_number"], str) else ""
        if key.startswith("name:") or phone not in set(people["phone"]):
            st_, how = status_of(x)
            out.append({"Meet Display Name": x["display_name"], "Phone": phone,
                        "Member On File": "", "Joined (GMT)": "", "Status": st_,
                        "How We Know": how + " · joined without the app link"})
    return pd.DataFrame(out, columns=["Meet Display Name", "Phone", "Member On File", "Joined (GMT)", "Status",
                                      "How We Know"])


def meet_daily_totals(days: int = 30) -> pd.DataFrame:
    dates = query_df("""SELECT DISTINCT tracking_date AS d FROM google_meet_tracker WHERE tracking_date >= date('now', ?)
                        UNION SELECT DISTINCT tracking_date FROM meet_stays WHERE tracking_date >= date('now', ?)
                        ORDER BY d""", (f"-{int(days)} days", f"-{int(days)} days"))
    rows = []
    for d in dates["d"]:
        summary = meet_day_summary(d)
        rows.append({"Date": d, "Joined": len(summary),
                     "Stayed to the end": int((summary["Status"] == STATUS_STAYED).sum())})
    return pd.DataFrame(rows, columns=["Date", "Joined", "Stayed to the end"])


def meet_full_export() -> pd.DataFrame:
    return query_df(
        """SELECT tracking_date AS "Date", member_display_name AS "Meet Display Name", phone_number AS "Phone",
                  strftime('%H:%M:%S', join_time) AS "Joined (GMT)"
           FROM google_meet_tracker ORDER BY join_time DESC""")


# ----- members ---------------------------------------------------------------
MEMBER_FIELDS = ("full_name", "phone_number", "whatsapp_number", "email", "country",
                 "emergency_contact_name", "emergency_contact_phone", "parent_guardian_phone",
                 "is_minor", "source", "birth_day", "birth_month")


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


def set_member_birthday(member_id: int, day: int, month: int):
    with closing(get_conn()) as conn, conn:
        conn.execute("UPDATE members SET birth_day = ?, birth_month = ? WHERE id = ?", (day, month, member_id))


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
                  m.birth_day, m.birth_month,
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
                  m.birth_day, m.birth_month,
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
def server_checkin_time(member_id: int, event_id: int) -> str:
    """The arrival time exactly as the database recorded it."""
    with closing(get_conn()) as conn:
        row = conn.execute("""SELECT strftime('%H:%M:%S', check_in_timestamp) AS t FROM attendance
                              WHERE member_id = ? AND event_id = ?""", (member_id, event_id)).fetchone()
        return row["t"] if row else datetime.now(timezone.utc).strftime("%H:%M:%S")


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
                      (SELECT COUNT(*) FROM events WHERE is_open = 1 AND uses_checkin = 1) AS open_events,
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
                  CASE WHEN NOT e.prereg_open THEN 'Off'
                       WHEN e.prereg_deadline IS NOT NULL
                            AND e.prereg_deadline <= strftime('%Y-%m-%d %H:%M', 'now') THEN 'Closed'
                       ELSE 'On' END AS "Pre-Registration",
                  CASE WHEN NOT e.uses_checkin THEN 'Not used'
                       WHEN e.is_open THEN 'Open' ELSE 'Closed' END AS "Check-In"
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
    if not isinstance(intl, str):
        return ""
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
    if not isinstance(iso, str) or not iso:   # empty dates can arrive from pandas as NaN/NA
        return ""
    try:
        return date.fromisoformat(iso).strftime("%a %d %b %Y")
    except ValueError:
        return str(iso)


MODE_BOTH, MODE_CHECKIN, MODE_PREREG = "both", "checkin", "prereg"
MODES = [MODE_BOTH, MODE_CHECKIN, MODE_PREREG]
MODE_LABELS = {
    MODE_BOTH: "Pre-registration and check-in at the venue",
    MODE_CHECKIN: "Check-in at the venue only",
    MODE_PREREG: "Pre-registration only (no check-in at the venue)",
}


def event_mode(ev: dict) -> str:
    if not ev.get("uses_checkin", 1):
        return MODE_PREREG
    return MODE_BOTH if ev.get("prereg_open") else MODE_CHECKIN


def mode_flags(mode: str) -> tuple[bool, bool]:
    """(prereg_open, uses_checkin) for a mode."""
    return mode != MODE_CHECKIN, mode != MODE_PREREG


def mode_input(key: str, current: str = MODE_BOTH):
    return st.selectbox("How members take part", MODES, index=MODES.index(current), format_func=MODE_LABELS.get,
                        key=key, help="Pre-registration only suits programmes where registering is all that's needed. "
                                      "Those events get no check-in QR code, and when registration closes members are "
                                      "simply told it has closed.")


def parse_deadline(value):
    try:
        return datetime.strptime(str(value), "%Y-%m-%d %H:%M") if value else None
    except ValueError:
        return None


def now_local() -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None)   # Ghana time is UTC all year


def prereg_deadline_passed(ev: dict) -> bool:
    deadline = parse_deadline(ev.get("prereg_deadline"))
    return bool(deadline and now_local() >= deadline)


def fmt_deadline(value) -> str:
    """e.g. 'Fri 9 Oct 2026, 6:00 PM'"""
    d = parse_deadline(value)
    if not d:
        return ""
    hour = d.strftime("%I").lstrip("0") or "12"
    return f"{d.strftime('%a')} {d.day} {d.strftime('%b %Y')}, {hour}:{d.strftime('%M %p')}"


def deadline_inputs(key: str, current=None, label="Pre-registration closes"):
    """Optional closing date + time. Returns 'YYYY-MM-DD HH:MM' or None."""
    d = parse_deadline(current)
    key = f"{key}_{(current or 'none').replace(' ', '_')}"   # fresh widgets whenever the saved value changes
    c1, c2 = st.columns(2)
    day = c1.date_input(f"{label} on (optional)", value=d.date() if d else None, format="DD/MM/YYYY",
                        key=f"{key}_dl_date",
                        help="Leave empty to keep pre-registration open until you turn it off yourself.")
    times = [f"{h:02d}:{m:02d}" for h in range(24) for m in (0, 30)] + ["23:59"]
    current_time = d.strftime("%H:%M") if d else "23:59"
    if current_time not in times:
        times = sorted(times + [current_time])
    at = c2.selectbox("At (Ghana time)", times, index=times.index(current_time), key=f"{key}_dl_time",
                      format_func=lambda t: fmt_clock(t) + (" (end of day)" if t == "23:59" else ""))
    return f"{day.isoformat()} {at}" if day else None


def fmt_clock(hhmm: str) -> str:
    t = datetime.strptime(hhmm, "%H:%M")
    return f"{int(t.strftime('%I'))}:{t.strftime('%M %p')}"


def deadline_problem(deadline, event_date):
    if not deadline or not event_date:
        return None
    try:
        if parse_deadline(deadline).date() > date.fromisoformat(event_date):
            return "Pre-registration should close on or before the event date."
    except ValueError:
        return None
    return None


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
.stApp [data-testid="stIconMaterial"], .stApp .stMarkdown [data-testid="stIconMaterial"],
.stApp .stMarkdown h1 [data-testid="stIconMaterial"], .stApp .stMarkdown h2 [data-testid="stIconMaterial"],
.stApp .stMarkdown h3 [data-testid="stIconMaterial"], .stApp .stMarkdown h4 [data-testid="stIconMaterial"],
.stApp .stMarkdown span[role="img"][translate="no"], .stApp .stMarkdown h1 span[role="img"][translate="no"],
.stApp .stMarkdown h2 span[role="img"][translate="no"], .stApp .stMarkdown h3 span[role="img"][translate="no"],
.stApp .stMarkdown h4 span[role="img"][translate="no"] {{
    font-family: 'Material Symbols Rounded' !important; font-weight: 400 !important; }}
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

/* ---- footer centring ---- */
.st-key-ig_footer_btn > div, .st-key-ig_footer_btn [data-testid="stElementContainer"] {{ width:100% !important;
    display:flex !important; justify-content:center !important; }}
.st-key-ig_footer_btn button {{ margin:0 auto; }}
.st-key-ig_go {{ display:none; }}

/* ---- cards made from keyed containers ---- */
[class*="st-key-glass"] {{ background:#fff; border:1px solid var(--line); border-radius:18px; padding:1.1rem 1.2rem;
    box-shadow:0 12px 30px -26px rgba(31,26,51,.45); }}
[class*="st-key-glass_panel"], .st-key-glass_topbar {{ background:transparent; border:none; box-shadow:none; padding:0; }}
.st-key-glass_topbar {{ margin-bottom:.4rem; }}
[class*="st-key-glass_pa_"] {{ padding:.8rem 1rem; margin-bottom:.45rem; }}
[class*="st-key-glass_pa_"] code, .ig-row code {{ color:var(--royal) !important; background:#F1ECF9 !important;
    border-radius:6px; padding:.08rem .4rem; }}
.ig-rowmeta {{ color:var(--muted); font-size:.84rem; margin-top:.2rem; }}
[class*="st-key-wa_"] a {{ background:#1E9E57 !important; border-color:#1E9E57 !important;
    box-shadow:0 10px 22px -14px rgba(30,158,87,.9); }}
[class*="st-key-wa_"] a p {{ color:#fff !important; font-weight:700 !important; }}
a[data-testid="stBaseLinkButton-primary"] {{ background:var(--royal) !important; border-color:var(--royal) !important;
    box-shadow:0 10px 22px -14px rgba(75,46,131,.9); }}
a[data-testid="stBaseLinkButton-primary"] p {{ color:#fff !important; }}

/* ---- green "you're in" card ---- */
.ig-done-success {{ background: radial-gradient(120% 90% at 50% 0%, #245A47 0%, var(--ink) 66%); }}
.ig-done-success .ig-done-seal {{ background:#43C482; color:#0C2A1A; box-shadow:0 0 0 8px rgba(67,196,130,.2); }}
.ig-done-success .ig-done-kicker {{ color:#9BE6BE; }}
.ig-bar {{ height:3px; border-radius:3px; background:rgba(255,255,255,.12); overflow:hidden; margin:.9rem auto 0; max-width:16rem; }}
.ig-bar i {{ display:block; height:100%; background:#43C482; animation: igdrain var(--secs) linear forwards; }}
@keyframes igdrain {{ from {{ width:100%; }} to {{ width:0%; }} }}

/* ---- refusal / waiting card ---- */
.ig-deny {{ background:#fff; border:1px solid #EBCBD1; border-top:4px solid #B3263A; border-radius:18px;
           padding:1.4rem 1.2rem; text-align:center; margin:.6rem 0 1rem; box-shadow:0 14px 32px -26px rgba(31,26,51,.5); }}
.ig-deny-icon {{ width:52px; height:52px; border-radius:50%; margin:0 auto .75rem; display:flex; align-items:center;
                justify-content:center; background:#FBEDEF; color:#B3263A; }}
.ig-deny-icon svg {{ width:25px; height:25px; }}
.ig-deny h3 {{ font-size:1.45rem !important; margin:0 0 .35rem; padding:0; color:var(--ink) !important; }}
.ig-deny p {{ margin:0 auto; max-width:24rem; font-size:.93rem; line-height:1.6; color:var(--muted); }}
.ig-deny.ig-wait {{ border-color:#EADFC4; border-top-color:var(--gold); }}
.ig-wait .ig-deny-icon {{ background:#FBF4E3; color:var(--gold-text); }}

/* ---- masterclass token ---- */
.ig-tokencard {{ background: radial-gradient(120% 90% at 50% 0%, #3A2E66 0%, var(--ink) 62%); color:var(--paper);
                border-radius:24px; padding:1.8rem 1.4rem 1.5rem; text-align:center; margin:.4rem 0 1rem;
                box-shadow:0 26px 54px -30px rgba(31,26,51,.9); }}
.ig-tokencard .ig-kicker {{ color:var(--gold); }}
.ig-token {{ font-family:'Cormorant Garamond', Georgia, serif; font-size:2.6rem; font-weight:700; letter-spacing:.05em;
            color:#fff; margin:.4rem 0 .3rem; font-variant-numeric: lining-nums; }}
.ig-status {{ display:inline-block; font-size:.66rem; font-weight:700; letter-spacing:.16em; text-transform:uppercase;
             padding:.3rem .75rem; border-radius:999px; margin-top:.3rem; }}
.ig-status-ok {{ background:rgba(67,196,130,.16); color:#9BE6BE; border:1px solid rgba(67,196,130,.5); }}
.ig-status-wait {{ background:rgba(201,162,75,.16); color:var(--gold); border:1px solid rgba(201,162,75,.55); }}
.ig-tokencard p {{ color:#D9D3E6; font-size:.95rem; line-height:1.6; max-width:25rem; margin:.8rem auto 0; }}
.ig-tokencard .ig-done-meta {{ margin-top:1.2rem; padding-top:.9rem; border-top:1px solid rgba(255,255,255,.12);
                              font-size:.78rem; color:#A79FBD; }}

/* ---- closing code (admin) ---- */
.ig-code {{ background: radial-gradient(120% 90% at 50% 0%, #3A2E66 0%, var(--ink) 65%); border-radius:20px;
           padding:1.3rem 1rem 1.1rem; text-align:center; color:#fff; margin:.3rem 0 .8rem; }}
.ig-code .n {{ font-family:'Cormorant Garamond', Georgia, serif; font-size:3.6rem; font-weight:700; letter-spacing:.18em;
              line-height:1; color:#fff; font-variant-numeric: lining-nums; }}
.ig-code .k {{ font-size:.66rem; letter-spacing:.24em; text-transform:uppercase; color:var(--gold); font-weight:700; margin-bottom:.5rem; }}
.ig-code .t {{ font-size:.82rem; color:#CFC6E3; margin-top:.5rem; }}
.ig-pillrow {{ display:flex; flex-wrap:wrap; gap:.4rem; margin:.2rem 0 .8rem; }}
.ig-tag {{ display:inline-block; font-size:.72rem; font-weight:700; padding:.2rem .6rem; border-radius:999px; }}
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
    colour = TYPE_COLOURS.get(ev["event_type"], ROYAL)
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


# ---------------------------------------------------------------------------
# Birthdays
# ---------------------------------------------------------------------------
MONTHS = list(range(1, 13))
DEFAULT_BIRTHDAY_MESSAGE = (
    "Happy birthday, {name}! The whole Ignite Prayer Network family is celebrating you today. "
    "May this new year draw you closer to God and be full of His favour. Have a blessed day!"
)


def fmt_birthday(day, month) -> str:
    try:
        return f"{int(day)} {calendar.month_name[int(month)]}"
    except (TypeError, ValueError, IndexError):
        return ""


def birthday_problem(day, month, required: bool):
    if not day and not month:
        return "Choose the day and month of your birthday." if required else None
    if not day or not month:
        return "Choose both the day and the month of your birthday."
    if int(day) > calendar.monthrange(2024, int(month))[1]:   # 2024 is a leap year, so 29 Feb is allowed
        return f"{calendar.month_name[int(month)]} doesn't have {int(day)} days."
    return None


def with_birthday_column(df: pd.DataFrame, after: str) -> pd.DataFrame:
    if df.empty or "birth_day" not in df:
        return df.drop(columns=["birth_day", "birth_month"], errors="ignore")
    df = df.copy()
    bday = [fmt_birthday(d, m) if pd.notna(d) and pd.notna(m) else "" for d, m in zip(df["birth_day"], df["birth_month"])]
    df = df.drop(columns=["birth_day", "birth_month"])
    df.insert(list(df.columns).index(after) + 1 if after in df else len(df.columns), "Birthday", bday)
    return df


def birthday_on(day: int, month: int, year: int) -> date:
    """29 February is celebrated on 28 February in ordinary years."""
    if month == 2 and day == 29 and not calendar.isleap(year):
        day = 28
    return date(year, month, day)


def today_local() -> date:
    return datetime.now(timezone.utc).date()   # Ghana time is UTC all year


def members_with_birthdays() -> list[dict]:
    with closing(get_conn()) as conn:
        return [dict(r) for r in conn.execute(
            """SELECT id, full_name, phone_number, whatsapp_number, country, birth_day, birth_month
               FROM members WHERE birth_day IS NOT NULL AND birth_month IS NOT NULL
               ORDER BY birth_month, birth_day, full_name COLLATE NOCASE""").fetchall()]


def upcoming_birthdays(days: int, start: date | None = None) -> list[dict]:
    """Birthdays from `start` (today) through the next `days` days, soonest first."""
    start = start or today_local()
    out = []
    for m in members_with_birthdays():
        nxt = birthday_on(m["birth_day"], m["birth_month"], start.year)
        if nxt < start:
            nxt = birthday_on(m["birth_day"], m["birth_month"], start.year + 1)
        away = (nxt - start).days
        if away <= days:
            out.append({**m, "date": nxt, "days_away": away})
    return sorted(out, key=lambda m: (m["days_away"], m["full_name"].lower()))


def birthday_counts() -> tuple[int, int]:
    with closing(get_conn()) as conn:
        r = conn.execute("""SELECT SUM(birth_month IS NOT NULL), COUNT(*) FROM members""").fetchone()
        return int(r[0] or 0), int(r[1] or 0)


def birthday_message(name: str) -> str:
    template = get_app_setting("birthday_message") or DEFAULT_BIRTHDAY_MESSAGE
    first = (name or "").strip().split(" ")[0] or "friend"
    return template.replace("{name}", first)


def whatsapp_wish_url(member: dict) -> str:
    number = re.sub(r"\D", "", member.get("whatsapp_number") or member.get("phone_number") or "")
    return f"https://wa.me/{number}?text={quote(birthday_message(member['full_name']))}"


def _ics_text(value: str) -> str:
    return (value.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n"))


def _ics_fold(line: str) -> str:
    raw = line.encode("utf-8")
    if len(raw) <= 74:
        return line
    parts, current = [], b""
    for ch in line:
        b = ch.encode("utf-8")
        if len(current) + len(b) > 73:
            parts.append(current.decode("utf-8"))
            current = b""
        current += b
    parts.append(current.decode("utf-8"))
    return "\r\n ".join(parts)


def birthday_calendar_bytes() -> bytes:
    """An .ics file with every member's birthday as a yearly, all-day event
    and a reminder at 8am on the day. Phones and Google/Outlook calendars
    can import it."""
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Ignite Prayer Network//Birthdays//EN",
             "CALSCALE:GREGORIAN", "METHOD:PUBLISH", "X-WR-CALNAME:Ignite birthdays"]
    for m in members_with_birthdays():
        day, month = int(m["birth_day"]), int(m["birth_month"])
        if month == 2 and day == 29:
            start, rule = "20250228", "RRULE:FREQ=YEARLY;BYMONTH=2;BYMONTHDAY=-1"
        else:
            start, rule = f"2025{month:02d}{day:02d}", "RRULE:FREQ=YEARLY"
        end = (date(int(start[:4]), int(start[4:6]), int(start[6:])) + timedelta(days=1)).strftime("%Y%m%d")
        phone = fmt_phone(m.get("whatsapp_number") or m["phone_number"])
        lines += [
            "BEGIN:VEVENT",
            f"UID:ignite-member-{m['id']}-birthday@igniteprayernetwork",
            f"DTSTAMP:{stamp}",
            f"DTSTART;VALUE=DATE:{start}",
            f"DTEND;VALUE=DATE:{end}",
            rule,
            f"SUMMARY:{_ics_text(m['full_name'] + chr(39) + 's birthday')}",
            f"DESCRIPTION:{_ics_text('Ignite member. WhatsApp: ' + phone + '. Send wishes: ' + whatsapp_wish_url(m))}",
            "TRANSP:TRANSPARENT",
            "BEGIN:VALARM", "ACTION:DISPLAY",
            f"DESCRIPTION:{_ics_text(m['full_name'] + chr(39) + 's birthday today')}",
            "TRIGGER:PT8H", "END:VALARM",
            "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")
    return ("\r\n".join(_ics_fold(l) for l in lines) + "\r\n").encode("utf-8")


_MONTH_LOOKUP = {name.lower()[:3]: i for i, name in enumerate(calendar.month_name) if name}


def parse_birthday(text):
    """Read a birthday from an imported cell: 12/03, 12/03/1990, 1990-03-12,
    12 March, March 12, 12th Mar 1990. Day comes before month in numbers."""
    t = str(text or "").strip().lower()
    if not t or t == "nan":
        return None
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", t)
    if m:
        day, month = int(m.group(3)), int(m.group(2))
    else:
        m = re.match(r"^(\d{1,2})[/.\-](\d{1,2})(?:[/.\-]\d{2,4})?$", t)
        if m:
            day, month = int(m.group(1)), int(m.group(2))
        else:
            words = re.findall(r"[a-z]+|\d+", t)
            month = next((_MONTH_LOOKUP.get(w[:3]) for w in words if w[:3] in _MONTH_LOOKUP), None)
            nums = [int(w) for w in words if w.isdigit() and int(w) <= 31]
            if not month or not nums:
                return None
            day = nums[0]
    if not 1 <= month <= 12 or birthday_problem(day, month, True):
        return None
    return day, month


def birthday_inputs(key_prefix: str, required: bool, day=None, month=None):
    """Day and month pickers. Returns (day, month); either may be None."""
    c1, c2 = st.columns(2)
    star = " *" if required else ""
    d = c1.selectbox(f"Day{star}", list(range(1, 32)), index=(int(day) - 1) if day else None,
                     placeholder="Day", key=f"{key_prefix}_bd")
    m = c2.selectbox(f"Month{star}", MONTHS, index=(int(month) - 1) if month else None,
                     format_func=lambda i: calendar.month_name[i], placeholder="Month", key=f"{key_prefix}_bm")
    return d, m


def reset_flow_state():
    for p in FLOW_PREFIXES:
        st.session_state.pop(f"{p}_pending", None)
        st.session_state.pop(f"{p}_known", None)
        st.session_state.pop(f"{p}_bday", None)


def confirm_and_reset(kind: str, title: str, message: str, context: str,
                      auto_reset: bool = True, link: tuple | None = None, seconds: int = CONFIRM_SECONDS,
                      stamp: str | None = None):
    """Show the confirmation screen, throw away the form and anything the
    person typed (a new form counter means brand-new, empty widgets)."""
    st.session_state.confirmation = {
        "kind": kind, "title": title, "message": message, "context": context,
        "time": stamp or datetime.now(timezone.utc).strftime("%H:%M:%S"), "shown_at": None,
        "auto_reset": auto_reset, "link": link, "seconds": seconds,
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
    seal, css = (ICON_CHECK, "ig-done ig-done-success") if c["kind"] == "success" else (ICON_INFO, "ig-done ig-done-info")
    secs = c.get("seconds", CONFIRM_SECONDS)
    bar = f'<div class="ig-bar" style="--secs:{secs}s"><i></i></div>' if c.get("auto_reset") else ""
    st.markdown(
        f"""<div class="{css}">
              <div class="ig-done-seal">{seal}</div>
              <div class="ig-done-kicker">Submitted · thank you</div>
              <h2>{esc(c['title'])}</h2>
              <p>{esc(c['message'])}</p>
              <div class="ig-done-meta">{esc(c['context'])} &nbsp;·&nbsp; recorded {esc(c['time'])} GMT (server time)</div>
              {bar}
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
            left = secs - int(time.time() - c["shown_at"])
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

        need_bday = source == "Join link"
        st.markdown(f'<div class="ig-section">Your birthday{"" if need_bday else " (optional)"}</div>',
                    unsafe_allow_html=True)
        b_day, b_month = birthday_inputs(k("birthday"), need_bday)
        st.caption("Just the day and month, so the family can celebrate with you.")

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
    bday_error = birthday_problem(b_day, b_month, need_bday)
    if bday_error:
        errors.append(bday_error)
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
        if b_day and b_month and not (get_member(existing) or {}).get("birth_month"):
            set_member_birthday(existing, b_day, b_month)
        return existing
    return create_member({
        "full_name": full_name.strip(), "phone_number": pending["phone"], "whatsapp_number": whatsapp,
        "email": email.strip() or None, "country": lives_in if lives_in != "Other country" else None,
        "emergency_contact_name": ec_name.strip(), "emergency_contact_phone": ec_phone,
        "parent_guardian_phone": parent or None, "is_minor": int(is_minor),
        "source": source, "birth_day": b_day, "birth_month": b_month,
    })


def page_top():
    """Flyer slot at the very top, then the brand. Returns the flyer slot."""
    banner = st.container(key="flyer_banner")
    brand_row()
    return banner


FLYER_LOOKS = {
    "frosted": "Frosted: the flyer glows softly behind frosted forms (recommended)",
    "rich": "Rich: more of the flyer's colour shows through",
    "off": "Off: plain ivory page, flyer only at the top",
}
FLYER_OVERLAY = {"frosted": (.70, .84, .93), "rich": (.42, .64, .84)}


@st.cache_data(ttl=3600, show_spinner=False, max_entries=50)
def flyer_backdrop_uri(event_id: int, fingerprint: str):
    """A small, blurred copy of the flyer to sit behind the page (about 30-60 KB)."""
    flyer = get_flyer(event_id)
    if not flyer:
        return None
    img = Image.open(io.BytesIO(flyer)).convert("RGB")
    img.thumbnail((640, 640))
    img = img.filter(ImageFilter.GaussianBlur(radius=12))
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=72, optimize=True)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()


def flyer_backdrop(ev: dict):
    """Members see the programme's flyer as a soft, see-through backdrop with the
    forms frosted on top of it."""
    look = get_app_setting("flyer_look", "frosted")
    flyer = get_flyer(ev["id"])
    if look not in FLYER_OVERLAY or not flyer:
        return
    uri = flyer_backdrop_uri(ev["id"], hashlib.md5(flyer).hexdigest())
    if not uri:
        return
    a1, a2, a3 = FLYER_OVERLAY[look]
    st.markdown(f"""<style>
.stApp {{ background: var(--paper) !important; }}
.stApp::before {{ content:""; position:fixed; inset:0; z-index:0; pointer-events:none;
    background:
      linear-gradient(180deg, rgba(251,248,243,{a1}) 0%, rgba(251,248,243,{a2}) 42%, rgba(251,248,243,{a3}) 100%),
      url("{uri}") center top / cover no-repeat;
    transform: scale(1.06); }}
[data-testid="stAppViewContainer"], [data-testid="stMain"], header[data-testid="stHeader"] {{
    background: transparent !important; position: relative; z-index: 1; }}
[data-testid="stForm"], .ig-ticket, .ig-note, .ig-deny, [class*="st-key-glass"]:not([class*="st-key-glass_panel"]) {{
    background: rgba(255,255,255,.66) !important;
    -webkit-backdrop-filter: blur(18px) saturate(160%); backdrop-filter: blur(18px) saturate(160%);
    border: 1px solid rgba(255,255,255,.75) !important;
    box-shadow: inset 0 1px 0 rgba(255,255,255,.8), 0 22px 44px -26px rgba(31,26,51,.55) !important; }}
.ig-ticket-body {{ border-left-color: rgba(31,26,51,.14); }}
[data-baseweb="input"], [data-baseweb="select"] > div, [data-baseweb="textarea"] {{
    background: rgba(255,255,255,.78) !important; }}
.st-key-flyer_banner img {{ box-shadow: 0 0 0 1px rgba(255,255,255,.7), 0 26px 50px -24px rgba(31,26,51,.7) !important; }}
.ig-lead, .ig-word-sub {{ text-shadow: 0 1px 0 rgba(255,255,255,.6); }}
</style>""", unsafe_allow_html=True)


def show_flyer(banner, ev: dict):
    flyer = get_flyer(ev["id"])
    if flyer:
        banner.image(flyer, width="stretch")
        flyer_backdrop(ev)


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
        confirm_and_reset("success", "Welcome. You're checked in.", msg, context,
                          seconds=CHECKIN_CONFIRM_SECONDS, stamp=server_checkin_time(member_id, ev["id"]))
    else:
        confirm_and_reset("info", "You're already checked in",
                          "Your arrival was recorded earlier, so there's nothing more to do. Enjoy the programme.", context,
                          seconds=CHECKIN_CONFIRM_SECONDS + 1, stamp=server_checkin_time(member_id, ev["id"]))


def checkin_page():
    if show_confirmation():
        return
    banner = page_top()
    ev = resolve_checkin_event()
    if ev is None:
        return
    show_flyer(banner, ev)
    if not ev.get("uses_checkin", 1):
        page_heading("Check-in", "No check-in for this programme",
                     "This programme uses pre-registration only, so there's nothing to do here.")
        event_ticket(ev)
        return
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
def prereg_closed_note(ev: dict):
    ended = f'It ended {esc(fmt_deadline(ev["prereg_deadline"]))}.'
    extra = (" You can still join us: on the day, scan the QR code at the entrance to check in."
             if ev.get("uses_checkin", 1) else " Thank you for your interest.")
    st.markdown(f'<div class="ig-note"><strong>Pre-registration for this programme has closed.</strong> {ended}{extra}'
                '</div>', unsafe_allow_html=True)


def complete_prereg(member_id: int, ev: dict):
    fresh = get_event(ev["id"]) or ev          # the deadline may have passed while they were typing
    if not fresh["prereg_open"] or prereg_deadline_passed(fresh):
        late = ("Registration for this programme closed before your details came in. You're still welcome: "
                "on the day, scan the QR code at the entrance to check in." if fresh.get("uses_checkin", 1) else
                "Registration for this programme closed before your details came in, so we couldn't add you.")
        confirm_and_reset("info", "Pre-registration has closed", late, ev["event_name"], auto_reset=False)
        return
    if record_pre_registration(member_id, ev["id"]):
        done = ("On the day, scan the QR code at the entrance and enter your phone number to check in."
                if fresh.get("uses_checkin", 1) else
                "You're on the list, and there's nothing else you need to do. We look forward to seeing you.")
        confirm_and_reset("success", "Your place is reserved", done,
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
    if prereg_deadline_passed(ev):
        prereg_closed_note(ev)
        return
    if ev.get("prereg_deadline"):
        st.markdown(f'<div class="ig-note"><strong>Pre-registration closes {esc(fmt_deadline(ev["prereg_deadline"]))}'
                    '</strong> (Ghana time).</div>', unsafe_allow_html=True)

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
    if ev.get("uses_checkin", 1):
        st.markdown('<div class="ig-note">This reserves your place. It isn\'t your check-in: on the day, '
                    'scan the QR code at the entrance to confirm you\'ve arrived.</div>', unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Public page: member registration (shared in the WhatsApp group)
# ---------------------------------------------------------------------------
def join_birthday_step(member_id: int, link):
    """For people already on the list who haven't given a birthday yet.
    Their name is not shown, so a typed-in number reveals nothing."""
    nonce = st.session_state.form_nonce
    c1, c2 = st.columns([3, 2])
    c1.markdown('<div class="ig-step">Almost done</div>', unsafe_allow_html=True)
    if c2.button("Use a different number", key=f"join_bday_change_{nonce}", type="tertiary"):
        st.session_state.pop("join_bday", None)
        st.rerun()
    with st.form(key=f"join_bday_form_{nonce}"):
        st.markdown('<div class="ig-formtitle">Add your birthday</div><div class="ig-formnote">This number is '
                    "already on our member list. Tell us your birthday so the family can celebrate with you."
                    "</div>", unsafe_allow_html=True)
        b_day, b_month = birthday_inputs(f"join_bday_{nonce}", True)
        go = st.form_submit_button("Save my birthday", type="primary", width="stretch")
    if not go:
        return
    problem = birthday_problem(b_day, b_month, True)
    if problem:
        show_errors([problem])
        return
    set_member_birthday(member_id, b_day, b_month)
    confirm_and_reset("success", "Thank you",
                      "Your birthday is saved. You're all set on the Ignite member list.",
                      "Membership", auto_reset=False, link=link)


def join_page():
    if show_confirmation():
        return
    brand_row()
    page_heading("Membership", "Join the Ignite family",
                 "Register once so we know you're part of the family, and check in faster whenever you join us.")
    link = whatsapp_community_link()
    known = st.session_state.get("join_known")
    pending = st.session_state.get("join_pending")
    bday_for = st.session_state.get("join_bday")

    if known:
        # Number already registered. Don't reveal whose it is.
        st.session_state.pop("join_known", None)
        confirm_and_reset("info", "You're already registered",
                          "This number is already on our member list, so there's nothing more to do. Thank you.",
                          "Membership", auto_reset=False, link=link)

    if bday_for:
        join_birthday_step(bday_for, link)
        return

    if pending is None:
        found = phone_step("join", "Start with your phone number",
                           "Members abroad: choose your country first. We'll ask for a few details next.", "Continue")
        if found:
            member_id = find_member_id(found["phone"])
            if member_id and not (get_member(member_id) or {}).get("birth_month"):
                st.session_state.join_bday = member_id
            elif member_id:
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
# Public pages: Prophet Masterclass, live room gate, Google Meet
# ---------------------------------------------------------------------------
ICON_LOCK = ('<svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M12 2a5 5 0 0 1 5 5v3h1a2 2 0 0 1 2 2v8'
             'a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2v-8a2 2 0 0 1 2-2h1V7a5 5 0 0 1 5-5zm0 2a3 3 0 0 0-3 3v3h6V7a3 3 0 0 0-3-3z'
             'm0 9.5a1.75 1.75 0 0 0-1 3.19V18h2v-1.31a1.75 1.75 0 0 0-1-3.19z"/></svg>')
ICON_CLOCK = ('<svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M12 2a10 10 0 1 1 0 20 10 10 0 0 1 0-20z'
              'm0 2a8 8 0 1 0 0 16 8 8 0 0 0 0-16zm1 3v4.6l3.2 1.9-1 1.7L11 12.7V7h2z"/></svg>')


def deny_card(title: str, message: str, waiting: bool = False):
    st.markdown(f"""<div class="ig-deny{' ig-wait' if waiting else ''}">
                      <div class="ig-deny-icon">{ICON_CLOCK if waiting else ICON_LOCK}</div>
                      <h3>{esc(title)}</h3><p>{esc(message)}</p></div>""", unsafe_allow_html=True)


def redirect_browser(url: str, label: str, key: str):
    """Send the whole browser window to the room. A same-tab button stays on
    screen in case the phone blocks the automatic jump."""
    target = json.dumps(url)
    script = f"""<script>
      (function () {{
        const url = {target};
        try {{
          const d = window.parent.document;
          const a = d.createElement('a');
          a.href = url; a.target = '_top'; a.rel = 'noopener';
          d.body.appendChild(a); a.click();
          return;
        }} catch (e) {{}}
        try {{ window.top.location.href = url; }} catch (e) {{}}
      }})();
    </script>"""
    with st.container(key="ig_go"):
        if hasattr(st, "iframe"):
            st.iframe(script, height=1)
        else:
            st.components.v1.html(script, height=0)
    st.markdown(f'<a href="{html.escape(url, quote=True)}" target="_top" style="text-decoration:none">'
                f'<div class="ig-muted" style="margin:.3rem 0 .6rem">Not moving? Tap the button below.</div></a>',
                unsafe_allow_html=True)
    st.link_button(label, url, type="primary", width="stretch")


def pa_whatsapp_number() -> str:
    return re.sub(r"\D", "", get_app_setting("pa_whatsapp") or get_setting("PA_WHATSAPP", DEFAULT_PA_WHATSAPP))


def pa_whatsapp_url(reg: dict, session_name: str) -> str:
    text = (f"Hello, I've registered for the Prophet Masterclass ({session_name}).\n"
            f"Name: {reg['full_name']}\nRegistration token: {reg['reg_id']}\n"
            f"I'm sending my proof of payment so my access can be approved. Thank you.")
    number = pa_whatsapp_number()
    return f"https://wa.me/{number}?text={quote(text)}" if number else f"https://wa.me/?text={quote(text)}"


def live_link(base: str) -> str:
    return f"{base}/?mode=live"


def masterclass_link(base: str, session_id=None) -> str:
    return f"{base}/?mode=masterclass" + (f"&session={session_id}" if session_id else "")


def meet_link(base: str) -> str:
    return f"{base}/?mode=meet"


def session_ticket(sess: dict):
    paid = sess["session_type"] == "Paid"
    when = ""
    if sess.get("session_date"):
        try:
            when = date.fromisoformat(sess["session_date"]).strftime("%a %d %b %Y")
        except ValueError:
            when = ""
    sub_line = " · ".join(x for x in [when, "Paid session" if paid else "Free session"] if x)
    st.markdown(
        f"""<div class="ig-ticket">
              <div class="ig-ticket-date">{ICON_FLAME}</div>
              <div class="ig-ticket-body">
                <div class="ig-ticket-type" style="color:{AMBER if paid else GREEN}">Prophet Masterclass</div>
                <div class="ig-ticket-name">{esc(sess['session_name'])}</div>
                <div class="ig-ticket-venue">{esc(sub_line)}</div>
              </div>
            </div>""", unsafe_allow_html=True)


def token_card(reg: dict, sess: dict, is_new: bool):
    approved = reg["payment_status"] in MC_APPROVED
    status = ('<span class="ig-status ig-status-ok">Access approved</span>' if approved
              else '<span class="ig-status ig-status-wait">Pending verification</span>')
    if approved:
        note = ("Keep this token. When the masterclass starts, open the live room and enter it "
                "(or your phone number) to go straight in.")
    else:
        note = ("Your place is held. Send your proof of payment to the PA on WhatsApp with this token. "
                "Once it's confirmed, the same token opens the live room.")
    head = "You're registered" if is_new else "You're already registered"
    st.markdown(f"""<div class="ig-tokencard">
                      <div class="ig-kicker">{esc(head)}</div>
                      <div class="ig-token">{esc(reg['reg_id'])}</div>
                      {status}
                      <p>{esc(note)}</p>
                      <div class="ig-done-meta">{esc(sess['session_name'])} · {esc(first_name(reg['full_name']))}</div>
                    </div>""", unsafe_allow_html=True)


def masterclass_page():
    brand_row()
    result = st.session_state.get("mc_result")
    if result:
        sess = mc_session(result["session_id"]) or {"session_name": "Prophet Masterclass", "session_type": "Free",
                                                    "price_note": None}
        reg = result["reg"]
        token_card(reg, sess, result["is_new"])
        if reg["payment_status"] == MC_PENDING:
            if sess.get("price_note"):
                st.markdown(f'<div class="ig-note"><b>How to pay:</b> {esc(sess["price_note"])}</div>',
                            unsafe_allow_html=True)
            with st.container(key="wa_btn"):
                st.link_button("Send payment proof to the PA on WhatsApp", pa_whatsapp_url(reg, sess["session_name"]),
                               width="stretch", icon=":material/chat:")
        else:
            base = current_base_url()
            if base_url_ok(base):
                st.link_button("Go to the live room", live_link(base), type="primary", width="stretch")
        if st.button("Done", width="stretch", key="mc_done"):
            st.session_state.pop("mc_result", None)
            st.session_state.form_nonce += 1
            st.rerun()
        return

    sessions = mc_sessions(active_only=True)
    page_heading("Prophet Masterclass", "Reserve your seat",
                 "Register once and you'll get a personal token. It's your key to the live room.")
    if not sessions:
        deny_card("Enrolment isn't open yet", "There's no masterclass taking registrations right now. "
                  "Please check back soon or watch the WhatsApp group for the link.", waiting=True)
        return
    wanted = st.query_params.get("session")
    by_id = {x["id"]: x for x in sessions}
    if wanted and str(wanted).isdigit() and int(wanted) in by_id:
        sess = by_id[int(wanted)]
    elif len(sessions) == 1:
        sess = sessions[0]
    else:
        sess = by_id[st.selectbox("Choose a session", list(by_id),
                                  format_func=lambda i: f"{by_id[i]['session_name']} · {by_id[i]['session_type']}")]
    session_ticket(sess)
    if sess["session_type"] == "Paid" and sess.get("price_note"):
        st.markdown(f'<div class="ig-note"><b>This is a paid session.</b> {esc(sess["price_note"])}</div>',
                    unsafe_allow_html=True)

    nonce = st.session_state.form_nonce
    k = lambda n: f"mc_{n}_{nonce}"
    with st.form(key=k("form")):
        st.markdown('<div class="ig-formtitle">Your details</div>'
                    '<div class="ig-formnote">Use the phone number you\'ll have with you on the day.</div>',
                    unsafe_allow_html=True)
        name = st.text_input("Full name *", placeholder="e.g. Ama Serwaa Mensah", key=k("name"))
        email = st.text_input("Email *", placeholder="you@example.com", key=k("email"))
        c1, c2 = st.columns([1, 1.5])
        country = c1.selectbox("Country", COUNTRY_NAMES, format_func=country_label, key=k("cc"))
        raw = c2.text_input("Phone number *", placeholder="e.g. 024 123 4567", autocomplete="tel", key=k("ph"))
        go = st.form_submit_button("Register & get my token", type="primary", width="stretch")
    if not go:
        return
    phone = to_intl(raw, country)
    errors = []
    if len(name.strip()) < 2:
        errors.append("Enter your full name.")
    if not is_valid_email(email):
        errors.append("Enter a valid email address.")
    problem = phone_problem(phone) if raw.strip() else "Enter your phone number."
    if country == "Other country" and raw.strip() and not raw.strip().startswith(("+", "00")):
        problem = "For other countries, start with + and the country code."
    if problem:
        errors.append(problem)
    fresh = mc_session(sess["id"])
    if not fresh or not fresh["active_status"]:
        errors.append("Enrolment for this session has just closed.")
    if errors:
        show_errors(errors)
        return
    reg, is_new = mc_register(fresh, name, email, phone)
    st.session_state.mc_result = {"reg": reg, "session_id": fresh["id"], "is_new": is_new}
    st.session_state.form_nonce += 1
    st.rerun()


def live_gate_page():
    brand_row()
    page_heading("Join Live Masterclass Room", "Enter the room",
                 "Type your registration token or the phone number you registered with.")
    go_to = st.session_state.get("live_go")
    if go_to:
        st.markdown(f"""<div class="ig-done ig-done-success">
                          <div class="ig-done-seal">{ICON_CHECK}</div>
                          <div class="ig-done-kicker">Access granted</div>
                          <h2>Welcome, {esc(go_to['first'])}</h2>
                          <p>Taking you into {esc(go_to['session'])} now…</p></div>""", unsafe_allow_html=True)
        redirect_browser(go_to["url"], "Open the live room", "live")
        if st.button("Back", key="live_back", width="stretch"):
            st.session_state.pop("live_go", None)
            st.rerun()
        return

    locked = st.session_state.get("gate_locked_until", 0)
    if locked > time.time():
        deny_card("Too many tries", f"Please wait {int(locked - time.time())} seconds and try again.")
        return

    nonce = st.session_state.form_nonce
    with st.form(key=f"gate_form_{nonce}"):
        c1, c2 = st.columns([1, 1.6])
        country = c1.selectbox("Country (for phone numbers)", COUNTRY_NAMES, format_func=country_label,
                               key=f"gate_cc_{nonce}")
        entry = c2.text_input("Token or phone number", placeholder="IGNITE-8921 or 024 123 4567",
                              key=f"gate_in_{nonce}")
        go = st.form_submit_button("Join live room", type="primary", width="stretch", icon=":material/lock_open:")

    choice = st.session_state.get("gate_choices")
    if go:
        st.session_state.pop("gate_choices", None)
        choice = None
        if not entry.strip():
            show_errors(["Enter your token or phone number."])
            return
        matches = mc_lookup(entry, country)
        approved = [m for m in matches if m["payment_status"] in MC_APPROVED]
        if not approved:
            st.session_state.gate_fails = st.session_state.get("gate_fails", 0) + 1
            if st.session_state.gate_fails >= MAX_LOGIN_ATTEMPTS:
                st.session_state.gate_locked_until = time.time() + LOCKOUT_SECONDS
                st.session_state.gate_fails = 0
            time.sleep(0.8)
            if matches:   # registered but not paid yet
                m = matches[0]
                deny_card("Payment not confirmed yet",
                          f"Your token {m['reg_id']} is registered, but the PA hasn't confirmed your payment. "
                          "Send your proof of payment on WhatsApp and try again once it's approved.", waiting=True)
                with st.container(key="wa_gate"):
                    st.link_button("Message the PA on WhatsApp", pa_whatsapp_url(m, m["session_name"]),
                                   width="stretch", icon=":material/chat:")
            else:
                deny_card("Access denied", "We couldn't find an approved registration for that token or number on an "
                          "open session. Check for typos, or register first.")
                base = current_base_url()
                if base_url_ok(base):
                    st.link_button("Register for the masterclass", masterclass_link(base), width="stretch")
            return
        st.session_state.gate_fails = 0
        if len(approved) == 1:
            choice = approved
        else:
            st.session_state.gate_choices = approved
            choice = approved
    if not choice:
        return
    if len(choice) > 1:
        st.markdown('<div class="ig-formtitle" style="margin-top:.8rem">Which room?</div>', unsafe_allow_html=True)
    for m in choice:
        if len(choice) > 1 and not st.button(m["session_name"], key=f"pick_{m['reg_id']}", width="stretch"):
            continue
        st.session_state.pop("gate_choices", None)
        if not (m.get("streaming_url") or "").startswith("http"):
            deny_card("The room isn't open yet", "You're approved, but the stream link hasn't been added. "
                      "Please try again in a few minutes.", waiting=True)
            return
        mc_log_entry(m["reg_id"], m["session_id"])
        st.session_state.live_go = {"url": m["streaming_url"], "session": m["session_name"],
                                    "first": first_name(m["full_name"])}
        st.session_state.form_nonce += 1
        st.rerun()


def meet_confirm_form(code_state: dict):
    """End of prayer: members confirm they stayed by typing the closing code."""
    done = st.session_state.get("meet_stayed")
    if done:
        st.markdown(f"""<div class="ig-done ig-done-success">
                          <div class="ig-done-seal">{ICON_CHECK}</div>
                          <div class="ig-done-kicker">Recorded · {esc(done['time'])} GMT</div>
                          <h2>Thank you, {esc(done['first'])}</h2>
                          <p>You're marked as having prayed with us to the end. God bless you.</p></div>""",
                    unsafe_allow_html=True)
        if st.button("Done", key="meet_stay_done", width="stretch"):
            st.session_state.pop("meet_stayed", None)
            st.rerun()
        return
    locked = st.session_state.get("meet_code_locked", 0)
    if locked > time.time():
        deny_card("Too many tries", f"Please wait {int(locked - time.time())} seconds and try again.")
        return
    nonce = st.session_state.form_nonce
    with st.form(key=f"meet_stay_form_{nonce}"):
        st.markdown('<div class="ig-step">Prayer has ended</div><div class="ig-formtitle">Confirm you stayed</div>'
                    '<div class="ig-formnote">Enter the number you joined with today and the closing code '
                    'announced at the end of prayer.</div>', unsafe_allow_html=True)
        c1, c2 = st.columns([1, 1.5])
        country = c1.selectbox("Country", COUNTRY_NAMES, format_func=country_label, key=f"meet_scc_{nonce}")
        raw = c2.text_input("Phone number *", placeholder="e.g. 024 123 4567", autocomplete="tel", key=f"meet_sph_{nonce}")
        code = st.text_input("Closing code *", placeholder="4 digits", max_chars=4, key=f"meet_scode_{nonce}")
        go = st.form_submit_button("Confirm I stayed", type="primary", width="stretch", icon=":material/task_alt:")
    if not go:
        return
    phone = to_intl(raw, country)
    if not raw.strip() or phone_problem(phone):
        show_errors([phone_problem(phone) or "Enter your phone number."])
        return
    fresh = meet_code_state()
    if not fresh["active"]:
        deny_card("The closing code has expired", "It's only open for a few minutes at the end of prayer. "
                  "If you stayed, let one of the team know and they can mark you.", waiting=True)
        return
    day = today_local().isoformat()
    name = meet_joined_name(phone, day)
    if code.strip() != fresh["code"]:
        st.session_state.meet_code_fails = st.session_state.get("meet_code_fails", 0) + 1
        if st.session_state.meet_code_fails >= MAX_LOGIN_ATTEMPTS:
            st.session_state.meet_code_locked = time.time() + LOCKOUT_SECONDS
            st.session_state.meet_code_fails = 0
        time.sleep(0.6)
        deny_card("That code isn't right", "Check the code announced at the end of prayer and try again.")
        return
    if not name:
        deny_card("We don't have you joining today", "This number didn't join today's prayer through the Ignite "
                  "link. Next time, join from this page so your attendance counts.")
        return
    meet_record_stay(day, phone, name, "code")
    st.session_state.meet_code_fails = 0
    st.session_state.meet_stayed = {"first": first_name(name), "time": datetime.now(timezone.utc).strftime("%H:%M")}
    st.session_state.form_nonce += 1
    st.rerun()


def meet_page():
    brand_row()
    url = get_app_setting("meet_url").strip()
    is_open = get_app_setting("meet_open", "1") == "1"
    code_state = meet_code_state()
    go_to = st.session_state.get("meet_go")
    if go_to:
        page_heading("Daily prayer room", "Join us on Google Meet")
        st.markdown(f"""<div class="ig-done ig-done-success">
                          <div class="ig-done-seal">{ICON_CHECK}</div>
                          <div class="ig-done-kicker">Arrival recorded · {esc(go_to['time'])} GMT</div>
                          <h2>Welcome, {esc(go_to['first'])}</h2>
                          <p>Opening the Google Meet prayer room… Please come back to this page at the end of
                             prayer to confirm you stayed.</p></div>""", unsafe_allow_html=True)
        redirect_browser(go_to["url"], "Open Google Meet", "meet")
        if st.button("Back", key="meet_back", width="stretch"):
            st.session_state.pop("meet_go", None)
            st.rerun()
        return
    if code_state["active"] or st.session_state.get("meet_stayed"):
        page_heading("Daily prayer room", "Thank you for praying with us",
                     "Before you go, confirm you stayed to the end.")
        meet_confirm_form(code_state)
        if url.startswith("https://") and is_open:
            with st.expander("Only joining now?"):
                meet_join_form(url)
        return
    page_heading("Daily prayer room", "Join us on Google Meet",
                 "Tell us who you are, then we'll take you straight into the room.")
    if not url.startswith("https://") or not is_open:
        deny_card("The prayer room isn't open right now", "Please check the WhatsApp group for today's time "
                  "and come back then.", waiting=True)
        return
    meet_join_form(url)
    st.markdown('<div class="ig-note">At the end of prayer a closing code is announced. Come back to this page '
                'and enter it, so we know you stayed with us to the end.</div>', unsafe_allow_html=True)


def meet_join_form(url: str):
    nonce = st.session_state.form_nonce
    with st.form(key=f"meet_form_{nonce}"):
        st.markdown('<div class="ig-formnote">Use the exact name you show on Google Meet, so we can match you '
                    'in the room.</div>', unsafe_allow_html=True)
        display = st.text_input("Your Google Meet display name *", placeholder="e.g. Ama Mensah", key=f"meet_nm_{nonce}")
        c1, c2 = st.columns([1, 1.5])
        country = c1.selectbox("Country", COUNTRY_NAMES, format_func=country_label, key=f"meet_cc_{nonce}")
        raw = c2.text_input("Phone number *", placeholder="e.g. 024 123 4567", autocomplete="tel", key=f"meet_ph_{nonce}")
        go = st.form_submit_button("Join room", type="primary", width="stretch", icon=":material/videocam:")
    if not go:
        return
    phone = to_intl(raw, country)
    errors = []
    if len(display.strip()) < 2:
        errors.append("Enter your Google Meet display name.")
    problem = phone_problem(phone) if raw.strip() else "Enter your phone number."
    if country == "Other country" and raw.strip() and not raw.strip().startswith(("+", "00")):
        problem = "For other countries, start with + and the country code."
    if problem:
        errors.append(problem)
    if errors:
        show_errors(errors)
        return
    meet_log(display, phone)
    st.session_state.meet_go = {"url": url, "first": first_name(display),
                                "time": datetime.now(timezone.utc).strftime("%H:%M")}
    st.session_state.form_nonce += 1
    st.rerun()

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
def birthday_row(member: dict, when: str):
    left, right = st.columns([3, 2], vertical_alignment="center")
    where = f" · {member['country']}" if member.get("country") else ""
    left.markdown(f"**{esc(member['full_name'])}**  \n"
                  f"<span style='color:{MUTED};font-size:.88rem'>{when} · {esc(fmt_phone(member.get('whatsapp_number') or member['phone_number']))}{esc(where)}</span>",
                  unsafe_allow_html=True)
    right.link_button("Send wishes on WhatsApp", whatsapp_wish_url(member), icon=":material/cake:", width="stretch")


def birthdays_today_panel():
    soon = upcoming_birthdays(7)
    today = [m for m in soon if m["days_away"] == 0]
    later = [m for m in soon if m["days_away"] > 0]
    if not soon:
        return
    with st.container(border=True):
        if today:
            st.markdown(f"#### :material/cake: Birthday{'s' if len(today) > 1 else ''} today")
            for m in today:
                birthday_row(m, "Today")
        else:
            st.markdown("#### :material/cake: Birthdays this week")
        if later:
            if today:
                st.markdown("**Coming up this week**")
            for m in later:
                when = "Tomorrow" if m["days_away"] == 1 else m["date"].strftime("%A %d %B").replace(" 0", " ")
                birthday_row(m, when)


def birthdays_section():
    have, total = birthday_counts()
    c1, c2, c3 = st.columns(3)
    c1.metric("Birthdays on file", have)
    c2.metric("Still missing", total - have)
    c3.metric("This month", sum(1 for m in members_with_birthdays() if m["birth_month"] == today_local().month))

    st.markdown("#### Get a reminder on your phone")
    st.markdown("Download this calendar file and open it on your phone (or import it into Google Calendar or "
                "Outlook). Every member's birthday shows up as a yearly event with a reminder at 8am on the day, "
                "and the event has a link to send wishes on WhatsApp. Download it again every few weeks to pick "
                "up new members.")
    st.download_button("Download birthday calendar (.ics)", birthday_calendar_bytes(),
                       file_name="Ignite_birthdays.ics", mime="text/calendar", type="primary",
                       icon=":material/event:", disabled=have == 0, key="bday_ics")
    st.caption("The file contains names and phone numbers, so keep it to the team.")

    st.markdown("#### Birthdays by month")
    month = st.selectbox("Month", MONTHS, index=today_local().month - 1,
                         format_func=lambda i: calendar.month_name[i], key="bday_month")
    rows = [m for m in members_with_birthdays() if m["birth_month"] == month]
    if not rows:
        st.caption(f"No birthdays recorded for {calendar.month_name[month]} yet.")
    else:
        for m in rows:
            birthday_row(m, fmt_birthday(m["birth_day"], m["birth_month"]))

    st.markdown("#### Birthday message")
    with st.form("bday_msg_form"):
        text = st.text_area("Message that opens in WhatsApp", value=get_app_setting("birthday_message") or DEFAULT_BIRTHDAY_MESSAGE,
                            height=110, help="{name} is replaced with the member's first name.")
        c1, c2 = st.columns(2)
        saved = c1.form_submit_button("Save message", type="primary")
        reset = c2.form_submit_button("Use the default message")
    if saved or reset:
        set_app_setting("birthday_message", "" if reset else text.strip())
        notify("Birthday message updated.")
        st.rerun()

    if total - have:
        st.info("To collect birthdays from people already on the list, post the membership link again. Anyone who "
                "is already registered only gets asked for their birthday.", icon=":material/lightbulb:")


def overview_tab():
    birthdays_today_panel()
    s = overview_stats()
    c = st.columns(6)
    c[0].metric("Members", s["members"])
    c[1].metric("Living abroad", s["abroad"])
    c[2].metric("Events", s["events"], help=f"{s['open_events']} open for check-in")
    c[3].metric("Pre-registrations", s["preregs"])
    c[4].metric("Check-ins", s["checkins"])
    c[5].metric("Checked in today", s["today"])
    if DB_IS_PERSISTENT:
        st.success(f"Data is saved on the permanent disk ({DB_PATH}). Restarts and updates don't erase it. "
                   "A backup now and then is still wise.", icon=":material/verified:")
    else:
        st.warning("This copy isn't using a permanent disk, so a restart could clear the data. On Render, attach a "
                   "disk at /data. Until then, download a backup from the **Backup** tab after every event.",
                   icon=":material/backup:")
    mc = query_df("""SELECT
            (SELECT COUNT(*) FROM masterclass_registrations WHERE payment_status = 'Pending Verification') AS pending,
            (SELECT COUNT(*) FROM masterclass_registrations) AS mc_total,
            (SELECT COUNT(DISTINCT phone_number) FROM google_meet_tracker WHERE tracking_date = date('now')) AS meet_today
        """).iloc[0]
    m = st.columns(3)
    m[0].metric("Masterclass registrations", int(mc["mc_total"]))
    m[1].metric("Payments waiting for the PA", int(mc["pending"]))
    m[2].metric("In the Google Meet today", int(mc["meet_today"]))
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
            c5, c6 = st.columns([3, 2], vertical_alignment="bottom")
            with c5:
                mode = mode_input("new_event_mode")
            is_open = c6.toggle("Open day-of check-in", value=True,
                                help="Ignored for pre-registration-only events.")
            deadline = deadline_inputs("new_event")
            create = st.form_submit_button("Create event", type="primary")
        if create:
            name = ev_name.strip() or f"{ev_type} {datetime.now():%Y-%m-%d}"
            prereg_open, uses_checkin = mode_flags(mode)
            problem = deadline_problem(deadline, ev_date.isoformat() if ev_date else None)
            if problem:
                st.error(f"{problem} The event was not created.")
                return
            try:
                flyer = prepare_flyer(flyer_file) if flyer_file is not None else None
            except ValueError as err:
                st.error(f"{err} The event was not created.")
                return
            new_id = create_event(name, ev_type, ev_date.isoformat() if ev_date else None,
                                  venue, flyer, is_open, prereg_open, deadline if prereg_open else None,
                                  uses_checkin)
            st.session_state.manage_event = new_id
            notify(f"Created “{name}”. " + {
                MODE_BOTH: "Its registration link and check-in QR code are ready below.",
                MODE_CHECKIN: "It's attendance-only; its check-in QR code is ready below.",
                MODE_PREREG: "It's pre-registration only; its registration link is ready below.",
            }[mode])
            st.rerun()


def prereg_section(ev: dict, base):
    c1, c2 = st.columns([3, 1])
    if ev["prereg_open"] and prereg_deadline_passed(ev):
        c1.markdown(f":red[●] **Pre-registration closed** {fmt_deadline(ev['prereg_deadline'])} · "
                    f"{ev['preregs']} registered")
    elif ev["prereg_open"]:
        closes = f" · closes {fmt_deadline(ev['prereg_deadline'])}" if ev.get("prereg_deadline") else ""
        c1.markdown(f":green[●] **Pre-registration is on** · {ev['preregs']} registered so far{closes}")
    else:
        c1.markdown(":gray[●] **Pre-registration is off.** Attendance-only: members simply check in on the day.")
    prereg_only = not ev.get("uses_checkin", 1)
    if prereg_only:
        c2.markdown(f"<div style='text-align:right;color:{MUTED};font-size:.85rem;padding-top:.35rem'>"
                    "Pre-registration only</div>", unsafe_allow_html=True)
    elif c2.button("Turn off pre-registration" if ev["prereg_open"] else "Turn on pre-registration",
                   key=f"toggle_prereg_{ev['id']}", width="stretch"):
        set_event_flag(ev["id"], "prereg_open", not ev["prereg_open"])
        notify(f"Pre-registration turned {'off' if ev['prereg_open'] else 'on'} for “{ev['event_name']}”.")
        st.rerun()
    if not ev["prereg_open"]:
        return
    with st.form(f"deadline_form_{ev['id']}", border=True):
        st.markdown("**Closing date and time**")
        deadline = deadline_inputs(f"prereg_{ev['id']}", ev.get("prereg_deadline"))
        st.caption("After this time the registration link shows a “registration has closed” message. "
                   + ("Nothing else happens: this event has no check-in at the venue." if prereg_only
                      else "Day-of check-in is not affected."))
        save_dl = st.form_submit_button("Save closing time", type="primary")
    if save_dl:
        problem = deadline_problem(deadline, ev.get("event_date"))
        if problem:
            st.error(problem)
        else:
            update_event(ev["id"], ev["event_name"], ev["event_type"], ev["event_date"], ev["venue"],
                         ev["is_open"], ev["prereg_open"], deadline)
            notify(f"Pre-registration for “{ev['event_name']}” now closes {fmt_deadline(deadline)}." if deadline
                   else f"Removed the closing time for “{ev['event_name']}”.")
            st.rerun()
    if not base:
        st.warning("Set the app address above to get the registration link.")
        return
    link = register_link(base, ev["id"])
    when = f" on {fmt_date(ev['event_date'])}" if ev["event_date"] else ""
    where = f" at {ev['venue']}" if ev["venue"] else ""
    closes = (f"Registration closes {fmt_deadline(ev['prereg_deadline'])}.\n\n"
              if ev.get("prereg_deadline") and not prereg_deadline_passed(ev) else "")
    on_the_day = ("" if prereg_only else
                  "On the day, simply scan the QR code at the entrance and enter your phone number. ")
    message = (f"{ev['event_name']}{when}{where}\n\nReserve your place here: {link}\n\n{closes}"
               f"{on_the_day}See you there.\n{APP_NAME}")
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
    if not ev.get("uses_checkin", 1):
        st.info("This event is set to pre-registration only, so it has no check-in QR code. If you need check-in "
                "at the venue after all, change “How members take part” under **Edit details**.",
                icon=":material/info:")
        return
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
        c5, c6 = st.columns([3, 2], vertical_alignment="bottom")
        with c5:
            mode = mode_input(f"edit_mode_{ev['id']}_{event_mode(ev)}", event_mode(ev))
        is_open = c6.toggle("Day-of check-in open", value=bool(ev["is_open"]),
                            help="Ignored for pre-registration-only events.")
        deadline = deadline_inputs(f"edit_{ev['id']}", ev.get("prereg_deadline"))
        save = st.form_submit_button("Save changes", type="primary")
    if save:
        problem = deadline_problem(deadline, ev_date.isoformat() if ev_date else None)
        if not name.strip():
            st.error("The event needs a name.")
        elif problem:
            st.error(problem)
        else:
            prereg_open, uses_checkin = mode_flags(mode)
            update_event(ev["id"], name, ev_type, ev_date.isoformat() if ev_date else None, venue, is_open, prereg_open,
                         deadline, uses_checkin)
            notify(f"Saved changes to “{name.strip()}”.")
            st.rerun()


def flyer_section(ev: dict):
    current = get_flyer(ev["id"])
    if current:
        st.image(current, caption="Current flyer, as members see it", width=280)
    else:
        st.caption("This event has no flyer yet.")
    look_now = get_app_setting("flyer_look", "frosted")
    look = st.radio("How the flyer shows on the member pages", list(FLYER_LOOKS), format_func=FLYER_LOOKS.get,
                    index=list(FLYER_LOOKS).index(look_now) if look_now in FLYER_LOOKS else 0,
                    key=f"flyer_look_{ev['id']}",
                    help="Applies to every programme's check-in and pre-registration pages.")
    if look != look_now:
        set_app_setting("flyer_look", look)
        notify(f"Flyer background set to: {FLYER_LOOKS[look].split(':')[0]}.")
        st.rerun()
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
    if not ev.get("uses_checkin", 1):
        c1, c2 = st.columns([1, 2])
        c1.metric("Pre-registered", len(df))
        if ev.get("prereg_deadline"):
            c2.markdown(f"**Registration {'closed' if prereg_deadline_passed(ev) else 'closes'}**  \n"
                        f"{fmt_deadline(ev['prereg_deadline'])}")
        view = filter_people(df.drop(columns=["Arrived At"]), st.text_input("Search by name or phone", key="prereg_search"))
        st.dataframe(view, width="stretch", hide_index=True)
        download_pair(view, f"{safe_filename(ev['event_name'])}_preregistrations", "Pre-registrations", f"prereg_{ev['id']}")
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
    if not ev.get("uses_checkin", 1):
        st.info("This event is pre-registration only, so it has no check-ins. See the **Pre-registrations** tab.",
                icon=":material/info:")
        return
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
    bday_text = fmt_birthday(m.get("birth_day"), m.get("birth_month"))
    st.caption(f"{fmt_phone(m['phone_number'])} · {m.get('country') or 'Country not recorded'} · "
               + (f"birthday {bday_text} · " if bday_text else "") + "joined "
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
            st.markdown("**Birthday**")
            b_day, b_month = birthday_inputs(f"edit_member_{member_id}", False, m.get("birth_day"), m.get("birth_month"))
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
                "birth_day": b_day if (b_day and b_month) else None,
                "birth_month": b_month if (b_day and b_month) else None,
            }
            problem = phone_problem(data["phone_number"]) or birthday_problem(b_day, b_month, False)
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
        st.markdown("Removes this person, their check-ins and pre-registrations. "
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
    df = with_birthday_column(search_members(term, country), after="Country")
    if df.empty:
        st.caption("No members match." if (term or country != "All countries")
                   else "No members yet. Share the membership link or import your existing list.")
        return
    st.caption(f"{len(df)} member(s)")
    st.dataframe(df.drop(columns=["id"]), hide_index=True, width="stretch", height=320)
    everyone = with_birthday_column(all_members_export(), after="Country")
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
               f"Already registered? Open the link anyway and enter your number so you can add your birthday.\n\n"
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
        st.caption("Your file needs at least a name column and a phone number column. Email and birthday "
                   "columns are optional. Birthdays can be written like 12/03, 12 March or 1990-03-12.")
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
    c4, c5 = st.columns(2)
    has_bday = any(w in str(c).lower() for c in cols for w in ("birth", "dob"))
    bday_col = c4.selectbox("Column with birthdays", [IMPORT_NONE] + cols,
                            index=(_guess(cols, ["birth", "dob"]) + 1) if has_bday else 0)
    default_country = c5.selectbox("Numbers without a + are from", COUNTRY_NAMES[:-1], format_func=country_label)

    rows, seen = [], set()
    for _, r in raw.iterrows():
        name = str(r.get(name_col) or "").strip()
        phone = to_intl(str(r.get(phone_col) or ""), default_country)
        email = str(r.get(email_col) or "").strip() if email_col != IMPORT_NONE else ""
        bday = parse_birthday(r.get(bday_col)) if bday_col != IMPORT_NONE else None
        if not name or name.lower() == "nan":
            status = "Skipped: no name"
        elif phone_problem(phone):
            status = "Skipped: phone number looks wrong"
        elif phone in seen or find_member_id(phone):
            status = "Already on the list"
        else:
            status = "Will be added"
        seen.add(phone)
        rows.append({"Name": name, "Phone": phone, "Email": email if email.lower() != "nan" else "",
                     "Birthday": fmt_birthday(*bday) if bday else "", "_bday": bday, "Result": status})
    preview = pd.DataFrame(rows)
    new = preview[preview["Result"] == "Will be added"]
    st.dataframe(preview.drop(columns=["_bday"]), hide_index=True, width="stretch", height=280)
    st.caption(f"{len(new)} new · {int((preview['Result'] == 'Already on the list').sum())} already on the list · "
               f"{int(preview['Result'].str.startswith('Skipped').sum())} skipped")
    if st.button(f"Import {len(new)} member(s)", type="primary", disabled=new.empty, key="do_import"):
        # Work from the plain list, not the table: pandas turns an empty birthday
        # into NaN, which used to crash the import halfway through.
        to_add = [r for r in rows if r["Result"] == "Will be added"]
        added, failed = 0, []
        progress = st.progress(0.0, text="Importing members…")
        with closing(get_conn()) as conn, conn:          # one transaction: fast, even for big lists
            for i, r in enumerate(to_add, 1):
                bday = r["_bday"] if isinstance(r["_bday"], tuple) and len(r["_bday"]) == 2 else (None, None)
                try:
                    conn.execute(
                        f"""INSERT INTO members ({', '.join(MEMBER_FIELDS)}, consent_at)
                            VALUES ({', '.join('?' for _ in MEMBER_FIELDS)}, CURRENT_TIMESTAMP)""",
                        [{"full_name": r["Name"], "phone_number": r["Phone"], "whatsapp_number": r["Phone"],
                          "email": r["Email"] or None, "country": country_from_number(r["Phone"]),
                          "emergency_contact_name": None, "emergency_contact_phone": None,
                          "parent_guardian_phone": None, "is_minor": 0, "source": "Imported",
                          "birth_day": bday[0], "birth_month": bday[1]}[f] for f in MEMBER_FIELDS])
                    added += 1
                except sqlite3.IntegrityError:
                    pass                                   # already on the list
                except Exception as err:                   # one odd row never stops the rest
                    failed.append(f"{r['Name']} ({err})")
                if i % 25 == 0 or i == len(to_add):
                    progress.progress(i / len(to_add), text=f"Importing members… {i} of {len(to_add)}")
        st.session_state.pop("import_file", None)
        msg = f"Imported {added} member(s) from {upload.name}."
        if failed:
            msg += f" {len(failed)} row(s) couldn't be added: " + "; ".join(failed[:5])
        notify(msg)
        st.rerun()


def members_tab():
    t_dir, t_bday, t_link, t_import = st.tabs([":material/groups: Directory", ":material/cake: Birthdays",
                                               ":material/share: Membership link",
                                               ":material/upload_file: Import existing members"])
    with t_dir:
        directory_section()
    with t_bday:
        birthdays_section()
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


# ---------------------------------------------------------------------------
# Admin: Prophet Masterclass
# ---------------------------------------------------------------------------
def member_whatsapp_url(phone: str, text: str) -> str:
    return f"https://wa.me/{re.sub(r'[^0-9]', '', phone or '')}?text={quote(text)}"


def pa_grid_section():
    st.markdown("#### PA Verification Grid")
    st.caption("Everyone who registered for a paid session and is waiting for their payment to be checked. "
               "Confirm the money has arrived, then tap **Approve Payment**. Their token opens the live room straight away.")
    notice = st.session_state.pop("pa_notice", None)
    if notice:
        st.success(notice[0])
        if notice[1]:
            with st.container(key="wa_pa_tell"):
                st.link_button("Tell them on WhatsApp", notice[1], icon=":material/chat:")
    sessions = mc_sessions()
    paid = [x for x in sessions if x["session_type"] == "Paid"]
    if not paid:
        st.info("No paid sessions yet. Create one in **Sessions** and set the type to Paid.")
        return
    opts = [0] + [x["id"] for x in paid]
    names = {0: "All paid sessions", **{x["id"]: x["session_name"] for x in paid}}
    c1, c2 = st.columns([1, 1])
    pick = c1.selectbox("Session", opts, format_func=lambda i: names[i], key="pa_session")
    search = c2.text_input("Find a token, name or number", key="pa_search", placeholder="IGNITE-8921")
    df = mc_registrations(pick or None, MC_PENDING)
    df = df[df["Type"] == "Paid"]
    if search.strip():
        t = search.strip()
        digits = re.sub(r"\D", "", t).lstrip("0")
        mask = (df["Token"].str.contains(t.upper(), regex=False) |
                df["Full Name"].str.contains(t, case=False, regex=False))
        if digits:
            mask |= df["Phone"].str.contains(digits, regex=False)
        df = df[mask]
    m = st.columns(3)
    m[0].metric("Waiting for approval", len(df))
    m[1].metric("Approved (paid)", int(sum(x["approved"] for x in paid if not pick or x["id"] == pick)))
    m[2].metric("Joined the room", int(sum(x["joined"] for x in paid if not pick or x["id"] == pick)))
    if df.empty:
        st.markdown('<div class="ig-note">Nothing waiting. Every paid registration has been checked.</div>',
                    unsafe_allow_html=True)
    for _, r in df.head(100).iterrows():
        with st.container(key=f"glass_pa_{r['Token']}"):
            a, b = st.columns([3, 1.2], vertical_alignment="center")
            a.markdown(f"<b>{esc(r['Full Name'])}</b> &nbsp;<code>{esc(r['Token'])}</code>"
                       f"<div class='ig-rowmeta'>{esc(fmt_phone(r['Phone']))} · {esc(r['Session'])} · "
                       f"registered {esc(r['Registered (GMT)'])} GMT</div>", unsafe_allow_html=True)
            if b.button("Approve Payment", key=f"approve_{r['Token']}", type="primary", width="stretch"):
                mc_set_status([r["Token"]], MC_PAID, current_admin().get("name", "PA"))
                log_action(f"Approved masterclass payment for {r['Full Name']} ({r['Token']})")
                base = current_base_url()
                room = f"\n\nJoin here when it starts: {live_link(base)}" if base_url_ok(base) else ""
                text = (f"Hello {first_name(r['Full Name'])}, your payment for the Prophet Masterclass "
                        f"({r['Session']}) is confirmed. Your token {r['Token']} is now active.{room}")
                st.session_state.pa_notice = (f"Approved {r['Full Name']} ({r['Token']}).",
                                              member_whatsapp_url(r["Phone"], text))
                st.rerun()
    if len(df) > 100:
        st.caption(f"Showing the first 100 of {len(df)}. Use the search box to find someone.")
    with st.expander("Recently approved (undo a mistake)"):
        done = mc_registrations(pick or None, MC_PAID).head(30)
        if done.empty:
            st.caption("No approved payments yet.")
        for _, r in done.iterrows():
            a, b = st.columns([3, 1], vertical_alignment="center")
            a.markdown(f"**{r['Full Name']}** `{r['Token']}` · {r['Session']} · by {r['Approved By'] or '?'}")
            if b.button("Move back to pending", key=f"undo_{r['Token']}", width="stretch"):
                mc_set_status([r["Token"]], MC_PENDING, current_admin().get("name", "PA"))
                log_action(f"Moved {r['Full Name']} ({r['Token']}) back to Pending Verification")
                st.rerun()


def mc_session_form(key: str, sess: dict | None = None):
    sess = sess or {}
    with st.form(key):
        name = st.text_input("Session name *", value=sess.get("session_name", ""),
                             placeholder="e.g. Prophet Masterclass: The Prophetic Voice")
        c1, c2 = st.columns(2)
        stype = c1.selectbox("Session type", MC_TYPES, index=MC_TYPES.index(sess.get("session_type", "Free")),
                             help="Free: everyone is approved straight away. Paid: the PA approves each payment.")
        try:
            dval = date.fromisoformat(sess["session_date"]) if sess.get("session_date") else None
        except ValueError:
            dval = None
        sdate = c2.date_input("Date (optional)", value=dval, format="DD/MM/YYYY")
        url = st.text_input("Secret streaming link", value=sess.get("streaming_url") or "",
                            placeholder="https://youtube.com/live/… or https://zoom.us/j/…",
                            help="Only approved members are sent here. It never appears on a public page.")
        price = st.text_input("Payment instructions (paid sessions)", value=sess.get("price_note") or "",
                              placeholder="e.g. GHS 150 by MoMo to 024 000 0000 (Ignite Prayer Network)")
        active = st.toggle("Active (open for enrolment and the live room)", value=bool(sess.get("active_status", 1)))
        saved = st.form_submit_button("Save session" if sess else "Create session", type="primary")
    if not saved:
        return None
    errors = []
    if len(name.strip()) < 3:
        errors.append("Give the session a name.")
    if url.strip() and not url.strip().startswith("https://"):
        errors.append("The streaming link should start with https://")
    if errors:
        show_errors(errors)
        return None
    return {"name": name, "session_type": stype, "streaming_url": url, "session_date": sdate.isoformat() if sdate else None,
            "price_note": price, "active": active}


def mc_sessions_section():
    with st.expander("Create a new masterclass session", expanded=not mc_sessions()):
        data = mc_session_form("mc_new")
        if data:
            mc_create_session(data["name"], data["session_type"], data["streaming_url"], data["session_date"],
                              data["price_note"], data["active"])
            notify(f"Created the masterclass session “{data['name'].strip()}” ({data['session_type']}).")
            st.rerun()
    sessions = mc_sessions()
    if not sessions:
        return
    rows = pd.DataFrame([{
        "Session": x["session_name"], "Type": x["session_type"], "Status": "Active" if x["active_status"] else "Closed",
        "Date": fmt_date(x["session_date"]), "Registered": x["registered"], "Approved": x["approved"],
        "Pending": x["pending"], "Joined Room": x["joined"],
        "Stream Link": "Set" if (x["streaming_url"] or "").startswith("http") else "Missing"} for x in sessions])
    st.dataframe(rows, hide_index=True, width="stretch")
    by_id = {x["id"]: x for x in sessions}
    pick = st.selectbox("Manage a session", list(by_id), key="mc_manage",
                        format_func=lambda i: f"{by_id[i]['session_name']} · {by_id[i]['session_type']}")
    sess = by_id[pick]
    c1, c2 = st.columns([3, 2])
    with c1:
        data = mc_session_form(f"mc_edit_{pick}", sess)
        if data:
            if data["session_type"] != sess["session_type"] and sess["registered"]:
                st.warning("People have already registered under the old type. Their status stays as it is; "
                           "use the PA grid if anyone needs changing.")
            mc_update_session(pick, data["name"], data["session_type"], data["streaming_url"], data["session_date"],
                              data["price_note"], data["active"])
            notify(f"Updated the masterclass session “{data['name'].strip()}”.")
            st.rerun()
    with c2:
        base = current_base_url()
        if base_url_ok(base):
            st.markdown("**Enrolment link for this session**")
            st.code(masterclass_link(base, pick), language=None)
        with st.container(key="glass_mc_delete"):
            st.markdown("**Delete this session**")
            st.caption(f"Removes the session and its {sess['registered']} registration(s). This can't be undone.")
            sure = st.checkbox("Yes, delete it", key=f"mc_del_sure_{pick}")
            if st.button("Delete session", key=f"mc_del_{pick}", disabled=not sure):
                n = mc_delete_session(pick)
                notify(f"Deleted the masterclass session “{sess['session_name']}” and {n} registration(s).")
                st.rerun()


def mc_registrations_section():
    sessions = mc_sessions()
    if not sessions:
        st.info("No sessions yet.")
        return
    opts = [0] + [x["id"] for x in sessions]
    names = {0: "All sessions", **{x["id"]: x["session_name"] for x in sessions}}
    c1, c2, c3 = st.columns([1.3, 1, 1.2])
    pick = c1.selectbox("Session", opts, format_func=lambda i: names[i], key="mcr_session")
    status = c2.selectbox("Status", ["Any", MC_FREE, MC_PENDING, MC_PAID], key="mcr_status")
    search = c3.text_input("Search", key="mcr_search", placeholder="Name, token or number")
    df = mc_registrations(pick or None, None if status == "Any" else status)
    if search.strip():
        t = search.strip()
        digits = re.sub(r"\D", "", t).lstrip("0")
        mask = (df["Token"].str.contains(t.upper(), regex=False) |
                df["Full Name"].str.contains(t, case=False, regex=False) |
                df["Email"].str.contains(t, case=False, regex=False))
        if digits:
            mask |= df["Phone"].str.contains(digits, regex=False)
        df = df[mask]
    st.caption(f"{len(df)} registration(s). “Room Entries” counts every time a token was used to enter; "
               "a high number can mean a token is being shared.")
    view = df.copy()
    view["Phone"] = view["Phone"].map(fmt_phone)
    st.dataframe(view, hide_index=True, width="stretch", height=380)
    download_pair(df, "Ignite_masterclass_registrations", "Masterclass", "mcr")
    if current_admin().get("role") in ("Owner", "Admin") and not df.empty:
        with st.expander("Remove a registration"):
            tok = st.selectbox("Token", df["Token"].tolist(), key="mcr_del_tok",
                               format_func=lambda t: f"{t} · {df.loc[df['Token'] == t, 'Full Name'].iloc[0]}")
            if st.button("Remove", key="mcr_del_go"):
                mc_delete_registration(tok)
                notify(f"Removed masterclass registration {tok}.")
                st.rerun()


def mc_links_section():
    base = current_base_url()
    st.markdown("#### PA WhatsApp number")
    st.caption("People on paid sessions get a button that opens a WhatsApp chat with this number, "
               "already filled in with their name and token.")
    current = get_app_setting("pa_whatsapp") or get_setting("PA_WHATSAPP", DEFAULT_PA_WHATSAPP)
    st.markdown(f"Currently: **{fmt_phone(current)}**")
    with st.form("pa_number_form"):
        raw, country = admin_phone_input("pa_wa", "PA's WhatsApp number")
        if st.form_submit_button("Save PA number", type="primary"):
            phone = to_intl(raw, country)
            if phone_problem(phone):
                st.error(phone_problem(phone))
            else:
                set_app_setting("pa_whatsapp", phone)
                notify(f"Set the PA WhatsApp number to {fmt_phone(phone)}.")
                st.rerun()
    if not base_url_ok(base):
        st.warning("Set the app's web address (Members → Membership link) to get the share links and QR codes.")
        return
    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Masterclass enrolment link**")
        st.code(masterclass_link(base), language=None)
        poster_block(masterclass_link(base), "Prophet Masterclass", "", "", "", "Enrolment",
                     "Scan to reserve your seat", "Ignite_masterclass_enrol", "mc_enrol")
    with c2:
        st.markdown("**Join Live Masterclass Room**")
        st.code(live_link(base), language=None)
        poster_block(live_link(base), "Join the Live Masterclass Room", "", "", "", "Live room",
                     "Scan and enter your token", "Ignite_masterclass_live", "mc_live")


def masterclass_tab(pa_only: bool = False):
    if pa_only:
        tabs = st.tabs(["PA Verification Grid", "Registrations"])
        with tabs[0]:
            pa_grid_section()
        with tabs[1]:
            mc_registrations_section()
        return
    tabs = st.tabs(["PA Verification Grid", "Sessions", "Registrations", "Links & PA number"])
    with tabs[0]:
        pa_grid_section()
    with tabs[1]:
        mc_sessions_section()
    with tabs[2]:
        mc_registrations_section()
    with tabs[3]:
        mc_links_section()


# ---------------------------------------------------------------------------
# Admin: daily Google Meet
# ---------------------------------------------------------------------------
def meet_settings_section():
    st.markdown("#### Your Google Meet room")
    st.caption("Paste the link to your own Google Meet room (your paid Workspace account works). Members never see "
               "it on a page: they go through the Ignite join link, which records them and then opens your room.")
    with st.form("meet_settings"):
        url = st.text_input("Google Meet link", value=get_app_setting("meet_url"),
                            placeholder="Paste your link, e.g. https://meet.google.com/xxx-xxxx-xxx")
        c1, c2, c3 = st.columns(3)
        is_open = c1.toggle("Room open", value=get_app_setting("meet_open", "1") == "1",
                            help="When off, the join page says the room isn't open.")
        times = [""] + [f"{h:02d}:{m:02d}" for h in range(24) for m in (0, 15, 30, 45)]
        start = c2.selectbox("Prayer starts (GMT)", times, index=times.index(meet_setting("meet_start"))
                             if meet_setting("meet_start") in times else 0, format_func=lambda t: t or "Not set")
        end = c3.selectbox("Prayer ends (GMT)", times, index=times.index(meet_setting("meet_end"))
                           if meet_setting("meet_end") in times else 0, format_func=lambda t: t or "Not set")
        st.markdown("**When does someone count as having stayed?**")
        c4, c5, c6 = st.columns(3)
        min_pct = c4.number_input("In the call for at least (% of the prayer)", 10, 100,
                                  int(meet_setting("meet_min_pct")), step=5)
        grace = c5.number_input("And still there within (minutes of the end)", 0, 60, int(meet_setting("meet_grace")))
        code_min = c6.number_input("Closing code stays open for (minutes)", 2, 60, int(meet_setting("meet_code_minutes")))
        if st.form_submit_button("Save", type="primary"):
            if url.strip() and not url.strip().startswith("https://"):
                st.error("Paste the full Meet link, starting with https://")
            else:
                for k, v in {"meet_url": url.strip(), "meet_open": "1" if is_open else "0", "meet_start": start,
                             "meet_end": end, "meet_min_pct": str(int(min_pct)), "meet_grace": str(int(grace)),
                             "meet_code_minutes": str(int(code_min))}.items():
                    set_app_setting(k, v)
                notify("Updated the Google Meet room settings.")
                st.rerun()


def meet_code_section():
    state = meet_code_state()
    minutes = int(meet_setting("meet_code_minutes"))
    st.markdown("#### Closing code")
    if state["active"]:
        until = datetime.fromtimestamp(state["until"], timezone.utc).strftime("%H:%M")
        st.markdown(f"""<div class="ig-code"><div class="k">Today's closing code</div>
                          <div class="n">{esc(state['code'])}</div>
                          <div class="t">Open until {until} GMT. Say it in the room or post it in the Meet chat.</div></div>""",
                    unsafe_allow_html=True)
        if st.button("Close the code now", key="meet_code_close", width="stretch"):
            meet_close_code()
            notify("Closed today's closing code.")
            st.rerun()
    else:
        st.caption(f"In the last minutes of prayer, reveal a code. Members open the Ignite prayer link again and type it "
                   f"to confirm they stayed. It works on any Google plan and closes itself after {minutes} minutes, "
                   "so people who left early can't get it later.")
        if st.button("Reveal closing code", key="meet_code_open", type="primary", width="stretch",
                     icon=":material/key:"):
            code = meet_open_code(minutes)
            notify(f"Revealed today's closing code ({code}).")
            st.rerun()


def meet_report_section(day: date):
    with st.expander("Upload Google Meet's attendance report (most accurate)"):
        st.caption("Google sends this report to the meeting organiser on Workspace Business Plus, Enterprise and "
                   "Education plans, when attendance tracking is on. It shows how long each person was in the call. "
                   "Download it as Excel or CSV and upload it here. Each person's name or email is matched to the "
                   "people who joined through the Ignite link.")
        upload = st.file_uploader("Attendance report (.xlsx or .csv)", type=["xlsx", "csv"], key="meet_report_file")
        if not upload:
            return
        try:
            table = _read_report_table(upload)
            meeting_minutes = None
            ms, me = meet_setting("meet_start"), meet_setting("meet_end")
            if ms and me:
                h1, m1 = map(int, ms.split(":"))
                h2, m2 = map(int, me.split(":"))
                meeting_minutes = ((h2 * 60 + m2) - (h1 * 60 + m1)) % (24 * 60) or None
            df = parse_attendance_report(table, int(meet_setting("meet_grace")), int(meet_setting("meet_min_pct")),
                                         meeting_minutes)
        except ValueError as err:
            st.error(str(err))
            return
        except Exception:
            st.error("That file couldn't be read. Download the report from Google Sheets as .xlsx or .csv and try again.")
            return
        view = df[["Name", "Email", "Minutes", "Result"]].copy()
        view["Minutes"] = view["Minutes"].map(lambda v: "" if v is None or pd.isna(v) else f"{v:.0f}")
        st.caption(f"Prayer length used: {df.attrs['span']:.0f} min · stayed = in the call at least "
                   f"{meet_setting('meet_min_pct')}% of it and still there within {meet_setting('meet_grace')} min of the end.")
        st.dataframe(view, hide_index=True, width="stretch", height=260)
        st.caption(f"{int(df['Stayed'].sum())} stayed to the end · {int((~df['Stayed']).sum())} left early")
        if st.button(f"Save this report for {day.strftime('%a %d %b %Y')}", type="primary", key="meet_report_save"):
            matched, total = meet_save_report(day.isoformat(), df)
            st.session_state.pop("meet_report_file", None)
            notify(f"Saved the Google Meet report for {day.isoformat()}: {total} people, {matched} matched to the "
                   "Ignite join list.")
            st.rerun()


def meet_tab():
    left, right = st.columns([3, 2])
    with left:
        meet_settings_section()
    with right:
        meet_code_section()
    st.divider()
    st.markdown("#### Who prayed with us")
    c1, c2 = st.columns([1, 2])
    day = c1.date_input("Day", value=today_local(), format="DD/MM/YYYY", key="meet_day")
    summary = meet_day_summary(day.isoformat())
    stayed = int((summary["Status"] == STATUS_STAYED).sum())
    early = int((summary["Status"] == STATUS_EARLY).sum())
    unconf = int((summary["Status"] == STATUS_UNCONFIRMED).sum())
    m = st.columns(4)
    m[0].metric("Joined", len(summary))
    m[1].metric("Stayed to the end", stayed)
    m[2].metric("Left early", early)
    m[3].metric("Joined, not confirmed", unconf,
                help="Joined through the Ignite link but never typed the closing code, and no report covers them yet.")
    meet_report_section(day)
    if summary.empty:
        st.caption("Nobody joined through the app on this day.")
    else:
        filt = c2.radio("Show", ["Everyone", STATUS_STAYED, STATUS_EARLY, STATUS_UNCONFIRMED], horizontal=True,
                        key="meet_filter")
        view = summary if filt == "Everyone" else summary[summary["Status"] == filt]
        shown = view.copy()
        shown["Phone"] = shown["Phone"].map(lambda p: fmt_phone(p) if p else "")
        st.dataframe(shown, hide_index=True, width="stretch", height=340)
        download_pair(summary, f"Ignite_prayer_{day.isoformat()}", "Prayer attendance", "meet_day_dl")
        people = summary[summary["Phone"] != ""]
        if not people.empty:
            with st.expander("Correct someone's status"):
                labels = {r["Phone"]: f"{r['Meet Display Name']} · {fmt_phone(r['Phone'])} · {r['Status']}"
                          for _, r in people.iterrows()}
                who = st.selectbox("Person", list(labels), format_func=labels.get, key="meet_fix_who")
                b1, b2, b3 = st.columns(3)
                name = people.loc[people["Phone"] == who, "Meet Display Name"].iloc[0]
                if b1.button("Mark as stayed", key="meet_fix_yes", width="stretch"):
                    meet_record_stay(day.isoformat(), who, name, "manual", True)
                    notify(f"Marked {name} as stayed to the end on {day.isoformat()}.")
                    st.rerun()
                if b2.button("Mark as left early", key="meet_fix_no", width="stretch"):
                    meet_record_stay(day.isoformat(), who, name, "manual", False)
                    notify(f"Marked {name} as left early on {day.isoformat()}.")
                    st.rerun()
                if b3.button("Clear my correction", key="meet_fix_clear", width="stretch"):
                    meet_remove_stay(day.isoformat(), who, "manual")
                    st.rerun()
    totals = meet_daily_totals(30)
    if not totals.empty:
        st.markdown("#### Last 30 days")
        st.bar_chart(totals.set_index("Date"), color=[GOLD, ROYAL], stack=False, y_label="People", x_label="")
    with st.expander("Share the join link"):
        base = current_base_url()
        if base_url_ok(base):
            st.markdown("Share this instead of the Meet link, so attendance is recorded.")
            st.code(meet_link(base), language=None)
            msg = (f"Our prayer room is open. Join here so we can see you came:\n{meet_link(base)}\n\n"
                   f"Please stay to the end and enter the closing code on the same page.\n{APP_NAME}")
            st.link_button("Share on WhatsApp", f"https://wa.me/?text={quote(msg)}", type="primary")
            poster_block(meet_link(base), "Daily Prayer Room", "", "", "", "Google Meet",
                         "Scan to join the prayer room", "Ignite_google_meet", "meet")
        else:
            st.warning("Set the app's web address (Members → Membership link) to get the share link.")
    with st.expander("Export every join"):
        if True:
            full = meet_full_export()
            st.caption(f"{len(full)} join(s) recorded in total.")
            download_pair(full, "Ignite_meet_all", "Google Meet", "meet_all_dl")

def admin_page():
    if not admin_is_authenticated():
        admin_login()
        return
    user = current_admin()
    bar = st.container(key="glass_topbar")
    c1, c2 = bar.columns([4, 1], vertical_alignment="center")
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
        with tabs[0], st.container(key="glass_panel_usher"):
            attendance_tab(limited=True)
        with tabs[1], st.container(key="glass_panel_account"):
            account_section()
        return
    if user["role"] == "PA":
        tabs = st.tabs([":material/verified: Masterclass", ":material/person: My account"])
        with tabs[0], st.container(key="glass_panel_pa"):
            masterclass_tab(pa_only=True)
        with tabs[1], st.container(key="glass_panel_pa_account"):
            account_section()
        return

    tabs = st.tabs([":material/dashboard: Overview", ":material/event: Events", ":material/how_to_reg: Pre-registrations",
                    ":material/fact_check: Check-ins", ":material/group: Members",
                    ":material/school: Masterclass", ":material/videocam: Google Meet",
                    ":material/admin_panel_settings: Team", ":material/backup: Backup"])
    with tabs[0], st.container(key="glass_panel_overview"):
        overview_tab()
    with tabs[1], st.container(key="glass_panel_events"):
        events_tab()
    with tabs[2], st.container(key="glass_panel_prereg"):
        prereg_tab()
    with tabs[3], st.container(key="glass_panel_checkins"):
        attendance_tab()
    with tabs[4], st.container(key="glass_panel_members"):
        members_tab()
    with tabs[5], st.container(key="glass_panel_masterclass"):
        masterclass_tab()
    with tabs[6], st.container(key="glass_panel_meet"):
        meet_tab()
    with tabs[7], st.container(key="glass_panel_team"):
        team_tab()
    with tabs[8], st.container(key="glass_panel_backup"):
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
    elif st.query_params.get("mode") in ("join", "masterclass", "live", "meet"):
        mode = st.query_params.get("mode")
    else:
        mode = "checkin"
    titles = {"admin": "Admin", "register": "Reserve your place", "join": "Join", "checkin": "Check-in",
              "masterclass": "Prophet Masterclass", "live": "Live Masterclass Room", "meet": "Prayer Room"}
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
    elif mode == "masterclass":
        masterclass_page()
        public_footer()
    elif mode == "live":
        live_gate_page()
        public_footer()
    elif mode == "meet":
        meet_page()
        public_footer()
    else:
        checkin_page()
        public_footer()


if __name__ == "__main__":
    main()
