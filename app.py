"""
Ignite Prayer Network: Event Registration & Check-In System
===========================================================

One Streamlit app, three pages:

  * Day-of check-in (default). Each event's check-in QR code opens
        https://<your-app>.streamlit.app/?event=<id>
  * Pre-registration, shared before the programme (e.g. on WhatsApp):
        https://<your-app>.streamlit.app/?event=<id>&mode=register
  * Admin portal. Tap "Admin login" at the bottom of any public page
    (or open .../?view=admin) and enter the admin password.

How records are kept apart
  * Every event (Asteri, Shekinah Glory, Impromptu) has its own event_id, and
    every record is written and read with it, so events never mix.
  * Pre-registrations live in their own table (pre_registrations), separate
    from the day-of check-ins (attendance). Each has its own ledger and export.
  * Both share one members list, so someone who pre-registers only needs their
    phone number to check in on the day.

Optional settings (Streamlit Cloud: Settings -> Secrets):
    ADMIN_PASSWORD = "your-strong-password"
    APP_BASE_URL   = "https://your-app-name.streamlit.app"
"""

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
DEFAULT_ADMIN_PASSWORD = "IgniteAdmin2026"     # change before going live
EVENT_TYPES = ["Asteri", "Shekinah Glory", "Impromptu"]

SUCCESS_SECONDS = 2          # how long the welcome screen stays up
ADMIN_SESSION_MINUTES = 30   # admin is logged out after this much idle time
MAX_LOGIN_ATTEMPTS = 5       # wrong passwords allowed before a cool-down
LOCKOUT_SECONDS = 60
FLYER_MAX_UPLOAD_MB = 10
FLYER_MAX_WIDTH = 1200
LEDGER_REFRESH_SECONDS = 15

BRAND = "#E4572E"            # Ignite flame orange
BRAND_DARK = "#B83A17"
ACCENT = "#F3A712"           # gold
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
                  CASE WHEN e.prereg_open THEN 'Open' ELSE 'Closed' END AS "Registration",
                  CASE WHEN e.is_open THEN 'Open' ELSE 'Closed' END AS "Check-In"
           FROM events e
           ORDER BY COALESCE(e.event_date, date(e.created_at)) DESC, e.id DESC"""
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
          .ig-footer {{text-align:center; opacity:.55; font-size:.8rem; margin-top:.4rem;}}
        </style>
        """,
        unsafe_allow_html=True,
    )


def brand_header(subtitle: str = "", compact: bool = False):
    cls = "ig-hero ig-compact" if compact else "ig-hero"
    inner = (f'<div><div class="ig-title">{APP_NAME}</div><div class="ig-sub">{esc(subtitle)}</div></div>'
             if compact else
             f'<div class="ig-title">{APP_NAME}</div><div class="ig-sub">{esc(subtitle)}</div>')
    st.markdown(f'<div class="{cls}"><div class="ig-flame">🔥</div>{inner}</div>', unsafe_allow_html=True)


def event_card(ev: dict):
    colour = TYPE_COLOURS.get(ev["event_type"], BRAND)
    meta = " &nbsp;·&nbsp; ".join(
        p for p in (
            f"📅 {esc(fmt_date(ev.get('event_date')))}" if ev.get("event_date") else "",
            f"📍 {esc(ev.get('venue'))}" if ev.get("venue") else "",
        ) if p
    )
    st.markdown(
        f"""<div class="ig-event">
              <span class="ig-badge" style="background:{colour}">{esc(ev['event_type'])}</span>
              <div class="ig-event-name">{esc(ev['event_name'])}</div>
              {f'<div class="ig-event-meta">{meta}</div>' if meta else ''}
            </div>""",
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Public pages: shared pieces
# ---------------------------------------------------------------------------
def flash_and_reset(kind: str, title: str, message: str):
    """Store the message, bump the form counter (new widget keys mean blank
    forms) and rerun. The next run shows ONLY the message, then the form."""
    st.session_state.flash = (kind, title, message)
    st.session_state.form_nonce += 1
    st.rerun()


def show_flash_if_any():
    """Privacy screen: while the message is up, nothing else is drawn."""
    if "flash" not in st.session_state:
        return
    kind, title, message = st.session_state.pop("flash")
    css, icon = ("ig-success", "✅") if kind == "success" else ("ig-notice", "👋")
    placeholder = st.empty()
    placeholder.markdown(
        f'<div class="{css}"><div class="ig-check">{icon}</div><h2>{esc(title)}</h2><p>{esc(message)}</p></div>',
        unsafe_allow_html=True,
    )
    time.sleep(SUCCESS_SECONDS)
    placeholder.empty()
    st.rerun()


def phone_lookup_form(prefix: str, caption: str, not_found_hint: str):
    """Returns a member id after a valid submit, otherwise None."""
    nonce = st.session_state.form_nonce
    with st.form(key=f"{prefix}_quick_{nonce}"):
        st.caption(caption)
        phone_raw = st.text_input("Phone number", placeholder="e.g. 024 123 4567",
                                  autocomplete="off", key=f"{prefix}_q_phone_{nonce}")
        submitted = st.form_submit_button("Continue" if prefix == "reg" else "Submit Check-In",
                                          type="primary", width="stretch")
    if not submitted:
        return None
    phone = normalize_phone(phone_raw)
    member_id = find_member_id(phone) if is_valid_phone(phone) else None
    if member_id is None:
        st.error(f"We couldn't find that number. {not_found_hint}")
    return member_id


def new_member_form(prefix: str, button_label: str):
    """Full registration form. Returns (member_id, is_new) after a valid
    submit, otherwise None. Existing numbers are reused, never overwritten."""
    nonce = st.session_state.form_nonce
    k = lambda name: f"{prefix}_{name}_{nonce}"
    with st.form(key=k("register")):
        st.caption("Fill this in once. After that you'll only ever need your phone number.")
        full_name = st.text_input("Full name *", autocomplete="off", key=k("name"))
        c1, c2 = st.columns(2)
        phone_raw = c1.text_input("Phone number *", autocomplete="off", key=k("phone"))
        whatsapp_raw = c2.text_input("WhatsApp (if different)", autocomplete="off", key=k("wa"))
        email = st.text_input("Email (optional)", autocomplete="off", key=k("email"))

        st.markdown("**Emergency contact**")
        c3, c4 = st.columns(2)
        ec_name = c3.text_input("Contact name *", autocomplete="off", key=k("ecn"))
        ec_phone_raw = c4.text_input("Contact phone *", autocomplete="off", key=k("ecp"))

        is_minor = st.checkbox("I am under 18", key=k("minor"))
        parent_raw = st.text_input("Parent/guardian phone (required if under 18)",
                                   autocomplete="off", key=k("par"))
        consent = st.checkbox(
            f"I agree that {APP_NAME} may keep these details for attendance records and to "
            "contact my emergency contact or parent/guardian if needed. *",
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
        errors.append("Please enter your full name.")
    if not is_valid_phone(phone):
        errors.append("Please enter a valid phone number.")
    if whatsapp_raw.strip() and not is_valid_phone(whatsapp):
        errors.append("The WhatsApp number doesn't look right.")
    if email.strip() and not is_valid_email(email):
        errors.append("The email address doesn't look right.")
    if not ec_name.strip():
        errors.append("Please enter an emergency contact name.")
    if not is_valid_phone(ec_phone):
        errors.append("Please enter a valid emergency contact phone.")
    if is_minor and not parent:
        errors.append("A parent/guardian phone number is required for anyone under 18.")
    if parent and not is_valid_phone(parent):
        errors.append("The parent/guardian phone doesn't look right.")
    if not consent:
        errors.append("Please tick the consent box so we can keep your details.")
    if errors:
        for e in errors:
            st.error(e)
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
    event_card(ev)


def public_footer():
    st.write("")
    st.divider()
    with st.container(key="admin_entry"):
        if st.button("🔒 Admin login", key="open_admin", type="tertiary"):
            st.session_state.show_admin = True
            st.rerun()
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
        st.warning("That check-in link isn't valid any more. Please pick your event below.")
    events = get_events(open_only=True)
    if not events:
        st.info("There are no events open for check-in right now. Please check back later.")
        return None
    if len(events) == 1:
        return events[0]
    return st.selectbox("Which event are you attending?", events, format_func=event_label)


def checkin_page():
    show_flash_if_any()
    banner = page_top("Event Check-In")
    ev = resolve_checkin_event()
    if ev is None:
        return
    show_event_header(banner, ev)

    if not ev["is_open"]:
        st.warning("Check-in for this event is closed. Please speak to an usher if you need help.")
        return

    tab_returning, tab_new = st.tabs(["✅  Returning / pre-registered", "📝  First time here"])
    with tab_returning:
        member_id = phone_lookup_form(
            "chk", "Registered before, or signed up online? Just enter your phone number.",
            "If it's your first time, use the “First time here” tab.")
        if member_id:
            if record_check_in(member_id, ev["id"]):
                flash_and_reset("success", "Success! Welcome to Ignite Network",
                                "Your attendance has been recorded. Enjoy the service!")
            else:
                flash_and_reset("info", "You're already checked in",
                                "Your earlier check-in for this event is still on record.")
    with tab_new:
        result = new_member_form("chk", "Submit Check-In")
        if result:
            member_id, is_new = result
            msg = ("You're registered and checked in. We're glad you're here!" if is_new
                   else "This number was already registered, so we've checked you in.")
            if record_check_in(member_id, ev["id"]):
                flash_and_reset("success", "Success! Welcome to Ignite Network", msg)
            else:
                flash_and_reset("info", "You're already checked in",
                                "Your earlier check-in for this event is still on record.")


# ---------------------------------------------------------------------------
# Public page: pre-registration (shared before the day)
# ---------------------------------------------------------------------------
def register_page():
    show_flash_if_any()
    banner = page_top("Pre-Registration")

    raw = st.query_params.get("event")
    ev = get_event(int(raw)) if raw and str(raw).isdigit() else None
    if ev is None:
        st.error("This registration link isn't valid. Please ask the organisers for the correct link.")
        return
    show_event_header(banner, ev)

    if not ev["prereg_open"]:
        st.warning("Registration for this event has closed. You can still check in at the venue on the day.")
        return

    st.markdown("**Reserve your place.** On the day, scan the check-in QR code at the entrance "
                "and just enter your phone number.")
    tab_known, tab_new = st.tabs(["✅  I've registered with Ignite before", "📝  I'm new"])
    done_msg = f"You're registered for {ev['event_name']}. See you there!"
    with tab_known:
        member_id = phone_lookup_form(
            "reg", "Enter the phone number you used before.",
            "If you haven't registered with us before, use the “I'm new” tab.")
        if member_id:
            if record_pre_registration(member_id, ev["id"]):
                flash_and_reset("success", "You're registered! 🎉", done_msg)
            else:
                flash_and_reset("info", "You're already registered", done_msg)
    with tab_new:
        result = new_member_form("reg", "Register")
        if result:
            member_id, _ = result
            if record_pre_registration(member_id, ev["id"]):
                flash_and_reset("success", "You're registered! 🎉", done_msg)
            else:
                flash_and_reset("info", "You're already registered", done_msg)


# ---------------------------------------------------------------------------
# Admin: access
# ---------------------------------------------------------------------------
def admin_is_authenticated() -> bool:
    if st.session_state.get("admin_expires", 0) > time.time():
        st.session_state.admin_expires = time.time() + ADMIN_SESSION_MINUTES * 60
        return True
    st.session_state.pop("admin_expires", None)
    return False


def leave_admin():
    st.session_state.pop("admin_expires", None)
    st.session_state.show_admin = False
    if "view" in st.query_params:
        del st.query_params["view"]
    st.rerun()


def admin_login():
    brand_header("Admin Portal")
    locked_until = st.session_state.get("locked_until", 0)
    if locked_until > time.time():
        st.error(f"Too many attempts. Try again in {int(locked_until - time.time())} seconds.")
    else:
        with st.form("admin_login", clear_on_submit=True):
            st.markdown("**Sign in to manage events and attendance**")
            pw = st.text_input("Admin password", type="password")
            go = st.form_submit_button("Unlock", type="primary", width="stretch")
        if go:
            expected = get_setting("ADMIN_PASSWORD", DEFAULT_ADMIN_PASSWORD)
            if hmac.compare_digest(pw.encode(), expected.encode()):
                st.session_state.admin_expires = time.time() + ADMIN_SESSION_MINUTES * 60
                st.session_state.failed_logins = 0
                st.rerun()
            else:
                time.sleep(1)
                st.session_state.failed_logins = st.session_state.get("failed_logins", 0) + 1
                if st.session_state.failed_logins >= MAX_LOGIN_ATTEMPTS:
                    st.session_state.locked_until = time.time() + LOCKOUT_SECONDS
                    st.session_state.failed_logins = 0
                st.error("Incorrect password.")
    if st.button("← Back to check-in", key="back_to_checkin", type="tertiary"):
        leave_admin()


def notify(message: str):
    """Show a confirmation at the top of the admin page after a rerun."""
    st.session_state.admin_notice = message


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
    c1.download_button("⬇️  Export CSV", df.to_csv(index=False).encode("utf-8"), file_name=f"{stem}.csv",
                       mime="text/csv", width="stretch", key=f"{key}_csv")
    c2.download_button("⬇️  Export Excel", to_excel_bytes(df, sheet), file_name=f"{stem}.xlsx",
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

    st.info("💾 Free hosting can wipe the database when the app restarts or updates. "
            "Download a backup from the **Backup** tab after every event.")

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
    with st.expander("➕  Create a new event", expanded=not has_events):
        with st.form("new_event", clear_on_submit=True):
            c1, c2 = st.columns(2)
            ev_type = c1.selectbox("Event type", EVENT_TYPES)
            ev_name = c2.text_input("Event name", placeholder=f"e.g. Asteri {datetime.now().year + 1}")
            c3, c4 = st.columns(2)
            ev_date = c3.date_input("Date (optional)", value=None, format="DD/MM/YYYY")
            venue = c4.text_input("Venue (optional)")
            flyer_file = st.file_uploader("Programme flyer (optional, JPEG or PNG)", type=["jpg", "jpeg", "png"])
            c5, c6 = st.columns(2)
            prereg_open = c5.toggle("Open pre-registration", value=True)
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
            notify(f"Created “{name}”. Its registration link and check-in QR code are ready below.")
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
    st.download_button("⬇️  Download poster (A4 PNG)", poster, file_name=f"{stem}.png",
                       mime="image/png", width="stretch", key=f"{key}_poster")
    st.download_button("⬇️  Download QR code only", to_png(make_qr_image(link)),
                       file_name=f"{stem}_QR.png", mime="image/png", width="stretch", key=f"{key}_qr")


def prereg_section(ev: dict, base):
    status = "🟢 Registration is open" if ev["prereg_open"] else "🔴 Registration is closed"
    c1, c2 = st.columns([3, 1])
    c1.markdown(f"{status} · **{ev['preregs']}** registered so far")
    if c2.button("Close registration" if ev["prereg_open"] else "Reopen registration",
                 key=f"toggle_prereg_{ev['id']}", width="stretch"):
        set_event_flag(ev["id"], "prereg_open", not ev["prereg_open"])
        notify("Pre-registration closed." if ev["prereg_open"] else "Pre-registration reopened.")
        st.rerun()
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
        st.link_button("💬  Share on WhatsApp", f"https://wa.me/?text={quote(message)}",
                       type="primary", width="stretch")
    with right:
        poster_block(ev, link, "Register Now", "Scan to register for this event",
                     f"Register_{safe_filename(ev['event_name'])}", f"reg_{ev['id']}")


def checkin_qr_section(ev: dict, base):
    status = "🟢 Check-in is open" if ev["is_open"] else "🔴 Check-in is closed"
    c1, c2 = st.columns([3, 1])
    c1.markdown(f"{status} · **{ev['attendees']}** checked in")
    if c2.button("Close check-in" if ev["is_open"] else "Reopen check-in",
                 key=f"toggle_open_{ev['id']}", width="stretch"):
        set_event_flag(ev["id"], "is_open", not ev["is_open"])
        notify("Check-in closed." if ev["is_open"] else "Check-in reopened.")
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
        prereg_open = c5.toggle("Pre-registration open", value=bool(ev["prereg_open"]))
        is_open = c6.toggle("Day-of check-in open", value=bool(ev["is_open"]),
                            help="Turn this off after the event so no one can check in late.")
        save = st.form_submit_button("Save changes", type="primary")
    if save:
        if not name.strip():
            st.error("The event needs a name.")
        else:
            update_event(ev["id"], name, ev_type, ev_date.isoformat() if ev_date else None,
                         venue, is_open, prereg_open)
            notify("Event details saved.")
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
                notify("Flyer saved.")
                st.rerun()
            except ValueError as err:
                st.error(str(err))
    if current and st.button("Remove flyer", key=f"remove_flyer_{ev['id']}"):
        set_event_flyer(ev["id"], None)
        notify("Flyer removed.")
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
        c1.download_button("⬇️  Export check-ins first (Excel)", to_excel_bytes(get_ledger(ev["id"]), "Check-ins"),
                           file_name=f"{stem}_checkins.xlsx", mime=XLSX_MIME,
                           key=f"pre_delete_chk_{ev['id']}", width="stretch")
    if ev["preregs"]:
        c2.download_button("⬇️  Export pre-registrations first (Excel)",
                           to_excel_bytes(get_prereg_ledger(ev["id"]), "Pre-registrations"),
                           file_name=f"{stem}_preregistrations.xlsx", mime=XLSX_MIME,
                           key=f"pre_delete_reg_{ev['id']}", width="stretch")
    typed = st.text_input(f"Type the event name to confirm: {ev['event_name']}",
                          key=f"confirm_delete_{ev['id']}")
    confirmed = typed.strip().casefold() == ev["event_name"].strip().casefold()
    if st.button("🗑️  Delete this event", type="primary", disabled=not confirmed,
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
        ["📝 Pre-registration link", "📲 Check-in QR & poster", "✏️ Edit details", "🖼️ Flyer", "🗑️ Delete"])
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


def attendance_tab():
    events = get_events()
    if not events:
        st.info("No events yet.")
        return
    ev = event_picker("Event", "ledger_event", events)

    with st.expander("✍️  Manual check-in (for someone without a phone)"):
        with st.form(f"manual_checkin_{ev['id']}", clear_on_submit=True):
            phone_raw = st.text_input("Member's registered phone number")
            go = st.form_submit_button("Check them in", type="primary")
        if go:
            phone = normalize_phone(phone_raw)
            member_id = find_member_id(phone) if is_valid_phone(phone) else None
            if member_id is None:
                st.error("No member has that number. New members need to register on the check-in page.")
            elif record_check_in(member_id, ev["id"]):
                st.success(f"{get_member(member_id)['full_name']} is checked in.")
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
        st.dataframe(view, width="stretch", hide_index=True)
        st.caption(f"Arrival times are recorded by the server in GMT (Accra time). "
                   f"Last refreshed {datetime.now(timezone.utc):%Y-%m-%d %H:%M:%S}.")
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
        if st.button("🗑️  Remove member", type="primary", key=f"delete_member_{member_id}",
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
    c2.download_button("⬇️  Export all (Excel)", to_excel_bytes(everyone, "Members"),
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
        st.session_state.backup_name = f"ignite_backup_{datetime.now():%Y-%m-%d_%H%M}.db"
    if st.session_state.get("backup_bytes"):
        st.download_button("⬇️  Download backup", st.session_state.backup_bytes,
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
            for key in ("manage_event", "ledger_event", "prereg_event", "member_pick",
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

    c1, c2 = st.columns([5, 1])
    with c1:
        brand_header("Admin Portal", compact=True)
    with c2:
        st.write("")
        if st.button("Log out", width="stretch"):
            leave_admin()

    notice = st.session_state.pop("admin_notice", None)
    if notice:
        st.success(notice)

    tabs = st.tabs(["📊 Overview", "🎟️ Events", "📝 Pre-registrations", "📋 Check-ins", "👥 Members", "💾 Backup"])
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
        backup_tab()


# ---------------------------------------------------------------------------
# Router
# ---------------------------------------------------------------------------
def main():
    st.session_state.setdefault("form_nonce", 0)
    st.session_state.setdefault("show_admin", False)
    logged_in = st.session_state.get("admin_expires", 0) > time.time()
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
