"""
Ignite Prayer Network: QR Event Enrollment & Check-In System
=============================================================

One Streamlit app with two views:

  * Public check-in (default)   https://<your-app>.streamlit.app/?event=<id>
  * Admin portal (hidden)       https://<your-app>.streamlit.app/?view=admin

Every event (Asteri, Shekinah Glory, Impromptu) gets its own event_id.
Attendance rows are always written and read with that event_id, so one
event's ledger can never show another event's records.

Configuration (optional, via Streamlit secrets or .streamlit/secrets.toml):
    ADMIN_PASSWORD = "your-strong-password"
    APP_BASE_URL   = "https://ignite-checkin.streamlit.app"
If no secret is set, the defaults below are used.

Each event can carry a programme flyer (JPEG/PNG). It is stored inside the
database itself (events.flyer_bytes), so it travels with the data and there
are no loose image files to lose.
"""

import hmac
import io
import re
import sqlite3
import time
from contextlib import closing
from datetime import datetime, timezone

import pandas as pd
import qrcode
import streamlit as st
from PIL import Image, ImageOps, UnidentifiedImageError

# ---------------------------------------------------------------------------
# Settings: change these here, or override them with Streamlit secrets
# ---------------------------------------------------------------------------
DB_PATH = "ignite_network.db"
DEFAULT_ADMIN_PASSWORD = "IgniteAdmin2026"   # change before going live
EVENT_TYPES = ["Asteri", "Shekinah Glory", "Impromptu"]
SUCCESS_SECONDS = 2            # how long the welcome screen stays up
ADMIN_SESSION_MINUTES = 30     # admin is logged out after this much idle time
MAX_LOGIN_ATTEMPTS = 5         # wrong passwords before a cool-down
LOCKOUT_SECONDS = 60
FLYER_MAX_UPLOAD_MB = 10       # biggest flyer file an admin can upload
FLYER_MAX_WIDTH = 1200         # flyers are resized to this width to load fast on phones

st.set_page_config(
    page_title="Ignite Prayer Network | Check-In",
    page_icon="🔥",
    layout="centered",
    initial_sidebar_state="collapsed",
)


def get_setting(key: str, default: str) -> str:
    """Read a value from Streamlit secrets, falling back to the default."""
    try:
        return str(st.secrets[key])
    except Exception:
        return default


# ---------------------------------------------------------------------------
# Database layer
# ---------------------------------------------------------------------------
def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


@st.cache_resource
def init_db() -> bool:
    """Create tables once per server process. Safe to call repeatedly."""
    with closing(get_conn()) as conn, conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS events (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                event_name  TEXT NOT NULL,
                event_type  TEXT NOT NULL
                            CHECK (event_type IN ('Asteri', 'Shekinah Glory', 'Impromptu')),
                created_at  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                flyer_bytes BLOB            -- programme flyer, stored as a JPEG
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
                created_at              DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
            );

            -- The timestamp is set by the database, never by the browser.
            -- UNIQUE stops the same person being logged twice for one event,
            -- so the first arrival time is the one that stands.
            CREATE TABLE IF NOT EXISTS attendance (
                id                  INTEGER PRIMARY KEY AUTOINCREMENT,
                member_id           INTEGER NOT NULL REFERENCES members(id),
                event_id            INTEGER NOT NULL REFERENCES events(id),
                check_in_timestamp  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                UNIQUE (member_id, event_id)
            );

            CREATE INDEX IF NOT EXISTS idx_attendance_event ON attendance(event_id);

            -- Attendance is append-only. Nobody can quietly edit an arrival time.
            CREATE TRIGGER IF NOT EXISTS attendance_no_update
            BEFORE UPDATE ON attendance
            BEGIN
                SELECT RAISE(ABORT, 'Attendance records are read-only');
            END;
            """
        )
        # Upgrade path: a database created before the flyer feature has no
        # flyer_bytes column yet, so add it without touching existing rows.
        columns = {row["name"] for row in conn.execute("PRAGMA table_info(events)")}
        if "flyer_bytes" not in columns:
            conn.execute("ALTER TABLE events ADD COLUMN flyer_bytes BLOB")
    return True


def normalize_phone(raw: str) -> str:
    """Keep digits only and convert +233XXXXXXXXX to 0XXXXXXXXX, so one
    number is always stored the same way however it was typed."""
    digits = re.sub(r"\D", "", raw or "")
    if digits.startswith("233") and len(digits) == 12:
        digits = "0" + digits[3:]
    return digits


def is_valid_phone(phone: str) -> bool:
    return 9 <= len(phone) <= 15


def get_events() -> list[dict]:
    """Plain dicts (not sqlite rows) so Streamlit widgets can hold them."""
    with closing(get_conn()) as conn:
        rows = conn.execute(
            """SELECT id, event_name, event_type, created_at,
                      flyer_bytes IS NOT NULL AS has_flyer
               FROM events ORDER BY id DESC"""
        ).fetchall()
        return [dict(r) for r in rows]


def get_event(event_id: int):
    with closing(get_conn()) as conn:
        row = conn.execute(
            "SELECT id, event_name, event_type FROM events WHERE id = ?", (event_id,)
        ).fetchone()
        return dict(row) if row else None


def create_event(name: str, event_type: str, flyer: bytes | None = None) -> int:
    with closing(get_conn()) as conn, conn:
        cur = conn.execute(
            "INSERT INTO events (event_name, event_type, flyer_bytes) VALUES (?, ?, ?)",
            (name.strip(), event_type, flyer),
        )
    get_flyer.clear()
    return cur.lastrowid


def set_event_flyer(event_id: int, flyer: bytes | None):
    """Replace an event's flyer, or remove it by passing None."""
    with closing(get_conn()) as conn, conn:
        conn.execute("UPDATE events SET flyer_bytes = ? WHERE id = ?", (flyer, event_id))
    get_flyer.clear()


@st.cache_data(ttl=600, show_spinner=False)
def get_flyer(event_id: int) -> bytes | None:
    """Cached, so a queue of people checking in doesn't reload the image
    from the database on every tap. Cleared whenever a flyer changes."""
    with closing(get_conn()) as conn:
        row = conn.execute(
            "SELECT flyer_bytes FROM events WHERE id = ?", (event_id,)
        ).fetchone()
        return bytes(row["flyer_bytes"]) if row and row["flyer_bytes"] else None


def find_member_id(phone: str):
    with closing(get_conn()) as conn:
        row = conn.execute(
            "SELECT id FROM members WHERE phone_number = ?", (phone,)
        ).fetchone()
        return row["id"] if row else None


def create_member(data: dict) -> int:
    with closing(get_conn()) as conn, conn:
        cur = conn.execute(
            """INSERT INTO members (full_name, phone_number, whatsapp_number, email,
                                    emergency_contact_name, emergency_contact_phone,
                                    parent_guardian_phone)
               VALUES (:full_name, :phone_number, :whatsapp_number, :email,
                       :emergency_contact_name, :emergency_contact_phone,
                       :parent_guardian_phone)""",
            data,
        )
        return cur.lastrowid


def record_check_in(member_id: int, event_id: int) -> bool:
    """Returns True for a new check-in, False if already checked in."""
    try:
        with closing(get_conn()) as conn, conn:
            conn.execute(
                "INSERT INTO attendance (member_id, event_id) VALUES (?, ?)",
                (member_id, event_id),
            )
        return True
    except sqlite3.IntegrityError:
        return False


def get_ledger(event_id: int) -> pd.DataFrame:
    """Attendance for ONE event only. The WHERE clause is what keeps
    Asteri, Shekinah Glory and impromptu records apart."""
    with closing(get_conn()) as conn:
        return pd.read_sql_query(
            """SELECT strftime('%Y-%m-%d %H:%M:%S', a.check_in_timestamp) AS "Check-In Time",
                      m.full_name               AS "Full Name",
                      m.phone_number            AS "Phone",
                      m.whatsapp_number         AS "WhatsApp",
                      m.email                   AS "Email",
                      m.emergency_contact_name  AS "Emergency Contact",
                      m.emergency_contact_phone AS "Emergency Phone",
                      m.parent_guardian_phone   AS "Parent/Guardian Phone"
               FROM attendance a
               JOIN members m ON m.id = a.member_id
               WHERE a.event_id = ?
               ORDER BY a.check_in_timestamp ASC""",
            conn,
            params=(event_id,),
        )


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def make_qr_png(data: str) -> bytes:
    qr = qrcode.QRCode(
        error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=10, border=4
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def prepare_flyer(uploaded_file) -> bytes:
    """Check the upload really is an image, straighten phone photos, flatten
    transparency onto white, shrink it for mobile and save it as a JPEG.
    Raises ValueError with a friendly message if something is wrong."""
    if uploaded_file.size > FLYER_MAX_UPLOAD_MB * 1024 * 1024:
        raise ValueError(f"That file is over {FLYER_MAX_UPLOAD_MB} MB. Please use a smaller image.")
    try:
        img = Image.open(uploaded_file)
        img.load()
    except (UnidentifiedImageError, OSError, Image.DecompressionBombError):
        raise ValueError("That file doesn't look like a valid JPEG or PNG image.")

    img = ImageOps.exif_transpose(img)  # fixes sideways photos taken on phones
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGBA")
        background = Image.new("RGB", img.size, "white")
        background.paste(img, mask=img.getchannel("A"))
        img = background
    else:
        img = img.convert("RGB")

    if img.width > FLYER_MAX_WIDTH:
        new_height = round(img.height * FLYER_MAX_WIDTH / img.width)
        img = img.resize((FLYER_MAX_WIDTH, new_height), Image.LANCZOS)

    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85, optimize=True, progressive=True)
    return buf.getvalue()


def event_label(ev) -> str:
    return f"{ev['event_type']}: {ev['event_name']} (#{ev['id']})"


def to_excel_bytes(df: pd.DataFrame, sheet: str) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet[:31])
    return buf.getvalue()


def safe_filename(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]+", "_", text).strip("_")


def inject_css():
    st.markdown(
        """
        <style>
          #MainMenu, footer, [data-testid="stToolbar"] {visibility: hidden;}
          .ignite-brand {text-align:center; margin-bottom:0.5rem;}
          .ignite-brand h1 {font-size:2rem; margin:0; color:#E4572E;}
          .ignite-brand p  {margin:0.2rem 0 0; opacity:0.75;}
          .ignite-success {text-align:center; padding:3rem 1rem; border-radius:16px;
                           background:#1B998B; color:white;}
          .ignite-success h2 {color:white; font-size:2rem; margin-bottom:0.5rem;}
          .ignite-info {text-align:center; padding:3rem 1rem; border-radius:16px;
                        background:#F3A712; color:#222;}
          /* Flyer banner: full width, rounded, and never taller than about
             half the phone screen, so the form is still visible below it */
          .st-key-flyer_banner img {
              width:100%; max-height:55vh; object-fit:contain;
              border-radius:14px; box-shadow:0 4px 18px rgba(0,0,0,0.15);}
          .st-key-flyer_banner {margin-bottom:0.75rem;}
          div.stButton > button, div.stFormSubmitButton > button {
              width:100%; padding:0.75rem; font-size:1.1rem; font-weight:600;}
        </style>
        """,
        unsafe_allow_html=True,
    )


def brand_header(subtitle: str = ""):
    st.markdown(
        f"""<div class="ignite-brand">
              <h1>🔥 Ignite Prayer Network</h1>
              <p>{subtitle}</p>
            </div>""",
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Public check-in view
# ---------------------------------------------------------------------------
def flash_and_reset(kind: str, message: str):
    """Store the message, bump the form counter (new widget keys = blank
    forms) and rerun. The next run shows ONLY the message, then the form."""
    st.session_state.flash = (kind, message)
    st.session_state.form_nonce += 1
    st.rerun()


def show_flash_if_any():
    """Privacy screen: while the message is up, nothing else is drawn."""
    if "flash" not in st.session_state:
        return
    kind, message = st.session_state.pop("flash")
    css_class = "ignite-success" if kind == "success" else "ignite-info"
    title = "Success! Welcome to Ignite Network" if kind == "success" else "Already checked in"
    placeholder = st.empty()
    placeholder.markdown(
        f"""<div class="{css_class}"><h2>{title}</h2><p>{message}</p></div>""",
        unsafe_allow_html=True,
    )
    time.sleep(SUCCESS_SECONDS)
    placeholder.empty()
    st.rerun()


def resolve_event():
    """Take the event from ?event=<id>, or let the member pick one."""
    raw = st.query_params.get("event")
    if raw and str(raw).isdigit():
        ev = get_event(int(raw))
        if ev:
            return ev
        st.warning("That check-in link isn't valid. Please pick your event below.")

    events = get_events()
    if not events:
        st.info("No events are open for check-in yet.")
        return None
    return st.selectbox("Select your event", events, format_func=event_label)


def show_flyer(container, event_id: int):
    flyer = get_flyer(event_id)
    if flyer:
        container.image(flyer, width="stretch")


def checkin_page():
    show_flash_if_any()

    # Reserve the top slot first. The flyer is drawn into it once we know
    # which event this is, so it sits above everything else on the page.
    banner = st.container(key="flyer_banner")
    brand_header("Event Check-In")

    ev = resolve_event()
    if ev is None:
        return
    show_flyer(banner, ev["id"])
    st.markdown(
        f"<p style='text-align:center'><b>{ev['event_name']}</b> · {ev['event_type']}</p>",
        unsafe_allow_html=True,
    )

    nonce = st.session_state.form_nonce
    tab_existing, tab_new = st.tabs(["✅ Existing Member", "📝 New Member"])

    # Quick check-in with just a phone number
    with tab_existing:
        with st.form(key=f"quick_{nonce}"):
            phone_raw = st.text_input(
                "Your phone number", placeholder="e.g. 024 123 4567",
                autocomplete="off", key=f"q_phone_{nonce}",
            )
            submitted = st.form_submit_button("Submit Check-In")
        if submitted:
            phone = normalize_phone(phone_raw)
            member_id = find_member_id(phone) if is_valid_phone(phone) else None
            if member_id is None:
                st.error("We couldn't find that number. Please use the New Member tab.")
            elif record_check_in(member_id, ev["id"]):
                flash_and_reset("success", "Your attendance has been recorded.")
            else:
                flash_and_reset("info", "You're already checked in for this event.")

    # Full registration for first-timers
    with tab_new:
        with st.form(key=f"register_{nonce}"):
            full_name = st.text_input("Full name *", autocomplete="off", key=f"n_name_{nonce}")
            phone_raw = st.text_input("Phone number *", autocomplete="off", key=f"n_phone_{nonce}")
            whatsapp_raw = st.text_input(
                "WhatsApp number (if different)", autocomplete="off", key=f"n_wa_{nonce}"
            )
            email = st.text_input("Email (optional)", autocomplete="off", key=f"n_email_{nonce}")
            ec_name = st.text_input("Emergency contact name *", autocomplete="off", key=f"n_ecn_{nonce}")
            ec_phone_raw = st.text_input("Emergency contact phone *", autocomplete="off", key=f"n_ecp_{nonce}")
            parent_raw = st.text_input(
                "Parent/guardian phone (required for under-18s)",
                autocomplete="off", key=f"n_par_{nonce}",
            )
            submitted_new = st.form_submit_button("Submit Check-In")

        if submitted_new:
            phone = normalize_phone(phone_raw)
            ec_phone = normalize_phone(ec_phone_raw)
            parent = normalize_phone(parent_raw)
            whatsapp = normalize_phone(whatsapp_raw) or phone

            errors = []
            if len(full_name.strip()) < 2:
                errors.append("Please enter your full name.")
            if not is_valid_phone(phone):
                errors.append("Please enter a valid phone number.")
            if not ec_name.strip():
                errors.append("Please enter an emergency contact name.")
            if not is_valid_phone(ec_phone):
                errors.append("Please enter a valid emergency contact phone.")
            if parent and not is_valid_phone(parent):
                errors.append("The parent/guardian phone doesn't look right.")
            if email and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email.strip()):
                errors.append("The email address doesn't look right.")

            if errors:
                for e in errors:
                    st.error(e)
            else:
                member_id = find_member_id(phone)
                if member_id is None:
                    member_id = create_member({
                        "full_name": full_name.strip(),
                        "phone_number": phone,
                        "whatsapp_number": whatsapp,
                        "email": email.strip() or None,
                        "emergency_contact_name": ec_name.strip(),
                        "emergency_contact_phone": ec_phone,
                        "parent_guardian_phone": parent or None,
                    })
                    msg = "You're registered and checked in."
                else:
                    # Number already on file: check them in, don't overwrite anything
                    msg = "This number was already registered, so we've checked you in."
                if record_check_in(member_id, ev["id"]):
                    flash_and_reset("success", msg)
                else:
                    flash_and_reset("info", "You're already checked in for this event.")


# ---------------------------------------------------------------------------
# Admin portal
# ---------------------------------------------------------------------------
def admin_is_authenticated() -> bool:
    expires = st.session_state.get("admin_expires", 0)
    if expires > time.time():
        # Sliding timeout: any activity keeps the session alive
        st.session_state.admin_expires = time.time() + ADMIN_SESSION_MINUTES * 60
        return True
    st.session_state.pop("admin_expires", None)
    return False


def admin_login():
    brand_header("Admin Portal")
    locked_until = st.session_state.get("locked_until", 0)
    if locked_until > time.time():
        st.error(f"Too many attempts. Try again in {int(locked_until - time.time())} seconds.")
        return

    with st.form("admin_login", clear_on_submit=True):
        pw = st.text_input("Admin password", type="password")
        go = st.form_submit_button("Unlock")

    if go:
        expected = get_setting("ADMIN_PASSWORD", DEFAULT_ADMIN_PASSWORD)
        if hmac.compare_digest(pw.encode(), expected.encode()):
            st.session_state.admin_expires = time.time() + ADMIN_SESSION_MINUTES * 60
            st.session_state.failed_logins = 0
            st.rerun()
        else:
            time.sleep(1)  # slows down guessing
            st.session_state.failed_logins = st.session_state.get("failed_logins", 0) + 1
            if st.session_state.failed_logins >= MAX_LOGIN_ATTEMPTS:
                st.session_state.locked_until = time.time() + LOCKOUT_SECONDS
                st.session_state.failed_logins = 0
            st.error("Incorrect password.")


def event_manager_tab():
    st.subheader("Create an event")
    with st.form("new_event", clear_on_submit=True):
        ev_type = st.selectbox("Event type", EVENT_TYPES)
        ev_name = st.text_input(
            "Event name", placeholder=f"e.g. Asteri {datetime.now().year + 1}"
        )
        flyer_file = st.file_uploader(
            "Programme flyer (optional, JPEG or PNG)",
            type=["jpg", "jpeg", "png"],
            help="Shown at the top of the check-in page for this event.",
        )
        create = st.form_submit_button("Create event")
    if create:
        name = ev_name.strip() or f"{ev_type} {datetime.now():%Y-%m-%d}"
        flyer, flyer_ok = None, True
        if flyer_file is not None:
            try:
                flyer = prepare_flyer(flyer_file)
            except ValueError as err:
                flyer_ok = False
                st.error(f"{err} The event was not created.")
        if flyer_ok:
            new_id = create_event(name, ev_type, flyer)
            extra = " with its flyer" if flyer else ""
            st.success(f"Created '{name}' (event #{new_id}){extra}.")

    st.divider()
    st.subheader("QR codes")
    base_url = st.text_input(
        "Public app URL",
        value=st.session_state.get(
            "base_url", get_setting("APP_BASE_URL", "https://your-app.streamlit.app")
        ),
        help="The address of this app once deployed. Set APP_BASE_URL in secrets to fix it permanently.",
    ).rstrip("/")
    st.session_state.base_url = base_url

    events = get_events()
    if not events:
        st.info("Create an event first.")
        return
    ev = st.selectbox("Event", events, format_func=event_label, key="qr_event")
    link = f"{base_url}/?event={ev['id']}"
    png = make_qr_png(link)

    col1, col2 = st.columns([1, 1])
    with col1:
        st.image(png, caption=ev["event_name"], width=260)
    with col2:
        st.code(link, language=None)
        st.download_button(
            "Download QR (PNG)", png,
            file_name=f"QR_{safe_filename(ev['event_name'])}.png", mime="image/png",
        )
        st.caption("Print this and place it at the entrance. Each QR code opens check-in for this event only.")

    flyer_manager(ev)


def flyer_manager(ev: dict):
    """Add, replace or remove the flyer for an event that already exists."""
    st.divider()
    st.subheader("Programme flyer")
    current = get_flyer(ev["id"])
    if current:
        st.image(current, caption="Current flyer (as members will see it)", width=260)
    else:
        st.caption("This event has no flyer yet.")

    with st.form(f"flyer_form_{ev['id']}", clear_on_submit=True):
        new_file = st.file_uploader(
            "Upload a new flyer" if current else "Upload a flyer",
            type=["jpg", "jpeg", "png"],
        )
        save = st.form_submit_button("Save flyer")
    if save:
        if new_file is None:
            st.warning("Choose an image first.")
        else:
            try:
                set_event_flyer(ev["id"], prepare_flyer(new_file))
                st.success("Flyer saved.")
                st.rerun()
            except ValueError as err:
                st.error(str(err))

    if current and st.button("Remove flyer", key=f"remove_flyer_{ev['id']}"):
        set_event_flyer(ev["id"], None)
        st.rerun()


def ledger_tab():
    events = get_events()
    if not events:
        st.info("No events yet.")
        return
    ev = st.selectbox("Choose the event ledger to open", events, format_func=event_label, key="ledger_event")
    auto = st.toggle("Auto-refresh every 15 seconds", value=True)

    def render_ledger():
        df = get_ledger(ev["id"])
        st.metric("Checked in", len(df))
        search = st.text_input("Search by name or phone", key="ledger_search")
        view = df
        if search:
            mask = (
                df["Full Name"].str.contains(search, case=False, na=False)
                | df["Phone"].str.contains(re.sub(r"\D", "", search) or search, na=False)
            )
            view = df[mask]
        st.dataframe(view, width="stretch", hide_index=True)
        st.caption(
            f"Times are recorded by the server in GMT (Accra time). "
            f"Last refreshed {datetime.now(timezone.utc):%Y-%m-%d %H:%M:%S}."
        )

        stem = f"{safe_filename(ev['event_name'])}_attendance"
        c1, c2 = st.columns(2)
        c1.download_button(
            "Export CSV", view.to_csv(index=False).encode("utf-8"),
            file_name=f"{stem}.csv", mime="text/csv",
        )
        c2.download_button(
            "Export Excel", to_excel_bytes(view, ev["event_name"]),
            file_name=f"{stem}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # Fragment reruns only this block, so the rest of the page stays still
    st.fragment(run_every=15 if auto else None)(render_ledger)()


def admin_page():
    if not admin_is_authenticated():
        admin_login()
        return

    brand_header("Admin Portal")
    if st.button("Log out"):
        st.session_state.pop("admin_expires", None)
        st.rerun()

    tab_events, tab_ledger = st.tabs(["🎟️ Event Manager", "📋 Attendance Ledger"])
    with tab_events:
        event_manager_tab()
    with tab_ledger:
        ledger_tab()


# ---------------------------------------------------------------------------
# Router
# ---------------------------------------------------------------------------
def main():
    init_db()
    inject_css()
    st.session_state.setdefault("form_nonce", 0)

    if st.query_params.get("view") == "admin":
        admin_page()
    else:
        checkin_page()


if __name__ == "__main__":
    main()
