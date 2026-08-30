"""
==============================================================================
 FOOD DELIVERY LOGISTICS, OPERATIONS & PAYROLL DASHBOARD
==============================================================================
A production-ready Streamlit + SQLite application for managing food delivery
driver operations, workforce KPIs, payroll, bulk data uploads, rider-level
lookup, and automated supervisor action-alert generation.

--------------------------------------------------------------------------
INSTALLATION
--------------------------------------------------------------------------
    pip install streamlit pandas openpyxl

--------------------------------------------------------------------------
EXECUTION
--------------------------------------------------------------------------
    streamlit run app.py

The app auto-creates `logistics.db` (SQLite) in the same folder on first
launch, and auto-upgrades older database files with any new columns added
in later versions of this app. Use the "Seed Sample Data" button in the
sidebar to populate the database with realistic demo data for testing.
==============================================================================
"""

import hashlib
import io
import zipfile
import json
import os
import random
import re
import smtplib
import sqlite3
from datetime import datetime, timedelta
from email.mime.text import MIMEText

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# gspread is optional -- only needed for the Google Sheets sync feature.
# The whole app still works without it (Excel upload etc. are
# unaffected); the Google Sheet sync panel just shows a clear "please
# install" message instead of crashing the app on import.
try:
    import gspread
    from google.oauth2.service_account import Credentials as _GCredentials
    GSPREAD_AVAILABLE = True
except ImportError:  # noqa: BLE001
    GSPREAD_AVAILABLE = False

# ==============================================================================
# CONFIG / CONSTANTS
# ==============================================================================
DB_PATH = "logistics.db"
OWNER_CONFIG_PATH = "hq_owner.json"
VALIDITY_TARGETS_PATH = "hq_validity_targets.json"
DEFAULT_MIN_ORDERS_FOR_VALID = 330
DEFAULT_MIN_DAYS_FOR_VALID = 26

DRIVER_STATUSES = ["Active", "Terminated", "Suspended"]
VEHICLE_TYPES = ["Company Car", "Own Car"]
VALIDITY_STATUSES = ["Valid", "Invalid"]

st.set_page_config(
    page_title="Delivery Logistics & Payroll Dashboard",
    page_icon="\U0001F6F5",
    layout="wide",
)

# ==============================================================================
# GLOBAL PROFESSIONAL STYLING / ANIMATIONS
# ==============================================================================


def inject_custom_css():
    """One-time global CSS polish: card-style metrics with hover lift,
    a gradient accent bar, smoother tab styling, and a soft fade-in-up
    entrance for key sections -- gives the whole dashboard a more
    'designed' feel without touching Streamlit internals that change
    between versions (we only target stable data-testid selectors)."""
    st.markdown(
        """
        <style>
          @keyframes fadeInUp {
            from { opacity: 0; transform: translateY(14px); }
            to   { opacity: 1; transform: translateY(0); }
          }
          @keyframes shimmerBar {
            0%   { background-position: 0% 50%; }
            50%  { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
          }

          .app-gradient-bar {
            height: 5px;
            border-radius: 6px;
            margin-bottom: 6px;
            background: linear-gradient(90deg, #22c55e, #3b82f6, #a855f7, #22c55e);
            background-size: 300% 300%;
            animation: shimmerBar 6s ease infinite;
          }

          [data-testid="stMetric"] {
            background: linear-gradient(180deg, rgba(34,197,94,0.06), rgba(59,130,246,0.05));
            border: 1px solid rgba(120,120,120,0.15);
            border-radius: 14px;
            padding: 12px 14px 8px 14px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.06);
            transition: transform 0.18s ease, box-shadow 0.18s ease;
            animation: fadeInUp 0.5s ease both;
          }
          [data-testid="stMetric"]:hover {
            transform: translateY(-4px);
            box-shadow: 0 10px 22px rgba(0,0,0,0.12);
          }

          [data-testid="stDataFrame"], [data-testid="stExpander"] {
            border-radius: 12px;
            overflow: hidden;
            animation: fadeInUp 0.6s ease both;
          }

          button[kind="primary"] {
            transition: transform 0.15s ease;
          }
          button[kind="primary"]:hover {
            transform: translateY(-2px) scale(1.01);
          }

          .stTabs [data-baseweb="tab-list"] {
            gap: 6px;
          }
          .stTabs [data-baseweb="tab"] {
            border-radius: 10px 10px 0 0;
            transition: background 0.15s ease;
          }

          @keyframes glowPulse {
            0%, 100% { box-shadow: 0 0 0 1px rgba(64,47,181,0.55), 0 0 18px 3px rgba(207,48,170,0.35), 0 0 34px 8px rgba(64,47,181,0.22); }
            50%      { box-shadow: 0 0 0 1px rgba(207,48,170,0.55), 0 0 22px 4px rgba(64,47,181,0.40), 0 0 40px 10px rgba(207,48,170,0.26); }
          }
          div[data-testid="stTextInput"] input {
            background: #0b0a12 !important;
            color: #f1f0ff !important;
            border: 1px solid rgba(120,110,180,0.35) !important;
            border-radius: 14px !important;
            padding: 12px 18px !important;
            font-size: 15px !important;
            transition: box-shadow 0.25s ease, border-color 0.25s ease;
          }
          div[data-testid="stTextInput"] input::placeholder { color: #9a94b8 !important; }
          div[data-testid="stTextInput"] input:focus {
            outline: none !important;
            border-color: rgba(207,48,170,0.6) !important;
            animation: glowPulse 2.4s ease-in-out infinite;
          }

          /* ---- Reusable animated hover-pop stat cards (used everywhere) ---- */
          .stat-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(190px, 1fr));
            gap: 16px;
            margin: 10px 0 18px 0;
          }
          .stat-card2 {
            position: relative;
            border-radius: 16px;
            padding: 18px 20px;
            border: 1px solid rgba(255,255,255,0.14);
            box-shadow: 0 4px 14px rgba(0,0,0,0.10);
            animation: fadeInUp 0.5s ease both;
            transition: transform 0.22s cubic-bezier(.2,.8,.3,1.3), box-shadow 0.22s ease;
            cursor: default;
          }
          .stat-card2.v-a { background: linear-gradient(135deg, rgba(34,197,94,0.15), rgba(59,130,246,0.08)); }
          .stat-card2.v-b { background: linear-gradient(135deg, rgba(168,85,247,0.15), rgba(59,130,246,0.08)); }
          .stat-card2.v-c { background: linear-gradient(135deg, rgba(239,68,68,0.15), rgba(249,115,22,0.08)); }
          .stat-card2.v-d { background: linear-gradient(135deg, rgba(250,204,21,0.15), rgba(34,197,94,0.08)); }
          .stat-card2:hover {
            transform: translateY(-9px) scale(1.05);
            box-shadow: 0 20px 38px rgba(0,0,0,0.26);
            z-index: 6;
          }
          .stat-icon2 { font-size: 21px; line-height: 1; }
          .stat-label2 {
            font-size: 11.5px; letter-spacing: 1.1px; text-transform: uppercase;
            opacity: 0.64; margin-top: 7px; font-weight: 700;
          }
          .stat-value2 { font-size: 25px; font-weight: 900; margin-top: 3px; }
          .stat-tip {
            position: absolute; left: 50%; bottom: calc(100% + 10px);
            transform: translate(-50%, 6px);
            background: #111018; color: #f3f1ff;
            padding: 8px 13px; border-radius: 9px; font-size: 11.5px;
            white-space: nowrap; opacity: 0; pointer-events: none;
            box-shadow: 0 12px 24px rgba(0,0,0,0.4);
            transition: opacity 0.18s ease, transform 0.18s ease;
          }
          .stat-tip::after {
            content: ""; position: absolute; top: 100%; left: 50%;
            transform: translateX(-50%);
            border: 6px solid transparent; border-top-color: #111018;
          }
          .stat-card2:hover .stat-tip { opacity: 1; transform: translate(-50%, 0); }

          /* ---- Clickable stat cards: the button itself IS the card
             (styled to match the original stat-card2 look), so a click
             always lands on the real control -- no invisible overlay
             positioning that could go wrong. ---- */
          div[class*="st-key-clickcard_"] div[data-testid="stButton"] button {
            width: 100%;
            min-height: 104px;
            border-radius: 16px !important;
            border: 1px solid rgba(255,255,255,0.14) !important;
            box-shadow: 0 4px 14px rgba(0,0,0,0.10);
            padding: 14px 10px !important;
            white-space: pre-line;
            line-height: 1.5;
            font-weight: 700;
            transition: transform 0.22s cubic-bezier(.2,.8,.3,1.3), box-shadow 0.22s ease;
          }
          div[class*="st-key-clickcard_"] div[data-testid="stButton"] button:hover {
            transform: translateY(-6px) scale(1.03);
            box-shadow: 0 16px 30px rgba(0,0,0,0.22);
          }
          div[class*="__va"] div[data-testid="stButton"] button {
            background: linear-gradient(135deg, rgba(34,197,94,0.18), rgba(59,130,246,0.10)) !important;
          }
          div[class*="__vb"] div[data-testid="stButton"] button {
            background: linear-gradient(135deg, rgba(168,85,247,0.18), rgba(59,130,246,0.10)) !important;
          }
          div[class*="__vc"] div[data-testid="stButton"] button {
            background: linear-gradient(135deg, rgba(239,68,68,0.16), rgba(249,115,22,0.10)) !important;
          }
          div[class*="__vd"] div[data-testid="stButton"] button {
            background: linear-gradient(135deg, rgba(250,204,21,0.18), rgba(34,197,94,0.10)) !important;
          }
          div[class*="__active"] div[data-testid="stButton"] button {
            outline: 2px solid rgba(59,130,246,0.9) !important;
            box-shadow: 0 0 0 5px rgba(59,130,246,0.18), 0 20px 38px rgba(0,0,0,0.26) !important;
          }

          /* ---- Drilldown reveal panel (appears under whichever row
             the clicked card was in) ---- */
          div[class*="st-key-drill_panel_box"] {
            border-radius: 16px;
            padding: 14px 18px 4px 18px;
            margin: 6px 0 20px 0;
            border: 1px solid rgba(59,130,246,0.35);
            background: linear-gradient(135deg, rgba(59,130,246,0.10), rgba(34,197,94,0.06));
            animation: drillIn 0.32s cubic-bezier(.2,.8,.3,1.15) both;
          }
          @keyframes drillIn {
            from { opacity: 0; transform: translateY(-12px) scaleY(0.97); }
            to   { opacity: 1; transform: translateY(0) scaleY(1); }
          }
        </style>
        <div class="app-gradient-bar"></div>
        """,
        unsafe_allow_html=True,
    )


def stat_cards(cards: list) -> None:
    """Render a responsive grid of animated hover-pop stat cards.
    Each card lifts, scales, and reveals a small tooltip on hover --
    every dict may include: icon, label, value, tip (tooltip text), variant
    ('a'/'b'/'c'/'d' for different accent colors)."""
    html = ['<div class="stat-grid">']
    for c in cards:
        variant = c.get("variant", "a")
        tip_html = f'<div class="stat-tip">{c["tip"]}</div>' if c.get("tip") else ""
        html.append(
            f'<div class="stat-card2 v-{variant}">'
            f'<div class="stat-icon2">{c.get("icon", "")}</div>'
            f'<div class="stat-label2">{c["label"]}</div>'
            f'<div class="stat-value2">{c["value"]}</div>'
            f"{tip_html}"
            f"</div>"
        )
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)


def render_clickable_stat_row(cards: list, row_key: str, state_key: str = "active_drill") -> None:
    """Real, clickable stat cards -- each card IS an actual Streamlit
    button (styled via CSS to keep the same gradient/icon/rounded-card
    look), not a decorative div with a separate invisible button placed
    behind it. A click always lands on the real control this way --
    that's the fix for the earlier version, where the invisible overlay
    button could end up misaligned with the pretty card drawn on top of
    it and silently eat clicks.

    state_key: which session_state key tracks the active card. Defaults
    to "active_drill" (Operations Dashboard's original behavior,
    unchanged); pass a different key for an independent set of cards
    elsewhere (e.g. Rider Lookup) so the two don't fight over the same
    "currently open" state."""
    cols = st.columns(len(cards))
    for i, c in enumerate(cards):
        card_id = f"{row_key}_{i}"
        variant = c.get("variant", "a")
        is_active = st.session_state.get(state_key) == card_id
        label = f"{c.get('icon', '')}\n\n{c['label']}\n\n**{c['value']}**"
        active_flag = "__active" if is_active else ""
        with cols[i]:
            with st.container(key=f"clickcard_{state_key}_{card_id}__v{variant}{active_flag}"):
                clicked = st.button(label, key=f"clickbtn_{state_key}_{card_id}", use_container_width=True)
                if c.get("tip"):
                    st.caption(c["tip"])
                if clicked:
                    st.session_state[state_key] = None if is_active else card_id
                    st.rerun()


def render_drilldown_panel(drill_defs: dict) -> None:
    """Shows an animated reveal panel for whichever card the user just
    clicked (tracked in st.session_state['active_drill']), IF that card
    belongs to the group currently on screen (drill_defs). drill_defs is
    {card_id: (title, dataframe, note)} -- always built fresh from
    whatever month/filters are active, so this works the same for every
    month, not just one hardcoded snapshot."""
    active = st.session_state.get("active_drill")
    if not active or active not in drill_defs:
        return
    title, df, note = drill_defs[active]
    with st.container(key="drill_panel_box"):
        c1, c2 = st.columns([6, 1])
        c1.markdown(f"##### \U0001F50E {title}  \u2022  {len(df)} row(s)")
        if c2.button("\u2715 Close", key="close_drill_btn", use_container_width=True):
            st.session_state["active_drill"] = None
            st.rerun()
        if note:
            st.caption(note)
        if df.empty:
            st.caption("Nothing to show here for the current month/filters.")
        else:
            st.dataframe(df, use_container_width=True, hide_index=True)


# ==============================================================================
# HQ ACCESS CONTROL  -- one Admin (you, email+password) + tracked Viewers
# ==============================================================================


def _hash_password(password: str, salt: str) -> str:
    return hashlib.sha256((salt + password).encode("utf-8")).hexdigest()


def _load_owner_config():
    if not os.path.exists(OWNER_CONFIG_PATH):
        return None
    try:
        with open(OWNER_CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:  # noqa: BLE001
        return None


def _save_owner_config(email: str, password: str) -> dict:
    salt = os.urandom(16).hex()
    config = {
        "email": email.strip().lower(),
        "salt": salt,
        "password_hash": _hash_password(password, salt),
    }
    with open(OWNER_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f)
    return config


def _verify_owner(email: str, password: str, config: dict) -> bool:
    if not config:
        return False
    if email.strip().lower() != config.get("email", ""):
        return False
    return _hash_password(password, config.get("salt", "")) == config.get("password_hash", "")


def is_admin() -> bool:
    """True only for the signed-in HQ Admin -- use this to gate anything
    that uploads, edits, seeds, or clears data."""
    return st.session_state.get("auth_role") == "admin"


ONLINE_WINDOW_MINUTES = 5  # "online now" = active within this many minutes


def _ensure_viewer_table(conn) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS viewer_log (
            email      TEXT PRIMARY KEY,
            name       TEXT,
            first_seen TEXT,
            last_seen  TEXT,
            visits     INTEGER NOT NULL DEFAULT 1,
            status     TEXT NOT NULL DEFAULT 'pending'
        )
        """
    )
    cols = [r[1] for r in conn.execute("PRAGMA table_info(viewer_log)").fetchall()]
    if "blocked" in cols and "status" not in cols:
        conn.execute("ALTER TABLE viewer_log ADD COLUMN status TEXT NOT NULL DEFAULT 'pending'")
        conn.execute("UPDATE viewer_log SET status = CASE WHEN blocked = 1 THEN 'revoked' ELSE 'approved' END")
    elif "status" not in cols:
        conn.execute("ALTER TABLE viewer_log ADD COLUMN status TEXT NOT NULL DEFAULT 'pending'")


def _get_viewer_status(email: str):
    conn = get_connection()
    _ensure_viewer_table(conn)
    row = conn.execute("SELECT status FROM viewer_log WHERE email = ?", (email.strip().lower(),)).fetchone()
    conn.close()
    return row[0] if row else None


def _record_viewer_request(email: str, name: str) -> bool:
    conn = get_connection()
    _ensure_viewer_table(conn)
    email = email.strip().lower()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    existing = conn.execute("SELECT visits FROM viewer_log WHERE email = ?", (email,)).fetchone()
    is_new = existing is None
    if is_new:
        conn.execute(
            "INSERT INTO viewer_log (email, name, first_seen, last_seen, visits, status) "
            "VALUES (?, ?, ?, ?, 1, 'pending')",
            (email, name, now, now),
        )
    else:
        conn.execute(
            "UPDATE viewer_log SET name = ?, last_seen = ?, visits = visits + 1 WHERE email = ?",
            (name, now, email),
        )
    conn.commit()
    conn.close()
    return is_new


def _touch_viewer_visit(email: str, name: str) -> None:
    conn = get_connection()
    _ensure_viewer_table(conn)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn.execute(
        "UPDATE viewer_log SET name = ?, last_seen = ?, visits = visits + 1 WHERE email = ?",
        (name, now, email.strip().lower()),
    )
    conn.commit()
    conn.close()


def _set_viewer_status(email: str, status: str) -> None:
    conn = get_connection()
    _ensure_viewer_table(conn)
    conn.execute("UPDATE viewer_log SET status = ? WHERE email = ?", (status, email.strip().lower()))
    conn.commit()
    conn.close()


def _send_access_request_email(viewer_name: str, viewer_email: str) -> bool:
    try:
        email_cfg = st.secrets.get("email", {})
        sender = email_cfg.get("address")
        app_password = email_cfg.get("app_password")
        owner_config = _load_owner_config()
        admin_email = owner_config["email"] if owner_config else None
        if not (sender and app_password and admin_email):
            return False

        msg = MIMEText(
            f"{viewer_name} ({viewer_email}) just requested access to your "
            f"HQ dashboard.\n\n"
            f"Open the dashboard, check the sidebar's 'Access requests' "
            f"panel, and Approve or Deny them."
        )
        msg["Subject"] = "HQ Dashboard: new access request"
        msg["From"] = sender
        msg["To"] = admin_email

        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=10) as server:
            server.login(sender, app_password)
            server.sendmail(sender, [admin_email], msg.as_string())
        return True
    except Exception:  # noqa: BLE001 - never let an email problem break the request flow
        return False


def render_auth_gate() -> bool:
    if "auth_role" not in st.session_state:
        st.session_state.auth_role = None
        st.session_state.auth_email = None
        st.session_state.auth_name = None

    if st.session_state.auth_role == "admin":
        return True
    if st.session_state.auth_role == "viewer":
        status = _get_viewer_status(st.session_state.auth_email)
        if status == "approved":
            _touch_viewer_visit(st.session_state.auth_email, st.session_state.auth_name)
            return True
        if status == "pending":
            st.markdown(
                "<h1 style='text-align:center; margin-bottom:0;'>\U0001F3E2 HQ</h1>",
                unsafe_allow_html=True,
            )
            st.info(
                f"\u23F3 **Your access request is waiting for approval.**\n\n"
                f"{st.session_state.auth_name} ({st.session_state.auth_email}) "
                f"-- the HQ Admin has been notified. Check back shortly, or "
                f"press the button below to check again."
            )
            if st.button("\U0001F504 Check again"):
                st.rerun()
            return False
        st.session_state.auth_role = None
        st.error("Your access to this dashboard has been removed by the HQ Admin.")
        return False

    st.markdown(
        "<h1 style='text-align:center; margin-bottom:0;'>\U0001F3E2 HQ</h1>"
        "<p style='text-align:center; opacity:0.65; margin-top:0;'>"
        "Delivery Logistics &amp; Payroll Dashboard</p>",
        unsafe_allow_html=True,
    )

    config = _load_owner_config()

    if config is None:
        st.info(
            "\U0001F510 **First-time setup** -- create the HQ Admin account. "
            "Only this account can upload files, seed/clear data, or make "
            "any changes. Everyone else can view the dashboards but not "
            "edit anything."
        )
        st.caption(
            "Note: this is basic password protection meant for a trusted "
            "team, not enterprise-grade security -- don't reuse a "
            "sensitive password here."
        )
        with st.form("hq_setup_form"):
            email = st.text_input("Your email (this becomes the HQ Admin login)")
            pw1 = st.text_input("Choose a password", type="password")
            pw2 = st.text_input("Confirm password", type="password")
            submitted = st.form_submit_button("Create Admin Account", type="primary")
        if submitted:
            if not email or "@" not in email:
                st.error("Please enter a valid email address.")
            elif len(pw1) < 6:
                st.error("Password should be at least 6 characters.")
            elif pw1 != pw2:
                st.error("Passwords don't match.")
            else:
                _save_owner_config(email, pw1)
                st.session_state.auth_role = "admin"
                st.session_state.auth_email = email.strip().lower()
                st.success("Admin account created!")
                st.rerun()
        return False

    tab_admin, tab_viewer = st.tabs(["\U0001F511 Admin Sign In", "\U0001F64B Request Access"])

    with tab_admin:
        with st.form("hq_login_form"):
            email = st.text_input("Email")
            pw = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Sign In", type="primary")
        if submitted:
            if _verify_owner(email, pw, config):
                st.session_state.auth_role = "admin"
                st.session_state.auth_email = config["email"]
                st.rerun()
            else:
                st.error("Incorrect email or password.")

    with tab_viewer:
        st.write(
            "Approved viewers can see every dashboard, report, and rider "
            "lookup, but cannot upload files, seed demo data, clear data, "
            "or make any changes."
        )
        st.caption(
            "Your request goes to the HQ Admin for approval -- you won't "
            "see the dashboard until they approve you."
        )
        with st.form("hq_viewer_form"):
            v_name = st.text_input("Your name")
            v_email = st.text_input("Your email")
            submitted = st.form_submit_button("Request Access", use_container_width=True)
        if submitted:
            if not v_name.strip():
                st.error("Please enter your name.")
            elif not v_email or "@" not in v_email:
                st.error("Please enter a valid email address.")
            elif _get_viewer_status(v_email) == "revoked":
                st.error("This email's access has been removed by the HQ Admin.")
            else:
                is_new_request = _record_viewer_request(v_email, v_name.strip())
                if is_new_request:
                    _send_access_request_email(v_name.strip(), v_email.strip().lower())
                st.session_state.auth_role = "viewer"
                st.session_state.auth_email = v_email.strip().lower()
                st.session_state.auth_name = v_name.strip()
                st.rerun()

    return False


def render_hq_banner():
    config = _load_owner_config()
    owner_email = config["email"] if config else ""
    role = st.session_state.get("auth_role")
    if role == "admin":
        who = owner_email
        role_label = "Admin (edit access)"
    else:
        who = st.session_state.get("auth_email", "")
        role_label = "Viewer (read-only)"
    st.markdown(
        f"""
        <div style="display:flex; justify-content:flex-end; align-items:center;
                    gap:8px; padding:2px 4px 8px 0; flex-wrap:wrap;">
          <span style="font-weight:900; letter-spacing:2px; font-size:14px;">\U0001F3E2 HQ</span>
          <span style="opacity:0.4;">|</span>
          <span style="opacity:0.85; font-size:12.5px;">{who}</span>
          <span style="opacity:0.4;">|</span>
          <span style="opacity:0.7; font-size:12.5px;">{role_label}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_hq_access_panel():
    conn = get_connection()
    _ensure_viewer_table(conn)
    df = pd.read_sql_query(
        "SELECT email, name, first_seen, last_seen, visits, status FROM viewer_log ORDER BY last_seen DESC",
        conn,
    )
    conn.close()

    pending_df = df[df["status"] == "pending"]
    approved_df = df[df["status"] == "approved"]
    revoked_df = df[df["status"] == "revoked"]

    online_count = 0
    if not approved_df.empty:
        cutoff = datetime.now() - timedelta(minutes=ONLINE_WINDOW_MINUTES)
        last_seen_dt = pd.to_datetime(approved_df["last_seen"], errors="coerce")
        online_count = int((last_seen_dt >= cutoff).sum())

    if not pending_df.empty:
        with st.sidebar.expander(f"\U0001F64B Access requests ({len(pending_df)})", expanded=True):
            for _, row in pending_df.iterrows():
                st.markdown(f"**{row['name'] or '(no name)'}**  \n{row['email']}")
                st.caption(f"Requested {row['first_seen']}")
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("\u2705 Approve", key=f"approve_{row['email']}", use_container_width=True):
                        _set_viewer_status(row["email"], "approved")
                        st.rerun()
                with c2:
                    if st.button("\u274C Deny", key=f"deny_{row['email']}", use_container_width=True):
                        _set_viewer_status(row["email"], "revoked")
                        st.rerun()
                st.markdown("---")

    with st.sidebar.expander(
        f"\U0001F465 Who has access ({len(approved_df)}) \u2022 \U0001F7E2 {online_count} online", expanded=False
    ):
        if approved_df.empty:
            st.caption("No approved viewers yet.")
        else:
            cutoff = datetime.now() - timedelta(minutes=ONLINE_WINDOW_MINUTES)
            for _, row in approved_df.iterrows():
                last_seen_dt = pd.to_datetime(row["last_seen"], errors="coerce")
                is_online = bool(pd.notna(last_seen_dt) and last_seen_dt >= cutoff)
                status_label = "\U0001F7E2 Online" if is_online else "\u26AA Offline"

                st.markdown(f"**{row['name'] or '(no name)'}**  \n{row['email']}")
                st.caption(
                    f"{status_label} \u2022 First seen {row['first_seen']} \u2022 "
                    f"Last seen {row['last_seen']} \u2022 {row['visits']} visit(s)"
                )
                if st.button("\U0001F5D1\uFE0F Remove access", key=f"remove_{row['email']}", use_container_width=True):
                    _set_viewer_status(row["email"], "revoked")
                    st.rerun()
                st.markdown("---")
            st.caption(
                f"'Online' means active within the last {ONLINE_WINDOW_MINUTES} minutes. "
                "Names/emails are self-reported, not verified."
            )

    if not revoked_df.empty:
        with st.sidebar.expander(f"\U0001F6AB Removed / denied ({len(revoked_df)})", expanded=False):
            for _, row in revoked_df.iterrows():
                st.markdown(f"**{row['name'] or '(no name)'}**  \n{row['email']}")
                if st.button("\u21A9\uFE0F Restore access", key=f"restore_{row['email']}", use_container_width=True):
                    _set_viewer_status(row["email"], "approved")
                    st.rerun()
                st.markdown("---")


def _active_mask(df: pd.DataFrame) -> pd.Series:
    """A rider counts as 'Active' for a month when their employment
    status isn't Terminated/Suspended AND they show real activity that
    month -- days worked (from the validity/attendance sheets) or
    completed orders. A roster's own status field is often unreliable on
    its own (many months' files never mark anyone anything but Active
    unless they also have a note elsewhere), so activity is what tells
    apart someone genuinely working from someone just still listed."""
    worked = (df["valid_days_in_month"].fillna(0) > 0) | (df["total_orders"].fillna(0) > 0)
    not_terminated = ~df["status"].isin(["Terminated", "Suspended"])
    return not_terminated & worked


def render_header(filters: dict):
    merged = load_merged()

    headcount = active = 0
    orders_m = 0
    gross_m = 0.0

    if not merged.empty:
        if filters["month"]:
            month_df = merged[merged["month_year"] == filters["month"]]
            month_df = month_df[
                month_df["supervisor_name"].isin(filters["supervisors"])
                & month_df["vehicle_type"].isin(filters["vehicle_types"])
            ]
            # Headcount is scoped to riders actually listed on THIS
            # month's roster/active-rider sheet -- not every driver_id
            # that happens to have a log row this month (an orders or
            # validity sheet can reference IDs the roster snapshot never
            # had, and those shouldn't inflate headcount).
            roster_month_df = month_df[month_df["in_roster"] == 1]
            headcount = roster_month_df["driver_id"].nunique()
            active = int(roster_month_df[_active_mask(roster_month_df)]["driver_id"].nunique())
            orders_m = int(month_df["total_orders"].sum())
            gross_m = float(month_df["gross_salary"].sum())
        else:
            roster = merged.drop_duplicates(subset="driver_id")
            roster_f = roster[
                roster["supervisor_name"].isin(filters["supervisors"])
                & roster["vehicle_type"].isin(filters["vehicle_types"])
            ]
            headcount = len(roster_f)
            active = int((roster_f["status"] == "Active").sum())

    # A server-side default for the very first paint (Streamlit's
    # light/dark TOGGLE is purely client-side -- the server has no idea
    # which one the browser is actually showing unless it's hard-coded
    # in config.toml, so this Python guess can be wrong and is only a
    # starting point; the JS below is what actually keeps it correct).
    theme_base = st.get_option("theme.base") or "light"
    stat_text_color = st.get_option("theme.textColor") or (
        "#FAFAFA" if theme_base == "dark" else "#31333F"
    )

    extrusion_layers = ""
    layer_count = 7
    for i in range(layer_count, 0, -1):
        shade = 10 + i * 4
        extrusion_layers += (
            f'<g transform="translate({i * 2.6},{i * 3.2})" fill="rgb({shade},{shade},{shade})">'
            f'<path d="M25 25 L120 25 L120 62 L68 62 L68 98 L25 98 Z"/>'
            f'<path d="M62 25 L195 25 L195 62 L110 62 L152 98 L110 98 Z"/>'
            f'<rect x="78" y="98" width="28" height="142"/>'
            f'<rect x="114" y="98" width="28" height="142"/>'
            f"</g>"
        )

    html = """
    <style>
      .header-row {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 30px;
        flex-wrap: wrap;
        padding: 8px 0 4px 0;
        font-family: Arial, Helvetica, sans-serif;
      }
      .stat-stack { display: flex; flex-direction: column; gap: 14px; }
      .stat-card {
        position: relative;
        min-width: 210px;
        text-align: center;
        padding: 18px 26px;
        border-radius: 16px;
        border: 1px solid rgba(120,120,120,0.18);
        box-shadow: 0 6px 18px rgba(0,0,0,0.12);
        animation: statPop 0.5s ease both;
        transition: transform 0.22s cubic-bezier(.2,.8,.3,1.3), box-shadow 0.22s ease;
      }
      .stat-card:hover {
        transform: translateY(-8px) scale(1.05);
        box-shadow: 0 18px 34px rgba(0,0,0,0.28);
        z-index: 6;
      }
      .stat-card.v-a { background: linear-gradient(135deg, rgba(34,197,94,0.18), rgba(59,130,246,0.10)); }
      .stat-card.v-b { background: linear-gradient(135deg, rgba(168,85,247,0.18), rgba(59,130,246,0.10)); }
      .stat-card.v-c { background: linear-gradient(135deg, rgba(239,68,68,0.16), rgba(249,115,22,0.10)); }
      .stat-card.v-d { background: linear-gradient(135deg, rgba(250,204,21,0.18), rgba(34,197,94,0.10)); }
      .stat-card .label {
        font-size: 13px; letter-spacing: 1.4px; text-transform: uppercase;
        opacity: 0.64; font-weight: 700;
        color: var(--stat-text, __STAT_TEXT__);
      }
      .stat-card .value {
        font-size: 32px; font-weight: 900; margin-top: 4px;
        color: var(--stat-text, __STAT_TEXT__);
      }
      .stat-card .htip {
        position: absolute; left: 50%; bottom: calc(100% + 10px); transform: translate(-50%, 6px);
        background: #111018; color: #f3f1ff; padding: 8px 13px; border-radius: 9px; font-size: 11.5px;
        white-space: nowrap; opacity: 0; pointer-events: none; box-shadow: 0 12px 24px rgba(0,0,0,0.4);
        transition: opacity 0.18s ease, transform 0.18s ease;
      }
      .stat-card .htip::after {
        content: ""; position: absolute; top: 100%; left: 50%; transform: translateX(-50%);
        border: 6px solid transparent; border-top-color: #111018;
      }
      .stat-card:hover .htip { opacity: 1; transform: translate(-50%, 0); }
      @keyframes statPop {
        from { opacity: 0; transform: translateY(10px) scale(0.96); }
        to   { opacity: 1; transform: translateY(0) scale(1); }
      }

      .logo-stage { perspective: 1800px; }
      .logo-outer {
        position: relative;
        width: 300px;
        height: 270px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        transform-style: preserve-3d;
        transition: transform 0.15s ease-out;
      }
      .logo-ring {
        position: absolute;
        border-radius: 50%;
        border: 2px solid rgba(34,197,94,0.35);
        animation: ringPulse 3.2s ease-out infinite;
      }
      .logo-ring.r2 { animation-delay: 1.05s; border-color: rgba(59,130,246,0.30); }
      .logo-ring.r3 { animation-delay: 2.1s;  border-color: rgba(168,85,247,0.26); }
      @keyframes ringPulse {
        0%   { width: 130px; height: 130px; opacity: 0.9; }
        100% { width: 330px; height: 330px; opacity: 0; }
      }
      .logo-glow {
        position: absolute;
        width: 270px;
        height: 270px;
        border-radius: 50%;
        filter: blur(50px);
        opacity: 0.55;
        z-index: 0;
        animation: colorPulse 8s linear infinite;
      }
      .logo-3d {
        position: relative;
        z-index: 1;
        width: 210px;
        transform-style: preserve-3d;
        filter: drop-shadow(0 22px 26px rgba(0,0,0,0.45));
        animation: tilt3d 6s ease-in-out infinite;
      }
      .logo-word {
        position: relative;
        z-index: 1;
        margin-top: 12px;
        font-weight: 900;
        font-size: 30px;
        letter-spacing: 6px;
        background: linear-gradient(90deg, #16a34a, #111111 65%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        animation: fadeInWord 1.2s ease both;
      }
      @keyframes fadeInWord {
        from { opacity: 0; transform: translateY(6px); }
        to   { opacity: 1; transform: translateY(0); }
      }
      @keyframes tilt3d {
        0%   { transform: rotateY(-26deg) rotateX(10deg) translateY(0px); }
        50%  { transform: rotateY(26deg)  rotateX(4deg)  translateY(-14px); }
        100% { transform: rotateY(-26deg) rotateX(10deg) translateY(0px); }
      }
      @keyframes colorPulse {
        0%   { background: #22c55e; }
        33%  { background: #3b82f6; }
        66%  { background: #a855f7; }
        100% { background: #22c55e; }
      }
    </style>

    <div class="header-row">
      <div class="stat-stack">
        <div class="stat-card v-a">
          <div class="label">Headcount</div>
          <div class="value">__HEADCOUNT__</div>
          <div class="htip">Total drivers matching your current filters</div>
        </div>
        <div class="stat-card v-b">
          <div class="label">Active Drivers</div>
          <div class="value">__ACTIVE__</div>
          <div class="htip">Currently active, not terminated or suspended</div>
        </div>
      </div>

      <div class="logo-stage">
        <div class="logo-outer" id="logoOuter">
          <div class="logo-ring r1"></div>
          <div class="logo-ring r2"></div>
          <div class="logo-ring r3"></div>
          <div class="logo-glow"></div>
          <svg class="logo-3d" viewBox="0 0 220 260" xmlns="http://www.w3.org/2000/svg">
            <defs>
              <linearGradient id="greenGrad" x1="0" y1="0" x2="1" y2="1">
                <stop offset="0%" stop-color="#4ade80"/>
                <stop offset="100%" stop-color="#15803d"/>
              </linearGradient>
              <linearGradient id="darkGrad" x1="0" y1="0" x2="1" y2="1">
                <stop offset="0%" stop-color="#3f3f46"/>
                <stop offset="100%" stop-color="#0a0a0a"/>
              </linearGradient>
            </defs>
            __EXTRUSION_LAYERS__
            <path d="M25 25 L120 25 L120 62 L68 62 L68 98 L25 98 Z" fill="url(#greenGrad)"/>
            <path d="M62 25 L195 25 L195 62 L110 62 L152 98 L110 98 Z" fill="url(#darkGrad)"/>
            <rect x="78" y="98" width="28" height="142" fill="url(#greenGrad)"/>
            <rect x="114" y="98" width="28" height="142" fill="url(#darkGrad)"/>
            <path d="M90 228 L90 118 L106 98 L122 118 L122 228 Z" fill="#ffffff" opacity="0.92"/>
            <path d="M25 25 L120 25 L120 62 L68 62 L68 98 L25 98 Z" fill="none" stroke="rgba(255,255,255,0.5)" stroke-width="1.5"/>
            <path d="M62 25 L195 25 L195 62 L110 62 L152 98 L110 98 Z" fill="none" stroke="rgba(255,255,255,0.35)" stroke-width="1.5"/>
          </svg>
          <div class="logo-word">TAMKEEN</div>
        </div>
      </div>

      <div class="stat-stack">
        <div class="stat-card v-c">
          <div class="label">Orders (month)</div>
          <div class="value">__ORDERS__</div>
          <div class="htip">Total orders logged for the selected month</div>
        </div>
        <div class="stat-card v-d">
          <div class="label">Gross (month)</div>
          <div class="value">__GROSS__</div>
          <div class="htip">Total gross salary across all riders this month</div>
        </div>
      </div>
    </div>

    <script>
      (function() {
        try {
          const doc = window.parent.document;

          // Streamlit's light/dark toggle is 100% client-side -- it
          // never triggers a Streamlit rerun, so this iframe's HTML
          // (built server-side, see __STAT_TEXT__ above) never gets
          // regenerated when the user flips the toggle. That's why a
          // one-time color read used to go stale: it captured whatever
          // theme was active at mount time and then never updated.
          //
          // Fix: keep --stat-text in sync with the REAL current theme
          // for as long as this widget is on screen -- read it right
          // away, then watch for theme changes (MutationObserver on
          // the attributes Streamlit flips when toggling) and also
          // poll on an interval as a cheap, reliable fallback in case
          // the observer misses how a particular Streamlit version
          // signals the switch.
          function applyStatTextColor() {
            try {
              const parentColor = getComputedStyle(doc.body).color;
              if (parentColor) {
                document.documentElement.style.setProperty('--stat-text', parentColor);
              }
            } catch (e) {
              // Leave the CSS fallback color in place if this fails.
            }
          }
          applyStatTextColor();

          try {
            const observerTargets = [doc.documentElement, doc.body].filter(Boolean);
            const observer = new MutationObserver(applyStatTextColor);
            observerTargets.forEach((t) =>
              observer.observe(t, { attributes: true, attributeFilter: ["class", "style", "data-theme"] })
            );
          } catch (e) {
            // MutationObserver unavailable/blocked -- polling below still covers it.
          }

          // Belt-and-suspenders: re-check every second. Cheap (a single
          // getComputedStyle call) and guarantees this never stays
          // stuck on a stale color for long, regardless of how the
          // theme switch is implemented under the hood.
          setInterval(applyStatTextColor, 1000);

          const candidates = [
            doc.querySelector('section.main'),
            doc.querySelector('[data-testid="stAppViewContainer"]'),
            doc.querySelector('[data-testid="stMain"]'),
          ].filter(Boolean);
          const target = candidates[0];
          const outer = document.getElementById('logoOuter');
          if (target && outer) {
            const onScroll = () => {
              const top = target.scrollTop || 0;
              const extraTilt = Math.sin(top / 140) * 30;
              const lift = Math.min(top / 10, 20);
              outer.style.transform = `rotateY(${extraTilt}deg) translateY(-${lift}px)`;
            };
            target.addEventListener('scroll', onScroll);
          }
        } catch (e) {
          // Silently ignore -- the base tilt/glow/ring CSS animation still runs fine.
        }
      })();
    </script>
    """
    html = (
        html.replace("__EXTRUSION_LAYERS__", extrusion_layers)
        .replace("__HEADCOUNT__", str(headcount))
        .replace("__ACTIVE__", str(active))
        .replace("__ORDERS__", f"{orders_m:,}")
        .replace("__GROSS__", f"SAR {gross_m:,.0f}")
        .replace("__STAT_TEXT__", stat_text_color)
    )
    components.html(html, height=320)


# ==============================================================================
# DATABASE LAYER
# ==============================================================================


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def _add_column_if_missing(conn: sqlite3.Connection, table: str, column_def: str) -> None:
    col_name = column_def.split()[0]
    existing = [r[1] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]
    if col_name not in existing:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column_def}")


def init_db() -> None:
    conn = get_connection()
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS drivers (
            driver_id        TEXT PRIMARY KEY,
            driver_name      TEXT NOT NULL,
            phone            TEXT,
            supervisor_name  TEXT,
            status           TEXT NOT NULL DEFAULT 'Active',
            vehicle_type     TEXT NOT NULL DEFAULT 'Own Car',
            join_date        TEXT,
            termination_date TEXT
        );
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS monthly_logs (
            log_id             INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_id           TEXT NOT NULL,
            month_year          TEXT NOT NULL,
            total_orders         INTEGER NOT NULL DEFAULT 0,
            gross_salary          REAL NOT NULL DEFAULT 0,
            total_deductions      REAL NOT NULL DEFAULT 0,
            net_salary            REAL NOT NULL DEFAULT 0,
            validity_status       TEXT NOT NULL DEFAULT 'Valid',
            valid_days_in_month   INTEGER NOT NULL DEFAULT 0,
            FOREIGN KEY (driver_id) REFERENCES drivers(driver_id),
            UNIQUE (driver_id, month_year)
        );
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS access_log (
            email TEXT PRIMARY KEY,
            name TEXT,
            first_seen TEXT,
            last_seen TEXT,
            visit_count INTEGER NOT NULL DEFAULT 1
        );
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS daily_logs (
            driver_id   TEXT NOT NULL,
            month_year  TEXT NOT NULL,
            day         INTEGER NOT NULL,
            orders      INTEGER,
            validity    TEXT,
            attendance  TEXT,
            PRIMARY KEY (driver_id, month_year, day)
        );
        """
    )

    _add_column_if_missing(conn, "drivers", "iqama_number TEXT")
    _add_column_if_missing(conn, "drivers", "sponsor_name TEXT")
    _add_column_if_missing(conn, "monthly_logs", "cancelled_orders INTEGER NOT NULL DEFAULT 0")
    _add_column_if_missing(conn, "monthly_logs", "pending_salary REAL NOT NULL DEFAULT 0")
    _add_column_if_missing(conn, "monthly_logs", "in_roster INTEGER NOT NULL DEFAULT 0")

    conn.commit()
    conn.close()


def upsert_driver(conn: sqlite3.Connection, row: dict) -> None:
    """Insert a driver, or update it in place if the driver_id already
    exists. Vehicle type is only overwritten when THIS row actually
    carries a real value for it."""
    existing_vehicle = conn.execute(
        "SELECT vehicle_type FROM drivers WHERE driver_id = ?", (row["driver_id"],)
    ).fetchone()
    vehicle_type = row.get("vehicle_type") or (existing_vehicle[0] if existing_vehicle else None) or "Own Car"

    conn.execute(
        """
        INSERT INTO drivers
            (driver_id, driver_name, phone, supervisor_name, status,
             vehicle_type, join_date, termination_date, iqama_number, sponsor_name)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(driver_id) DO UPDATE SET
            driver_name       = excluded.driver_name,
            phone             = COALESCE(NULLIF(excluded.phone, ''), drivers.phone),
            supervisor_name   = COALESCE(NULLIF(excluded.supervisor_name, ''), drivers.supervisor_name),
            status            = excluded.status,
            vehicle_type      = excluded.vehicle_type,
            join_date         = COALESCE(NULLIF(excluded.join_date, ''), drivers.join_date),
            termination_date  = excluded.termination_date,
            iqama_number      = COALESCE(NULLIF(excluded.iqama_number, ''), drivers.iqama_number),
            sponsor_name      = COALESCE(NULLIF(excluded.sponsor_name, ''), drivers.sponsor_name)
        """,
        (
            row["driver_id"], row["driver_name"], row.get("phone"),
            row.get("supervisor_name"), row.get("status", "Active"),
            vehicle_type, row.get("join_date"),
            row.get("termination_date"), row.get("iqama_number"), row.get("sponsor_name"),
        ),
    )


def upsert_monthly_log(conn: sqlite3.Connection, row: dict) -> None:
    conn.execute(
        """
        INSERT INTO monthly_logs
            (driver_id, month_year, total_orders, gross_salary,
             total_deductions, net_salary, validity_status, valid_days_in_month,
             cancelled_orders, pending_salary, in_roster)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(driver_id, month_year) DO UPDATE SET
            total_orders        = excluded.total_orders,
            gross_salary         = excluded.gross_salary,
            total_deductions      = excluded.total_deductions,
            net_salary             = excluded.net_salary,
            validity_status         = excluded.validity_status,
            valid_days_in_month      = excluded.valid_days_in_month,
            cancelled_orders          = excluded.cancelled_orders,
            pending_salary             = excluded.pending_salary,
            in_roster                   = excluded.in_roster
        """,
        (
            row["driver_id"], row["month_year"], row["total_orders"],
            row["gross_salary"], row["total_deductions"], row["net_salary"],
            row["validity_status"], row["valid_days_in_month"],
            row.get("cancelled_orders", 0), row.get("pending_salary", 0.0),
            row.get("in_roster", 0),
        ),
    )


_MONTHLY_LOG_FIELDS = [
    "total_orders", "cancelled_orders", "gross_salary", "total_deductions",
    "pending_salary", "net_salary", "validity_status", "valid_days_in_month",
    "in_roster",
]
_MONTHLY_LOG_DEFAULTS = {
    "total_orders": 0, "cancelled_orders": 0, "gross_salary": 0.0,
    "total_deductions": 0.0, "pending_salary": 0.0, "net_salary": 0.0,
    "validity_status": "Valid", "valid_days_in_month": 0, "in_roster": 0,
}


def get_monthly_log(conn: sqlite3.Connection, driver_id: str, month_year: str):
    row = conn.execute(
        f"SELECT {', '.join(_MONTHLY_LOG_FIELDS)} FROM monthly_logs "
        f"WHERE driver_id = ? AND month_year = ?",
        (driver_id, month_year),
    ).fetchone()
    if row is None:
        return None
    return dict(zip(_MONTHLY_LOG_FIELDS, row))


def merge_monthly_log(conn: sqlite3.Connection, driver_id: str, month_year: str, updates: dict) -> None:
    existing = get_monthly_log(conn, driver_id, month_year) or dict(_MONTHLY_LOG_DEFAULTS)
    merged = dict(existing)
    for key, value in updates.items():
        if value is not None:
            merged[key] = value

    if updates.get("net_salary") is None and (
        updates.get("gross_salary") is not None or updates.get("total_deductions") is not None
    ):
        merged["net_salary"] = round(merged["gross_salary"] - merged["total_deductions"], 2)

    upsert_monthly_log(conn, {"driver_id": driver_id, "month_year": month_year, **merged})


# ==============================================================================
# SAMPLE DATA SEEDING
# ==============================================================================

FIRST_NAMES = [
    "Ahmed", "Bilal", "Usman", "Hassan", "Ali", "Zeeshan", "Faisal", "Kamran",
    "Imran", "Waqas", "Adeel", "Saad", "Hamza", "Talha", "Omar", "Junaid",
    "Rizwan", "Naveed", "Shahzad", "Asif", "Farhan", "Danish", "Salman",
    "Yasir", "Nabeel", "Arslan", "Sami", "Fahad", "Rehan", "Noman",
]
LAST_NAMES = [
    "Khan", "Malik", "Butt", "Sheikh", "Raza", "Iqbal", "Chaudhry", "Aslam",
    "Farooq", "Javed", "Qureshi", "Abbasi", "Rashid", "Hussain", "Baig",
]
SUPERVISORS = ["Bilal Ahmed", "Sara Khan", "Usman Tariq", "Ayesha Malik"]
SPONSORS = ["Tamkeen Est.", "Al Faisal Trading", "Rawabi Sponsorship", "Gulf Link Est."]


def _random_driver_pool(n: int) -> list:
    pool = set()
    while len(pool) < n:
        pool.add(f"{random.choice(FIRST_NAMES)} {random.choice(LAST_NAMES)}")
    return list(pool)


def seed_sample_data() -> None:
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("DELETE FROM monthly_logs;")
    cur.execute("DELETE FROM drivers;")

    today = datetime.today()
    months = [(today.replace(day=1) - timedelta(days=30 * i)).strftime("%Y-%m") for i in range(4)]
    months = sorted(set(months))

    names = _random_driver_pool(40)
    drivers = []

    for i, name in enumerate(names, start=1):
        driver_id = f"DRV-{1000 + i}"
        status = random.choices(DRIVER_STATUSES, weights=[70, 18, 12])[0]
        vehicle_type = random.choices(VEHICLE_TYPES, weights=[55, 45])[0]
        supervisor = random.choice(SUPERVISORS)
        sponsor = random.choice(SPONSORS)
        iqama_number = str(random.randint(2000000000, 2999999999))
        join_dt = today - timedelta(days=random.randint(60, 720))
        join_date = join_dt.strftime("%Y-%m-%d")

        termination_date = None
        if status == "Terminated":
            term_dt = join_dt + timedelta(days=random.randint(45, 600))
            if term_dt > today:
                term_dt = today - timedelta(days=random.randint(1, 40))
            termination_date = term_dt.strftime("%Y-%m-%d")

        drivers.append(
            {
                "driver_id": driver_id,
                "driver_name": name,
                "phone": f"03{random.randint(00, 99):02d}{random.randint(1000000, 9999999)}",
                "supervisor_name": supervisor,
                "status": status,
                "vehicle_type": vehicle_type,
                "join_date": join_date,
                "termination_date": termination_date,
                "iqama_number": iqama_number,
                "sponsor_name": sponsor,
            }
        )
        upsert_driver(conn, drivers[-1])

    for d in drivers:
        for month in months:
            if d["termination_date"] and month > d["termination_date"][:7]:
                continue
            if d["join_date"][:7] > month:
                continue

            orders = random.randint(40, 480)
            cancelled_orders = random.randint(0, max(1, orders // 12))
            rate_per_order = random.uniform(45, 65)
            base_pay = random.uniform(8000, 15000)
            gross_salary = round(base_pay + orders * rate_per_order, 2)
            deductions = round(random.uniform(0, 4500), 2)
            pending_salary = round(random.choice([0, 0, 0, random.uniform(200, 3000)]), 2)
            net_salary = round(gross_salary - deductions, 2)
            valid_days = random.randint(18, 30)
            validity_status = random.choices(VALIDITY_STATUSES, weights=[80, 20])[0]

            upsert_monthly_log(
                conn,
                {
                    "driver_id": d["driver_id"],
                    "month_year": month,
                    "total_orders": orders,
                    "gross_salary": gross_salary,
                    "total_deductions": deductions,
                    "net_salary": net_salary,
                    "validity_status": validity_status,
                    "valid_days_in_month": valid_days,
                    "cancelled_orders": cancelled_orders,
                    "pending_salary": pending_salary,
                },
            )

    conn.commit()
    conn.close()


def clear_all_data() -> None:
    conn = get_connection()
    conn.execute("DELETE FROM monthly_logs;")
    conn.execute("DELETE FROM drivers;")
    conn.commit()
    conn.close()


# ==============================================================================
# VALIDITY TARGETS  -- "Valid" / "Invalid" here means the rider hit BOTH a
# minimum monthly order count AND a minimum days-worked count -- a whole-
# month performance target, not the raw day-by-day VALID/INVALID marks a
# Validity Report sheet might carry (a rider who worked every day but fell
# short on total orders should still show Invalid; the reverse too).
# ==============================================================================


def _load_validity_targets() -> dict:
    if not os.path.exists(VALIDITY_TARGETS_PATH):
        return {"min_orders": DEFAULT_MIN_ORDERS_FOR_VALID, "min_days": DEFAULT_MIN_DAYS_FOR_VALID}
    try:
        with open(VALIDITY_TARGETS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        return {
            "min_orders": int(data.get("min_orders", DEFAULT_MIN_ORDERS_FOR_VALID)),
            "min_days": int(data.get("min_days", DEFAULT_MIN_DAYS_FOR_VALID)),
        }
    except Exception:  # noqa: BLE001
        return {"min_orders": DEFAULT_MIN_ORDERS_FOR_VALID, "min_days": DEFAULT_MIN_DAYS_FOR_VALID}


def _save_validity_targets(min_orders: int, min_days: int) -> None:
    with open(VALIDITY_TARGETS_PATH, "w", encoding="utf-8") as f:
        json.dump({"min_orders": int(min_orders), "min_days": int(min_days)}, f)


def compute_performance_validity(total_orders, days_worked, min_orders: int, min_days: int) -> str:
    """Whole-month performance classification: Valid only when BOTH the
    order count and the days-worked count meet their targets for the
    month -- 500 orders in only 20 days is still Invalid; 26 days
    worked with only 200 orders is still Invalid too."""
    orders = total_orders or 0
    days = days_worked or 0
    if orders < min_orders or days < min_days:
        return "Invalid"
    return "Valid"


# ==============================================================================
# BACKUP / RESTORE  -- Streamlit Community Cloud wipes local files (the
# SQLite DB, hq_owner.json, hq_gsheet.json) on every redeploy, since
# they're server-local and never part of the git repo. Backing up
# before a code push and restoring right after is the practical
# workaround until/unless the app is moved to persistent cloud storage.
# ==============================================================================

def build_backup_zip() -> bytes:
    """Package every local file that would otherwise be lost on a
    redeploy (the database, the Admin login, the Google Sheet
    connection settings) into one downloadable .zip."""
    backup_files = [DB_PATH, OWNER_CONFIG_PATH, "hq_gsheet.json", VALIDITY_TARGETS_PATH]
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in backup_files:
            if os.path.exists(path):
                zf.write(path, arcname=os.path.basename(path))
    return buffer.getvalue()


def restore_backup_zip(uploaded_file) -> list:
    """Extracts a backup .zip built by build_backup_zip() and writes
    its files back to disk, overwriting whatever's currently there.
    Returns the list of filenames actually restored."""
    backup_files = [DB_PATH, OWNER_CONFIG_PATH, "hq_gsheet.json", VALIDITY_TARGETS_PATH]
    restored = []
    with zipfile.ZipFile(uploaded_file) as zf:
        names = set(zf.namelist())
        for path in backup_files:
            fname = os.path.basename(path)
            if fname in names:
                with zf.open(fname) as src, open(path, "wb") as dst:
                    dst.write(src.read())
                restored.append(fname)
    return restored


def delete_month_data(month_year: str) -> dict:
    """Wipes just ONE month's data -- monthly_logs rows and the
    company-level salary_summary row for that month -- while leaving
    the drivers table (roster/profiles) untouched. Use this before
    re-uploading a corrected file for that month, so the re-upload
    starts from a clean slate instead of merging on top of whatever
    (possibly wrong) numbers were there before."""
    conn = get_connection()
    _ensure_salary_summary_table(conn)
    logs_deleted = conn.execute(
        "SELECT COUNT(*) FROM monthly_logs WHERE month_year = ?", (month_year,)
    ).fetchone()[0]
    conn.execute("DELETE FROM monthly_logs WHERE month_year = ?", (month_year,))
    conn.execute("DELETE FROM salary_summary WHERE month_year = ?", (month_year,))
    conn.commit()
    conn.close()
    return {"logs_deleted": logs_deleted}


# ==============================================================================
# DATA ACCESS HELPERS (return pandas DataFrames for the UI layer)
# ==============================================================================


UNASSIGNED = "Unassigned"


def load_drivers() -> pd.DataFrame:
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM drivers", conn)
    conn.close()
    df["supervisor_name"] = df["supervisor_name"].fillna(UNASSIGNED).replace("", UNASSIGNED)
    return df


def load_monthly_logs() -> pd.DataFrame:
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM monthly_logs", conn)
    conn.close()
    return df


def load_merged() -> pd.DataFrame:
    conn = get_connection()
    df = pd.read_sql_query(
        """
        SELECT d.driver_id, d.driver_name, d.phone, d.supervisor_name,
               d.status, d.vehicle_type, d.join_date, d.termination_date,
               d.iqama_number, d.sponsor_name,
               m.log_id, m.month_year, m.total_orders, m.gross_salary,
               m.total_deductions, m.net_salary, m.validity_status,
               m.valid_days_in_month, m.cancelled_orders, m.pending_salary,
               m.in_roster
        FROM drivers d
        LEFT JOIN monthly_logs m ON d.driver_id = m.driver_id
        """,
        conn,
    )
    conn.close()
    df["supervisor_name"] = df["supervisor_name"].fillna(UNASSIGNED).replace("", UNASSIGNED)
    df["in_roster"] = df["in_roster"].fillna(0).astype(int)
    return df


def distinct_months() -> list:
    conn = get_connection()
    rows = conn.execute(
        "SELECT DISTINCT month_year FROM monthly_logs ORDER BY month_year DESC"
    ).fetchall()
    conn.close()
    return [r[0] for r in rows]


def month_display(ym: str) -> str:
    if not ym:
        return "N/A"
    try:
        return datetime.strptime(ym, "%Y-%m").strftime("%B %Y")
    except ValueError:
        return ym


def distinct_supervisors() -> list:
    conn = get_connection()
    rows = conn.execute(
        "SELECT DISTINCT COALESCE(NULLIF(supervisor_name, ''), ?) AS s FROM drivers ORDER BY s",
        (UNASSIGNED,),
    ).fetchall()
    conn.close()
    return [r[0] for r in rows]


# ==============================================================================
# SIDEBAR: DB CONTROLS + GLOBAL FILTERS
# ==============================================================================


def render_sidebar():
    st.sidebar.title("\U0001F6F5 Logistics Control Panel")
    st.sidebar.caption(f"Connected to `{DB_PATH}`")

    role = st.session_state.get("auth_role")
    if role == "admin":
        st.sidebar.success(f"\U0001F511 Signed in as Admin\n\n{st.session_state.get('auth_email', '')}")
    else:
        name = st.session_state.get("auth_name", "")
        email = st.session_state.get("auth_email", "")
        st.sidebar.info(f"\U0001F441\uFE0F Signed in as Viewer (read-only)\n\n{name} \u2022 {email}")

    if st.sidebar.button("Sign out", use_container_width=True):
        st.session_state.auth_role = None
        st.session_state.auth_email = None
        st.session_state.auth_name = None
        st.rerun()

    if is_admin():
        render_hq_access_panel()

        st.sidebar.markdown("### \U0001F331 Demo Data")
        if st.sidebar.button("Seed Sample Data", use_container_width=True):
            seed_sample_data()
            st.sidebar.success("Sample data seeded successfully!")
            st.rerun()

        with st.sidebar.expander("\u26A0\uFE0F Danger Zone"):
            month_options = distinct_months()
            if month_options:
                st.markdown("##### \U0001F5D3\uFE0F Delete One Month's Data")
                st.caption(
                    "Wipes only the selected month's orders/validity/payroll "
                    "figures -- drivers/roster profiles are kept. Use this "
                    "before re-uploading a corrected file for that month, so "
                    "the new upload doesn't merge on top of old numbers."
                )
                month_to_delete = st.selectbox(
                    "Month to delete", month_options, format_func=month_display,
                    key="danger_zone_month_pick",
                )
                confirm_month_delete = st.checkbox(
                    f"I understand this permanently deletes all "
                    f"{month_display(month_to_delete)} data",
                    key="danger_zone_month_confirm",
                )
                if st.button(
                    f"\U0001F5D1\uFE0F Delete {month_display(month_to_delete)} Data",
                    use_container_width=True,
                    disabled=not confirm_month_delete,
                ):
                    result = delete_month_data(month_to_delete)
                    st.success(
                        f"Deleted {result['logs_deleted']} log row(s) for "
                        f"{month_display(month_to_delete)}. You can now "
                        f"re-upload a corrected file for that month."
                    )
                    st.rerun()
                st.markdown("---")

            st.markdown("##### \U0001F9E8 Clear Everything")
            st.caption("Permanently wipes all drivers and monthly logs.")
            if st.button("Clear All Data", use_container_width=True):
                clear_all_data()
                st.success("Database cleared.")
                st.rerun()

        with st.sidebar.expander("\U0001F4BE Backup & Restore", expanded=False):
            st.caption(
                "\u26A0\uFE0F **Do this before every code update.** Streamlit Cloud wipes "
                "this app's saved data (database, Admin login, Google Sheet connection) "
                "every time you push new code and it redeploys -- those files live only "
                "on the server, never in GitHub. Download a backup first, push your code, "
                "then restore the backup right after."
            )
            st.download_button(
                "\u2B07\uFE0F Download Backup (.zip)",
                data=build_backup_zip(),
                file_name=f"logistics_backup_{datetime.today().strftime('%Y-%m-%d')}.zip",
                mime="application/zip",
                use_container_width=True,
            )
            st.markdown("---")
            restore_file = st.file_uploader(
                "Restore from a backup .zip", type=["zip"], key="restore_backup_uploader",
            )
            if restore_file is not None:
                st.warning(
                    "\u26A0\uFE0F This will overwrite the current database, Admin login, "
                    "and Google Sheet connection with whatever's in this backup file."
                )
                if st.button("\u267B\uFE0F Restore This Backup", use_container_width=True, type="primary"):
                    try:
                        restored = restore_backup_zip(restore_file)
                    except Exception as exc:  # noqa: BLE001
                        st.error(f"Couldn't restore this backup: {exc}")
                    else:
                        if restored:
                            st.success(f"Restored: {', '.join(restored)}. Reloading...")
                            st.rerun()
                        else:
                            st.error("This .zip didn't contain any recognizable backup files.")

    st.sidebar.markdown("---")
    st.sidebar.markdown("### \U0001F50E Global Filters")

    if is_admin():
        with st.sidebar.expander("\U0001F3AF Validity Targets", expanded=False):
            st.caption(
                "A rider counts as **Valid** for the month only when BOTH targets "
                "are met -- falling short on either one makes them Invalid."
            )
            targets = _load_validity_targets()
            new_min_orders = st.number_input(
                "Minimum orders this month", min_value=0, value=targets["min_orders"], step=10,
                key="validity_target_orders",
            )
            new_min_days = st.number_input(
                "Minimum days worked this month", min_value=0, max_value=31, value=targets["min_days"], step=1,
                key="validity_target_days",
            )
            if st.button("\U0001F4BE Save Targets", use_container_width=True):
                _save_validity_targets(new_min_orders, new_min_days)
                st.success("Saved.")
                st.rerun()

    months = distinct_months()
    supervisors = distinct_supervisors()

    if not months:
        st.sidebar.info("No monthly data yet. Seed sample data or upload a file.")
        selected_month = None
    else:
        selected_month = st.sidebar.selectbox(
            "Month", months, index=0, format_func=month_display
        )

    selected_supervisors = st.sidebar.multiselect(
        "Supervisor", supervisors, default=supervisors
    )
    selected_vehicle_types = st.sidebar.multiselect(
        "Vehicle Type", VEHICLE_TYPES, default=VEHICLE_TYPES
    )
    selected_validity = st.sidebar.multiselect(
        "Validity Status", VALIDITY_STATUSES, default=VALIDITY_STATUSES
    )

    return {
        "month": selected_month,
        "supervisors": selected_supervisors or supervisors,
        "vehicle_types": selected_vehicle_types or VEHICLE_TYPES,
        "validity": selected_validity or VALIDITY_STATUSES,
    }


def apply_filters(df: pd.DataFrame, filters: dict, month_scoped: bool = True) -> pd.DataFrame:
    out = df.copy()
    if filters["supervisors"]:
        out = out[out["supervisor_name"].isin(filters["supervisors"])]
    if filters["vehicle_types"]:
        out = out[out["vehicle_type"].isin(filters["vehicle_types"])]
    if month_scoped and filters["month"]:
        out = out[out["month_year"] == filters["month"]]
    if filters["validity"]:
        out = out[out["validity_status"].isin(filters["validity"]) | out["validity_status"].isna()]
    return out


# ==============================================================================
# TAB 1: OPERATIONS & WORKFORCE KPI DASHBOARD
# ==============================================================================


def render_dashboard(filters: dict):
    st.subheader("\U0001F4CA Operations & Workforce KPI Dashboard")

    merged = load_merged()
    if merged.empty:
        st.info("No data available yet. Use **Seed Sample Data** in the sidebar or upload a file.")
        return

    all_months = distinct_months()
    if all_months:
        oldest, newest = min(all_months), max(all_months)
        span_years = (
            (int(newest[:4]) - int(oldest[:4])) + (int(newest[5:7]) - int(oldest[5:7])) / 12
        )
        st.caption(
            f"\U0001F4DA Tracking **{len(all_months)} month(s)** of history "
            f"({month_display(oldest)} \u2192 {month_display(newest)}, "
            f"~{span_years:.1f} year(s) of data on file)."
        )

    roster = merged.drop_duplicates(subset="driver_id")
    roster_filtered = roster[
        roster["supervisor_name"].isin(filters["supervisors"])
        & roster["vehicle_type"].isin(filters["vehicle_types"])
    ]

    month_df = merged[merged["month_year"] == filters["month"]] if filters["month"] else merged.iloc[0:0]
    month_df = month_df[
        month_df["supervisor_name"].isin(filters["supervisors"])
        & month_df["vehicle_type"].isin(filters["vehicle_types"])
    ]

    if filters["month"]:
        # Scoped to THIS month's own records, and further scoped to
        # in_roster==1 -- riders actually listed in this month's
        # roster/active-rider sheet. An orders/validity/attendance sheet
        # can reference IDs the roster sheet never had (typos, a
        # different numbering scheme); those get their data recorded
        # too, but they must not inflate headcount beyond what the
        # roster sheet itself says.
        month_roster = month_df.drop_duplicates(subset="driver_id")
        roster_only = month_roster[month_roster["in_roster"] == 1]
        total_headcount = roster_only["driver_id"].nunique()
        # "Active" requires BOTH an employment status that isn't
        # Terminated/Suspended AND real activity this month (days
        # worked or orders completed) -- see _active_mask(). A roster's
        # own status column is often blank/unreliable on its own, so a
        # rider who's simply listed with no attendance/orders data at
        # all no longer counts as Active by default.
        active_drivers = int(roster_only[_active_mask(roster_only)]["driver_id"].nunique())
        suspended_drivers = int((roster_only["status"] == "Suspended").sum())
        terminated_this_month = int((roster_only["status"] == "Terminated").sum())
    else:
        total_headcount = len(roster_filtered)
        active_drivers = (roster_filtered["status"] == "Active").sum()
        suspended_drivers = int((roster_filtered["status"] == "Suspended").sum())
        terminated_this_month = 0

    # Vehicle type only ever comes from the roster sheet itself, so these
    # are scoped to in_roster==1 too -- otherwise a placeholder driver
    # created from an unmatched orders/validity row (which has no real
    # vehicle info) silently defaults to "Own Car" and skews the count.
    roster_only_for_vehicle = month_df[month_df["in_roster"] == 1].drop_duplicates(subset="driver_id")
    company_cars = (roster_only_for_vehicle["vehicle_type"] == "Company Car").sum()
    own_cars = (roster_only_for_vehicle["vehicle_type"] == "Own Car").sum()

    # "Valid" / "Invalid" here is a whole-month PERFORMANCE target, not
    # the raw day-by-day marks a Validity Report sheet might carry: a
    # rider needs to clear BOTH a minimum order count and a minimum
    # days-worked count for the month to count as Valid. See
    # compute_performance_validity() -- thresholds are Admin-adjustable
    # under Global Filters -> Validity Targets in the sidebar.
    targets = _load_validity_targets()
    month_df = month_df.copy()
    month_df["performance_validity"] = month_df.apply(
        lambda r: compute_performance_validity(
            r["total_orders"], r["valid_days_in_month"], targets["min_orders"], targets["min_days"]
        ),
        axis=1,
    )
    valid_drivers = (month_df["performance_validity"] == "Valid").sum()
    invalid_drivers = (month_df["performance_validity"] == "Invalid").sum()

    st.markdown(f"**Selected Month:** `{month_display(filters['month'])}`")
    st.caption(
        f"\U0001F446 Tap any card below to see exactly which riders make up that number. "
        f"Valid/Invalid targets: \u2265{targets['min_orders']} orders and \u2265{targets['min_days']} "
        f"days worked this month (change under Global Filters \u2192 Validity Targets)."
    )

    active_only = roster_only[_active_mask(roster_only)] if not roster_only.empty else roster_only
    terminated_only = roster_only[roster_only["status"] == "Terminated"]
    suspended_only = roster_only[roster_only["status"] == "Suspended"]
    company_car_only = roster_only_for_vehicle[roster_only_for_vehicle["vehicle_type"] == "Company Car"]
    own_car_only = roster_only_for_vehicle[roster_only_for_vehicle["vehicle_type"] == "Own Car"]
    valid_only = month_df[month_df["performance_validity"] == "Valid"]
    invalid_only = month_df[month_df["performance_validity"] == "Invalid"]

    roster_cols = ["driver_id", "driver_name", "supervisor_name", "status", "vehicle_type"]
    validity_cols = ["driver_id", "driver_name", "total_orders", "valid_days_in_month", "performance_validity"]

    drill_defs = {
        "row1_0": ("Total Headcount", roster_only[roster_cols].sort_values("driver_name") if not roster_only.empty else roster_only, None),
        "row1_1": ("Active Drivers", active_only[roster_cols].sort_values("driver_name") if not active_only.empty else active_only, None),
        "row1_2": ("Terminated (this month)", terminated_only[roster_cols + ["termination_date"]].sort_values("driver_name") if not terminated_only.empty else terminated_only, None),
        "row1_3": ("Suspended Drivers", suspended_only[roster_cols].sort_values("driver_name") if not suspended_only.empty else suspended_only, None),
        "row2_0": ("Company Cars (month)", company_car_only[["driver_id", "driver_name", "supervisor_name"]].sort_values("driver_name") if not company_car_only.empty else company_car_only, None),
        "row2_1": ("Own Cars (month)", own_car_only[["driver_id", "driver_name", "supervisor_name"]].sort_values("driver_name") if not own_car_only.empty else own_car_only, None),
        "row2_2": ("Valid Drivers (month)", valid_only[validity_cols].sort_values("driver_name") if not valid_only.empty else valid_only, None),
        "row2_3": ("Invalid Drivers (month)", invalid_only[validity_cols].sort_values("driver_name") if not invalid_only.empty else invalid_only, "These need supervisor follow-up."),
        "row3_0": ("Total Orders (month)", month_df[["driver_id", "driver_name", "total_orders"]].sort_values("total_orders", ascending=False), None),
        "row3_1": ("Cancelled Orders (month)", month_df[["driver_id", "driver_name", "cancelled_orders"]].sort_values("cancelled_orders", ascending=False), None),
        "row3_2": ("Gross Salary (month)", month_df[["driver_id", "driver_name", "gross_salary"]].sort_values("gross_salary", ascending=False), None),
        "row3_3": ("Pending Salary (month)", month_df[["driver_id", "driver_name", "pending_salary"]].sort_values("pending_salary", ascending=False), None),
    }

    render_clickable_stat_row([
        {"icon": "\U0001F465", "label": "Total Headcount", "value": total_headcount,
         "tip": "Riders listed in this month's roster/active-rider sheet", "variant": "a"},
        {"icon": "\U0001F7E2", "label": "Active Drivers", "value": int(active_drivers),
         "tip": "Not terminated/suspended, and shows orders or days worked this month", "variant": "a"},
        {"icon": "\U0001F6D1", "label": "Terminated (this month)", "value": terminated_this_month,
         "tip": "Marked Terminated in this month's uploaded roster", "variant": "c"},
        {"icon": "\u23F8\uFE0F", "label": "Suspended Drivers", "value": suspended_drivers,
         "tip": "Temporarily suspended / on leave, not terminated", "variant": "d"},
    ], row_key="row1")
    if (st.session_state.get("active_drill") or "").startswith("row1_"):
        render_drilldown_panel(drill_defs)

    render_clickable_stat_row([
        {"icon": "\U0001F697", "label": "Company Cars (month)", "value": int(company_cars),
         "tip": "Riders using a company-provided vehicle", "variant": "b"},
        {"icon": "\U0001F699", "label": "Own Cars (month)", "value": int(own_cars),
         "tip": "Riders using their own vehicle", "variant": "b"},
        {"icon": "\u2705", "label": "Valid Drivers (month)", "value": int(valid_drivers),
         "tip": f"Hit \u2265{targets['min_orders']} orders AND \u2265{targets['min_days']} days worked", "variant": "a"},
        {"icon": "\u274C", "label": "Invalid Drivers (month)", "value": int(invalid_drivers),
         "tip": f"Missed the \u2265{targets['min_orders']}-order or \u2265{targets['min_days']}-day target", "variant": "c"},
    ], row_key="row2")
    if (st.session_state.get("active_drill") or "").startswith("row2_"):
        render_drilldown_panel(drill_defs)

    render_clickable_stat_row([
        {"icon": "\U0001F4E6", "label": "Total Orders (month)", "value": f"{int(month_df['total_orders'].sum()):,}",
         "tip": "Sum of all completed orders this month", "variant": "a"},
        {"icon": "\u274C", "label": "Cancelled Orders (month)", "value": f"{int(month_df['cancelled_orders'].sum()):,}",
         "tip": "Orders cancelled across the whole team", "variant": "c"},
        {"icon": "\U0001F4B5", "label": "Gross Salary (month)", "value": f"SAR {month_df['gross_salary'].sum():,.0f}",
         "tip": "Total pay before deductions", "variant": "b"},
        {"icon": "\u23F3", "label": "Pending Salary (month)", "value": f"SAR {month_df['pending_salary'].sum():,.0f}",
         "tip": "Amount still owed, not yet paid out", "variant": "d"},
    ], row_key="row3")
    if (st.session_state.get("active_drill") or "").startswith("row3_"):
        render_drilldown_panel(drill_defs)

    st.markdown("---")
    st.markdown("#### Filtered Roster Detail")
    st.caption(
        "Every driver in your roster appears here regardless of whether they have "
        "payroll data for the selected month -- rows with no log for this month show "
        "blank figures instead of disappearing."
    )
    display_cols = [
        "driver_id", "driver_name", "supervisor_name", "status", "vehicle_type",
        "sponsor_name", "iqama_number", "join_date", "termination_date",
        "month_year", "total_orders", "cancelled_orders",
        "validity_status", "valid_days_in_month",
    ]

    roster_scope = merged[
        merged["supervisor_name"].isin(filters["supervisors"])
        & merged["vehicle_type"].isin(filters["vehicle_types"])
    ]

    if filters["month"]:
        this_month_rows = roster_scope[roster_scope["month_year"] == filters["month"]]
        covered_ids = set(this_month_rows["driver_id"])
        no_log_rows = roster_scope[~roster_scope["driver_id"].isin(covered_ids)].drop_duplicates(subset="driver_id").copy()
        for c in ["month_year", "validity_status"]:
            no_log_rows[c] = None
        for c in ["total_orders", "cancelled_orders", "valid_days_in_month"]:
            no_log_rows[c] = 0
        roster_view = pd.concat([this_month_rows, no_log_rows], ignore_index=True)
    else:
        roster_view = roster_scope.drop_duplicates(subset="driver_id")

    if filters["validity"]:
        roster_view = roster_view[
            roster_view["validity_status"].isin(filters["validity"]) | roster_view["validity_status"].isna()
        ]

    st.dataframe(
        roster_view[display_cols].sort_values("driver_name"),
        use_container_width=True,
        hide_index=True,
    )


# ==============================================================================
# TAB 2: FINANCIAL & PAYROLL SUMMARY
# ==============================================================================


def render_financials(filters: dict):
    st.subheader("\U0001F4B0 Financial & Payroll Summary")

    merged = load_merged()
    if merged.empty or not filters["month"]:
        st.info("No payroll data available yet. Seed sample data or upload a file.")
        return

    month_df = apply_filters(merged, filters, month_scoped=True)
    month_df = month_df[month_df["log_id"].notna()]

    total_gross = float(month_df["gross_salary"].sum())
    total_deductions = float(month_df["total_deductions"].sum())
    total_net = float(month_df["net_salary"].sum())
    total_pending = float(month_df["pending_salary"].sum())

    conn = get_connection()
    _ensure_salary_summary_table(conn)
    summary_row = conn.execute(
        "SELECT total_payable, tax_amount, invoice_amount FROM salary_summary WHERE month_year = ?",
        (filters["month"],),
    ).fetchone()
    conn.close()

    if summary_row and (summary_row[0] is not None or summary_row[1] is not None):
        company_total, tax_amount, invoice_amount = summary_row
        stat_cards([
            {"icon": "\U0001F4B0", "label": "Total Money Received", "value": f"SAR {(company_total or 0):,.2f}",
             "tip": "Company-level total payable amount for this billing cycle", "variant": "a"},
            {"icon": "\U0001F9FE", "label": "Tax Amount", "value": f"SAR {(tax_amount or 0):,.2f}",
             "tip": "Tax amount for this billing cycle, from the salary summary sheet", "variant": "c"},
            {"icon": "\U0001F4C4", "label": "Invoice Amount", "value": f"SAR {(invoice_amount or 0):,.2f}",
             "tip": "Invoiced amount for this billing cycle", "variant": "b"},
        ])
    else:
        st.caption(
            "\u2139\uFE0F No company-level salary summary (Total Money Received / "
            "Tax Amount) found for this month yet -- upload it from "
            "**Upload Monthly Data \u2192 Salary Data**."
        )

    stat_cards([
        {"icon": "\U0001F4B5", "label": "Gross Salary", "value": f"SAR {total_gross:,.0f}",
         "tip": "Total pay before deductions", "variant": "b"},
        {"icon": "\u2796", "label": "Deductions", "value": f"SAR {total_deductions:,.0f}",
         "tip": "Total amounts deducted this month", "variant": "c"},
        {"icon": "\u2705", "label": "Net Salary", "value": f"SAR {total_net:,.0f}",
         "tip": "Gross minus deductions", "variant": "a"},
        {"icon": "\u23F3", "label": "Pending Salary", "value": f"SAR {total_pending:,.0f}",
         "tip": "Still owed, not yet paid out", "variant": "d"},
    ])

    st.markdown("---")
    st.markdown("#### Individual Rider Breakdown")

    breakdown = month_df[
        [
            "driver_name", "iqama_number", "sponsor_name", "join_date",
            "total_orders", "cancelled_orders", "valid_days_in_month",
            "gross_salary", "total_deductions", "pending_salary", "net_salary",
            "vehicle_type", "status",
        ]
    ].rename(
        columns={
            "driver_name": "Rider Name",
            "iqama_number": "IQAMA",
            "sponsor_name": "Sponsor",
            "join_date": "Join Date",
            "total_orders": "Total Orders",
            "cancelled_orders": "Cancelled",
            "valid_days_in_month": "Days Worked",
            "gross_salary": "Gross Salary",
            "total_deductions": "Deductions",
            "pending_salary": "Pending",
            "net_salary": "Net Salary",
            "vehicle_type": "Vehicle Type",
            "status": "Status",
        }
    ).sort_values("Net Salary", ascending=False)

    st.dataframe(breakdown, use_container_width=True, hide_index=True)

    csv_bytes = breakdown.to_csv(index=False).encode("utf-8")
    st.download_button(
        "\u2B07\uFE0F Download Breakdown as CSV",
        data=csv_bytes,
        file_name=f"payroll_breakdown_{filters['month']}.csv",
        mime="text/csv",
    )


# ==============================================================================
# TAB 3: BULK EXCEL / CSV UPLOAD  (flexible column-mapping engine)
# ==============================================================================

NONE_OPTION = "\u2014 None / not in this file \u2014"

FIELD_ALIASES = {
    "driver_id": ["driver id", "driverid", "courier id", "rider id", "emp id", "employee id"],
    "driver_name": ["driver name", "name", "rider name", "courier name", "courier nmae", "full name"],
    "first_name": ["first name", "courier first name", "fname"],
    "last_name": ["last name", "courier last name", "lname", "surname"],
    "phone": ["phone", "mobile", "mobile number", "contact", "contact number"],
    "supervisor_name": ["supervisor", "supervisor name", "manager", "team lead"],
    "sponsor_name": ["sponsor", "sponser", "sponsor name", "sponsorship"],
    "iqama_number": ["iqama number", "iqama", "iqama no", "national id"],
    "vehicle_type": ["vehicle type", "vehicle", "car type"],
    "status": ["driver status", "employment status", "rider status"],
    "join_date": ["joining date", "join date", "start date", "doj", "date joined"],
    "termination_date": ["ending date", "termination date", "end date", "leaving date", "date left"],
    "month_year": ["month year", "month", "period", "payroll month"],
    "total_orders": ["total orders", "total order", "orders", "order count", "trips", "deliveries"],
    "cancelled_orders": ["cancelled orders", "cancel orders", "cancelled", "canceled orders"],
    "gross_salary": ["gross salary", "gross pay", "salary", "gross"],
    "total_deductions": ["deductions", "total deductions", "deduction"],
    "pending_salary": ["pending salary", "salary pending", "pending amount", "outstanding salary"],
    "net_salary": ["net salary", "net pay", "net"],
    "validity_status": ["validity", "validity status"],
    "valid_days_in_month": ["valid days", "days valid", "valid days in month", "attendance"],
}

FIELD_LABELS = {
    "driver_id": "Driver ID  (optional -- auto-generated from name if not mapped)",
    "driver_name": "Driver Name  (skip if using First + Last name below)",
    "first_name": "First Name  (optional, combine with Last Name)",
    "last_name": "Last Name  (optional, combine with First Name)",
    "phone": "Phone",
    "supervisor_name": "Supervisor Name",
    "sponsor_name": "Sponsor",
    "iqama_number": "IQAMA Number",
    "vehicle_type": "Vehicle Type  (Company Car / Own Car)",
    "status": "Driver Status  (Active/Terminated/Suspended -- auto-detected from Ending Date if skipped)",
    "join_date": "Join Date",
    "termination_date": "Termination / Ending Date",
    "month_year": "Payroll Month (e.g. 2026-06)",
    "total_orders": "Total Orders",
    "cancelled_orders": "Cancelled Orders",
    "gross_salary": "Gross Salary",
    "total_deductions": "Total Deductions",
    "pending_salary": "Pending Salary",
    "net_salary": "Net Salary  (auto-calculated as Gross minus Deductions if skipped)",
    "validity_status": "Validity Status  (Valid/Invalid)",
    "valid_days_in_month": "Days Worked / Attendance",
}

PAYROLL_FIELDS = [
    "total_orders", "cancelled_orders", "gross_salary", "total_deductions",
    "pending_salary", "net_salary", "validity_status", "valid_days_in_month",
]


def _normalize_header(h: str) -> str:
    return str(h).strip().lower().replace("_", " ")


_GENERIC_HEADER_STOPWORDS = {"total", "amount", "sum", "value", "grand total", "date"}


def _guess_column(columns: list, aliases: list) -> str:
    normed = {c: _normalize_header(c) for c in columns}
    for col, norm in normed.items():
        if norm in aliases:
            return col
    for col, norm in normed.items():
        if norm in _GENERIC_HEADER_STOPWORDS:
            continue
        for alias in aliases:
            if alias in norm or norm in alias:
                return col
    return NONE_OPTION


def _build_template_csv() -> bytes:
    template = pd.DataFrame(
        [
            {
                "driver_id": "DRV-1001",
                "driver_name": "Ali Raza",
                "phone": "03001234567",
                "supervisor_name": "Bilal Ahmed",
                "sponsor_name": "Tamkeen Est.",
                "iqama_number": "2123456789",
                "status": "Active",
                "vehicle_type": "Company Car",
                "join_date": "2025-01-15",
                "termination_date": "",
                "month_year": "2026-06",
                "total_orders": 320,
                "cancelled_orders": 6,
                "gross_salary": 45000,
                "total_deductions": 1500,
                "pending_salary": 0,
                "net_salary": 43500,
                "validity_status": "Valid",
                "valid_days_in_month": 28,
            }
        ]
    )
    return template.to_csv(index=False).encode("utf-8")


def _clean_number_value(v, as_int: bool = False):
    if pd.isna(v):
        return 0 if as_int else 0.0
    if isinstance(v, (int, float)):
        return int(round(v)) if as_int else float(v)
    s = str(v).strip()
    if not s or s.lower() in ("nan", "n/a", "na", "-", "--", "none"):
        return 0 if as_int else 0.0
    negative = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    for ch in ["Rs", "SR", "SAR", "PKR", "$", ",", "%", " "]:
        s = s.replace(ch, "")
    try:
        num = float(s)
    except ValueError:
        return 0 if as_int else 0.0
    if negative:
        num = -num
    return int(round(num)) if as_int else num


def _clean_id_value(v) -> str:
    if pd.isna(v):
        return ""
    if isinstance(v, float):
        s = str(int(v)) if v.is_integer() else str(v)
    else:
        s = str(v).strip()
    return s.strip().upper()


def _clean_date_value(v):
    if pd.isna(v):
        return None
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    if not s or s.lower() == "nan" or s.lower() == "nat":
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return s


_YYYY_MM_DD_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def _is_real_parsed_date(s) -> bool:
    """True only if s is a date _clean_date_value() actually managed to
    PARSE (always comes back as YYYY-MM-DD when it succeeds) -- not
    just any non-blank text sitting in an Ending/Termination Date
    column. Some supervisors reuse that column for free-text notes
    ('next month vacation', 'move to another city in Sep', a bare
    'terminate' with no date) instead of a real date. Treating ANY
    non-blank value there as proof of termination was marking people
    Terminated off a casual note about a FUTURE plan, not an actual
    end date -- this is the check that keeps that from happening."""
    return bool(s) and bool(_YYYY_MM_DD_RE.match(s))


def _clean_month_value(v) -> str:
    if pd.isna(v):
        return ""
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.strftime("%Y-%m")
    s = str(v).strip()
    if not s or s.lower() in ("nan", "nat"):
        return ""
    if len(s) == 7 and s[4] == "-":
        return s
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%B %Y", "%b %Y", "%b-%y", "%m/%Y", "%Y/%m", "%B-%y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m")
        except ValueError:
            continue
    return s


_MONTH_NAME_TO_NUM = {
    "jan": 1, "january": 1, "feb": 2, "february": 2, "mar": 3, "march": 3,
    "apr": 4, "april": 4, "may": 5, "jun": 6, "june": 6, "jul": 7, "july": 7,
    "aug": 8, "august": 8, "sep": 9, "sept": 9, "september": 9, "oct": 10,
    "october": 10, "nov": 11, "november": 11, "dec": 12, "december": 12,
}


def _infer_month_from_filename(filename: str) -> str:
    name = filename.rsplit(".", 1)[0]
    lower = name.lower().replace("-", "_").replace(" ", "_")
    tokens = lower.split("_")

    year_match = re.search(r"(20\d{2})", name)
    year = int(year_match.group(1)) if year_match else datetime.today().year

    for token in tokens:
        if token in _MONTH_NAME_TO_NUM:
            return f"{year:04d}-{_MONTH_NAME_TO_NUM[token]:02d}"

    ym_match = re.search(r"(20\d{2})[-_]?(\d{2})", name)
    if ym_match:
        y, m = int(ym_match.group(1)), int(ym_match.group(2))
        if 1 <= m <= 12:
            return f"{y:04d}-{m:02d}"

    return ""


def _clean_vehicle_type(v):
    if pd.isna(v):
        return None
    s = str(v).strip().lower()
    if not s:
        return None
    if "company" in s:
        return "Company Car"
    if "own" in s:
        return "Own Car"
    return None


def _slugify_name_to_id(name: str) -> str:
    slug = "".join(ch if ch.isalnum() else "-" for ch in name.upper()).strip("-")
    while "--" in slug:
        slug = slug.replace("--", "-")
    return f"GEN-{slug[:40]}"


def _process_upload_mapped(df: pd.DataFrame, mapping: dict, default_month: str, include_payroll: bool):
    def col(field):
        src = mapping.get(field, NONE_OPTION)
        return src if src != NONE_OPTION else None

    conn = get_connection()
    success_count = 0
    errors = []

    existing_drivers = load_drivers()
    known_driver_ids = set(existing_drivers["driver_id"])
    name_to_id = dict(zip(existing_drivers["driver_name"].str.strip().str.upper(), existing_drivers["driver_id"]))

    for idx, raw in df.iterrows():
        row_num = idx + 2
        if _row_is_summary(raw):
            continue
        try:
            name_col = col("driver_name")
            first_col, last_col = col("first_name"), col("last_name")
            if name_col:
                driver_name = str(raw[name_col]).strip()
            elif first_col or last_col:
                first = "" if not first_col or pd.isna(raw[first_col]) else str(raw[first_col]).strip()
                last = "" if not last_col or pd.isna(raw[last_col]) else str(raw[last_col]).strip()
                driver_name = f"{first} {last}".strip()
            else:
                driver_name = ""

            if not driver_name or driver_name.lower() == "nan":
                raise ValueError("no driver name found for this row")

            id_col = col("driver_id")
            driver_id = _clean_id_value(raw[id_col]) if id_col else ""
            if driver_id and driver_id not in known_driver_ids:
                name_match = _fuzzy_match_name_to_id(name_to_id, driver_name)
                if name_match:
                    driver_id = name_match
            if not driver_id:
                driver_id = _slugify_name_to_id(driver_name)

            join_date = _clean_date_value(raw[col("join_date")]) if col("join_date") else None
            termination_date = _clean_date_value(raw[col("termination_date")]) if col("termination_date") else None

            status_col = col("status")
            if status_col:
                status = str(raw[status_col]).strip()
                if status not in DRIVER_STATUSES:
                    status = "Terminated" if termination_date else "Active"
            else:
                status = "Terminated" if termination_date else "Active"

            veh_col = col("vehicle_type")
            vehicle_type = _clean_vehicle_type(raw[veh_col]) if veh_col else None

            phone_col, sup_col = col("phone"), col("supervisor_name")
            phone = None if not phone_col or pd.isna(raw[phone_col]) else str(raw[phone_col]).strip()
            supervisor_name = None if not sup_col or pd.isna(raw[sup_col]) else str(raw[sup_col]).strip()

            sponsor_col, iqama_col = col("sponsor_name"), col("iqama_number")
            sponsor_name = None if not sponsor_col or pd.isna(raw[sponsor_col]) else str(raw[sponsor_col]).strip()
            iqama_number = None if not iqama_col else _clean_id_value(raw[iqama_col])
            if iqama_number == "":
                iqama_number = None

            upsert_driver(
                conn,
                {
                    "driver_id": driver_id,
                    "driver_name": driver_name,
                    "phone": phone,
                    "supervisor_name": supervisor_name,
                    "status": status,
                    "vehicle_type": vehicle_type,
                    "join_date": join_date,
                    "termination_date": termination_date,
                    "sponsor_name": sponsor_name,
                    "iqama_number": iqama_number,
                },
            )

            if include_payroll:
                my_col = col("month_year")
                month_year = _clean_month_value(raw[my_col]) if my_col else ""
                if not month_year:
                    month_year = default_month
                if not month_year:
                    raise ValueError("no payroll month available (map a Month column or pick one above)")

                orders_col = col("total_orders")
                total_orders = _clean_number_value(raw[orders_col], as_int=True) if orders_col else None

                cancel_col = col("cancelled_orders")
                cancelled_orders = _clean_number_value(raw[cancel_col], as_int=True) if cancel_col else None

                gross_col = col("gross_salary")
                gross_salary = _clean_number_value(raw[gross_col]) if gross_col else None

                ded_col = col("total_deductions")
                total_deductions = _clean_number_value(raw[ded_col]) if ded_col else None

                pending_col = col("pending_salary")
                pending_salary = _clean_number_value(raw[pending_col]) if pending_col else None

                net_col = col("net_salary")
                net_salary = (
                    _clean_number_value(raw[net_col])
                    if net_col and not pd.isna(raw[net_col])
                    else None
                )

                val_col = col("validity_status")
                if val_col and not pd.isna(raw[val_col]):
                    validity_status = str(raw[val_col]).strip()
                    if validity_status not in VALIDITY_STATUSES:
                        validity_status = "Valid"
                else:
                    validity_status = None

                days_col = col("valid_days_in_month")
                valid_days_in_month = _clean_number_value(raw[days_col], as_int=True) if days_col else None

                merge_monthly_log(
                    conn,
                    driver_id,
                    month_year,
                    {
                        "total_orders": total_orders,
                        "cancelled_orders": cancelled_orders,
                        "gross_salary": gross_salary,
                        "total_deductions": total_deductions,
                        "pending_salary": pending_salary,
                        "net_salary": net_salary,
                        "validity_status": validity_status,
                        "valid_days_in_month": valid_days_in_month,
                    },
                )

            success_count += 1

        except Exception as exc:  # noqa: BLE001
            errors.append(f"Row {row_num}: {exc}")

    conn.commit()
    conn.close()
    return success_count, errors


def _unnamed_ratio(columns) -> float:
    if len(columns) == 0:
        return 1.0
    unnamed = sum(1 for c in columns if str(c).strip().lower().startswith("unnamed"))
    return unnamed / len(columns)


def _read_excel_smart(file_obj, sheet_name):
    file_obj.seek(0)
    df0 = pd.read_excel(file_obj, sheet_name=sheet_name, header=0)

    try:
        file_obj.seek(0)
        df1 = pd.read_excel(file_obj, sheet_name=sheet_name, header=1)
    except Exception:  # noqa: BLE001 - sheet too short for a second header row
        df0.columns = [str(c).strip() for c in df0.columns]
        return df0, 1

    if _unnamed_ratio(df1.columns) < _unnamed_ratio(df0.columns):
        df1.columns = [str(c).strip() for c in df1.columns]
        return df1, 2
    df0.columns = [str(c).strip() for c in df0.columns]
    return df0, 1


# ==============================================================================
# WHOLE-WORKBOOK AUTO-IMPORT (every sheet, auto-classified)
# ==============================================================================

VALIDITY_TOKENS = {"VALID", "INVALID"}
ATTENDANCE_TOKENS = {"P", "OFF", "A", "ABSENT", "PRESENT", "OFFDAY", "OFF DAY", "LEAVE"}

SUMMARY_ROW_TOKENS = {"total", "totals", "grand total", "grand totals", "sum", "overall", "subtotal"}

# Legend/footer rows some report sheets tack on below the real rider rows,
# e.g. "TOTAL VAID  93", "TOTAL INVALID  3", "CATAEGORY A  28" -- these put
# a summary label in the driver_id/name column and a plain COUNT (not that
# rider's data) next to it. Matched by cell PREFIX (not just exact token)
# so typos like "VAID" for "VALID" or "CATAEGORY" for "CATEGORY" still hit,
# since the point is the label shape, not exact spelling.
_SUMMARY_ROW_PREFIX_RE = re.compile(r"^(total|grand\s*total|subtotal|overall|sum|cat\w*gory)\b", re.I)


def _row_is_summary(raw_row) -> bool:
    for v in raw_row:
        if pd.isna(v):
            continue
        s = str(v).strip().lower()
        if s in SUMMARY_ROW_TOKENS or "grand total" in s:
            return True
        if _SUMMARY_ROW_PREFIX_RE.match(s):
            return True
    return False


# ---- Section-header status markers (roster sheets that group riders by a
# free-text label row -- e.g. "TERMINATE FOR THIS MONTH", "ON VOCATION" --
# instead of a per-row status/ending-date column) --------------------------
#
# Real roster exports often look like:
#   [96 normal rider rows]
#   [blank rows]
#   TERMINATE FOR THIS MONTH        <- everything below this, until the
#   [11 rider rows]                    next marker or end of sheet, is
#   [blank rows]                       Terminated -- NOT a per-row value.
#   ON VOCATION
#   [1 rider row]
#
# Without recognizing this, every one of those riders silently defaults to
# "Active" (the only fallback when no status/ending-date column exists),
# which is what made Active == Total Headcount and Terminated == 0 even
# though the file clearly lists terminated/on-leave riders.
_SECTION_STATUS_MARKERS = [
    (re.compile(r"terminat", re.I), "Terminated"),
    (re.compile(r"vacation|vocation", re.I), "Suspended"),
    (re.compile(r"suspend", re.I), "Suspended"),
    (re.compile(r"\bon\s*leave\b", re.I), "Suspended"),
    (re.compile(r"^\s*active\s*$", re.I), "Active"),
]


def _detect_status_section_marker(raw_row):
    """If this row is a section-header label (very few filled cells, and
    the text matches a known status word) return the status that should
    apply to every rider row that follows, until the next marker. Returns
    None for an ordinary rider row (or a blank spacer row)."""
    non_null = [str(v).strip() for v in raw_row if pd.notna(v) and str(v).strip() != ""]
    if not non_null or len(non_null) > 2:
        return None
    joined = " ".join(non_null)
    for pattern, status in _SECTION_STATUS_MARKERS:
        if pattern.search(joined):
            return status
    return None


# ==============================================================================
# SALARY WORKBOOK UPLOAD  (separate from the roster/orders/validity upload)
# ==============================================================================

SALARY_DETAIL_ALIASES = {
    "driver_id": ["courier id"],
    "driver_name": ["courier name"],
    "billing_cycle": ["billing cycle"],
    "gross_amount": ["total payable amount", "total salary"],
    "total_deductions": ["total deduction"],
    "final_salary": ["final salary"],
    "pending_amount": ["pending", "pend"],
}

SALARY_SUMMARY_ALIASES = {
    "billing_cycle": ["billing cycle"],
    "tax_amount": ["tax amount"],
    "total_payable": ["total payable amount"],
    "invoice_amount": ["invoice amount"],
}


def _guess_column_exact(columns: list, aliases: list) -> str:
    normed = {c: _normalize_header(c) for c in columns}
    for col, norm in normed.items():
        if norm in aliases:
            return col
    return NONE_OPTION


def _ensure_salary_summary_table(conn) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS salary_summary (
            month_year     TEXT PRIMARY KEY,
            total_payable  REAL,
            tax_amount     REAL,
            invoice_amount REAL
        )
        """
    )


def _classify_salary_sheet(df: pd.DataFrame) -> str:
    cols = list(df.columns)
    has_courier_id = _guess_column_exact(cols, SALARY_DETAIL_ALIASES["driver_id"]) != NONE_OPTION
    has_final_salary = _guess_column_exact(cols, SALARY_DETAIL_ALIASES["final_salary"]) != NONE_OPTION
    has_total_ded = _guess_column_exact(cols, SALARY_DETAIL_ALIASES["total_deductions"]) != NONE_OPTION
    has_tax = _guess_column_exact(cols, SALARY_SUMMARY_ALIASES["tax_amount"]) != NONE_OPTION

    if has_tax and not has_courier_id:
        return "salary_summary"
    if has_courier_id and (has_final_salary or has_total_ded):
        return "salary_detail"
    return "unrecognized"


def _extract_salary_detail(df: pd.DataFrame, default_month: str):
    cols = list(df.columns)
    id_col = _guess_column_exact(cols, SALARY_DETAIL_ALIASES["driver_id"])
    name_col = _guess_column_exact(cols, SALARY_DETAIL_ALIASES["driver_name"])
    billing_col = _guess_column_exact(cols, SALARY_DETAIL_ALIASES["billing_cycle"])
    gross_col = _guess_column_exact(cols, SALARY_DETAIL_ALIASES["gross_amount"])
    ded_col = _guess_column_exact(cols, SALARY_DETAIL_ALIASES["total_deductions"])
    final_col = _guess_column_exact(cols, SALARY_DETAIL_ALIASES["final_salary"])
    pending_col = _guess_column_exact(cols, SALARY_DETAIL_ALIASES["pending_amount"])

    if id_col == NONE_OPTION:
        return [], {}

    rows = []
    id_to_name = {}
    for _, raw in df.iterrows():
        if _row_is_summary(raw):
            continue
        driver_id = _clean_id_value(raw[id_col])
        if not driver_id:
            continue

        month_year = default_month
        if billing_col != NONE_OPTION and not pd.isna(raw[billing_col]):
            parsed = _clean_month_value(raw[billing_col])
            if parsed:
                month_year = parsed

        name = None
        if name_col != NONE_OPTION and not pd.isna(raw[name_col]):
            name = str(raw[name_col]).strip()
            if name:
                id_to_name[driver_id] = name

        rows.append({
            "driver_id": driver_id,
            "month_year": month_year,
            "gross_salary": _clean_number_value(raw[gross_col]) if gross_col != NONE_OPTION and not pd.isna(raw[gross_col]) else None,
            "total_deductions": _clean_number_value(raw[ded_col]) if ded_col != NONE_OPTION and not pd.isna(raw[ded_col]) else None,
            "net_salary": _clean_number_value(raw[final_col]) if final_col != NONE_OPTION and not pd.isna(raw[final_col]) else None,
            "pending_salary": _clean_number_value(raw[pending_col]) if pending_col != NONE_OPTION and not pd.isna(raw[pending_col]) else None,
        })
    return rows, id_to_name


def _extract_salary_summary(df: pd.DataFrame, default_month: str) -> list:
    cols = list(df.columns)
    billing_col = _guess_column_exact(cols, SALARY_SUMMARY_ALIASES["billing_cycle"])
    tax_col = _guess_column_exact(cols, SALARY_SUMMARY_ALIASES["tax_amount"])
    payable_col = _guess_column_exact(cols, SALARY_SUMMARY_ALIASES["total_payable"])
    invoice_col = _guess_column_exact(cols, SALARY_SUMMARY_ALIASES["invoice_amount"])

    records = []
    for _, raw in df.iterrows():
        if _row_is_summary(raw):
            continue
        month_year = default_month
        if billing_col != NONE_OPTION and not pd.isna(raw[billing_col]):
            parsed = _clean_month_value(raw[billing_col])
            if parsed:
                month_year = parsed
        records.append({
            "month_year": month_year,
            "tax_amount": _clean_number_value(raw[tax_col]) if tax_col != NONE_OPTION and not pd.isna(raw[tax_col]) else None,
            "total_payable": _clean_number_value(raw[payable_col]) if payable_col != NONE_OPTION and not pd.isna(raw[payable_col]) else None,
            "invoice_amount": _clean_number_value(raw[invoice_col]) if invoice_col != NONE_OPTION and not pd.isna(raw[invoice_col]) else None,
        })
    return records


def process_salary_workbook(uploaded_file, default_month: str) -> dict:
    xls = pd.ExcelFile(uploaded_file)

    sheet_report = []
    detail_rows = []
    id_to_name = {}
    summary_records = []

    for sheet_name in xls.sheet_names:
        try:
            df, _hdr = _read_excel_smart(uploaded_file, sheet_name)
        except Exception as exc:  # noqa: BLE001
            sheet_report.append((sheet_name, f"error reading sheet: {exc}", 0))
            continue

        df = df.dropna(axis=0, how="all").reset_index(drop=True)
        df.columns = _dedupe_headers([str(c).strip() for c in df.columns])
        if df.empty:
            sheet_report.append((sheet_name, "empty", 0))
            continue

        kind = _classify_salary_sheet(df)
        sheet_report.append((sheet_name, kind, len(df)))

        if kind == "salary_detail":
            rows, names = _extract_salary_detail(df, default_month)
            detail_rows.extend(rows)
            id_to_name.update(names)
        elif kind == "salary_summary":
            summary_records.extend(_extract_salary_summary(df, default_month))

    conn = get_connection()
    all_drivers_df = load_drivers()
    known_driver_ids = set(all_drivers_df["driver_id"])
    name_to_id = dict(zip(all_drivers_df["driver_name"].str.strip().str.upper(), all_drivers_df["driver_id"]))

    riders_updated = set()
    placeholders_created = 0
    for row in detail_rows:
        driver_id = row["driver_id"]
        if driver_id not in known_driver_ids:
            name = id_to_name.get(driver_id)
            matched = _fuzzy_match_name_to_id(name_to_id, name) if name else None
            if matched:
                driver_id = matched
            else:
                upsert_driver(
                    conn,
                    {
                        "driver_id": driver_id,
                        "driver_name": name or f"Unknown Rider ({driver_id})",
                        "status": "Active",
                    },
                )
                known_driver_ids.add(driver_id)
                placeholders_created += 1

        merge_monthly_log(
            conn,
            driver_id,
            row["month_year"],
            {
                "gross_salary": row["gross_salary"],
                "total_deductions": row["total_deductions"],
                "net_salary": row["net_salary"],
                "pending_salary": row["pending_salary"],
            },
        )
        riders_updated.add(driver_id)

    _ensure_salary_summary_table(conn)
    for rec in summary_records:
        conn.execute(
            """
            INSERT INTO salary_summary (month_year, total_payable, tax_amount, invoice_amount)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(month_year) DO UPDATE SET
                total_payable  = COALESCE(excluded.total_payable, salary_summary.total_payable),
                tax_amount     = COALESCE(excluded.tax_amount, salary_summary.tax_amount),
                invoice_amount = COALESCE(excluded.invoice_amount, salary_summary.invoice_amount)
            """,
            (rec["month_year"], rec["total_payable"], rec["tax_amount"], rec["invoice_amount"]),
        )

    conn.commit()
    conn.close()

    return {
        "sheet_report": sheet_report,
        "riders_updated": len(riders_updated),
        "placeholders_created": placeholders_created,
        "summary_written": len(summary_records),
    }


def _day_number_columns(columns) -> list:
    out = []
    for c in columns:
        s = str(c).strip()
        if s.isdigit() and 1 <= int(s) <= 31:
            out.append(c)
            continue
        date_part = s.split(" ")[0]
        matched = False
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
            try:
                datetime.strptime(date_part, fmt)
                matched = True
                break
            except ValueError:
                continue
        if matched:
            out.append(c)
    return out


def _day_column_to_daynum(col):
    """Turn a day-by-day column header -- either a plain day number
    ('1'..'31') or an actual date ('2026-08-15') -- into just the day-
    of-month integer, so a driver's daily orders/validity/attendance
    can be stored and charted against a simple 1..31 axis regardless
    of which style the sheet uses.

    DD/MM/YYYY is tried BEFORE MM/DD/YYYY -- not the other way round.
    Both are ambiguous for a header like '01/08/2026', but trying
    MM/DD first silently mis-happy-parses every one of '01/08'..'12/08'
    as 'month=that number, day=8' (since 8 is a valid day in any
    month), collapsing 12 different days' worth of data into a single
    'day 8' bucket. DD/MM/YYYY is both the far more common convention
    for the sheets this importer actually sees, and self-correcting
    where it matters: a header with day > 12 (e.g. '13/08/2026') can
    ONLY be DD/MM anyway, since no month exceeds 12."""
    s = str(col).strip()
    if s.isdigit() and 1 <= int(s) <= 31:
        return int(s)
    date_part = s.split(" ")[0]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(date_part, fmt).day
        except ValueError:
            continue
    return None


def _extract_daily_orders(df: pd.DataFrame) -> dict:
    """driver_id -> {day_number: orders} from a day-by-day orders
    sheet. Skipped entirely for a sheet whose orders come from a
    single Total-Orders-style column instead of per-day columns --
    there's no day-level detail to capture in that case."""
    cols = list(df.columns)
    id_col = _guess_column(cols, FIELD_ALIASES["driver_id"])
    day_cols = _day_number_columns(cols)
    if id_col == NONE_OPTION or not day_cols:
        return {}
    out = {}
    for _, raw in df.iterrows():
        if _row_is_summary(raw):
            continue
        driver_id = _clean_id_value(raw[id_col])
        if not driver_id:
            continue
        day_map = out.setdefault(driver_id, {})
        for c in day_cols:
            daynum = _day_column_to_daynum(c)
            if daynum is None or pd.isna(raw[c]):
                continue
            day_map[daynum] = day_map.get(daynum, 0) + _clean_number_value(raw[c], as_int=True)
    return out


def _extract_daily_validity(df: pd.DataFrame) -> dict:
    """driver_id -> {day_number: 'Valid'/'Invalid'} from a day-by-day
    validity sheet."""
    cols = list(df.columns)
    id_col = _guess_column(cols, FIELD_ALIASES["driver_id"])
    day_cols = _day_number_columns(cols)
    if id_col == NONE_OPTION or not day_cols:
        return {}
    out = {}
    for _, raw in df.iterrows():
        if _row_is_summary(raw):
            continue
        driver_id = _clean_id_value(raw[id_col])
        if not driver_id:
            continue
        day_map = out.setdefault(driver_id, {})
        for c in day_cols:
            daynum = _day_column_to_daynum(c)
            if daynum is None or pd.isna(raw[c]):
                continue
            val = str(raw[c]).strip().upper()
            if val in ("VALID", "INVALID"):
                day_map[daynum] = val.capitalize()
    return out


def _extract_daily_attendance(df: pd.DataFrame) -> dict:
    """driver_id -> {day_number: 'Present'/'Absent'/'Off'} from a
    day-by-day attendance sheet."""
    cols = list(df.columns)
    id_col = _guess_column(cols, FIELD_ALIASES["driver_id"])
    day_cols = _day_number_columns(cols)
    if id_col == NONE_OPTION or not day_cols:
        return {}
    out = {}
    for _, raw in df.iterrows():
        if _row_is_summary(raw):
            continue
        driver_id = _clean_id_value(raw[id_col])
        if not driver_id:
            continue
        day_map = out.setdefault(driver_id, {})
        for c in day_cols:
            daynum = _day_column_to_daynum(c)
            if daynum is None or pd.isna(raw[c]):
                continue
            val = str(raw[c]).strip().upper()
            if val in ("P", "PRESENT"):
                day_map[daynum] = "Present"
            elif val in ("A", "ABSENT"):
                day_map[daynum] = "Absent"
            elif val in ("OFF", "OFFDAY", "OFF DAY", "LEAVE"):
                day_map[daynum] = "Off"
    return out


def upsert_daily_logs(conn: sqlite3.Connection, month_year: str, daily_orders: dict,
                       daily_validity: dict, daily_attendance: dict) -> int:
    """Merge per-day orders/validity/attendance maps (each driver_id ->
    {day: value}) into the daily_logs table for month_year. Existing
    rows are updated in place (same UPSERT-and-COALESCE philosophy as
    merge_monthly_log) so re-syncing the same month never duplicates
    or wipes out a field a different sheet already filled in."""
    all_driver_ids = set(daily_orders) | set(daily_validity) | set(daily_attendance)
    rows_written = 0
    for driver_id in all_driver_ids:
        days = (
            set(daily_orders.get(driver_id, {}))
            | set(daily_validity.get(driver_id, {}))
            | set(daily_attendance.get(driver_id, {}))
        )
        for day in days:
            orders_val = daily_orders.get(driver_id, {}).get(day)
            validity_val = daily_validity.get(driver_id, {}).get(day)
            attendance_val = daily_attendance.get(driver_id, {}).get(day)
            conn.execute(
                """
                INSERT INTO daily_logs (driver_id, month_year, day, orders, validity, attendance)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(driver_id, month_year, day) DO UPDATE SET
                    orders     = COALESCE(excluded.orders, daily_logs.orders),
                    validity   = COALESCE(excluded.validity, daily_logs.validity),
                    attendance = COALESCE(excluded.attendance, daily_logs.attendance)
                """,
                (driver_id, month_year, day, orders_val, validity_val, attendance_val),
            )
            rows_written += 1
    return rows_written


def load_daily_logs(driver_id: str, month_year: str) -> pd.DataFrame:
    conn = get_connection()
    df = pd.read_sql_query(
        "SELECT day, orders, validity, attendance FROM daily_logs "
        "WHERE driver_id = ? AND month_year = ? ORDER BY day",
        conn,
        params=(driver_id, month_year),
    )
    conn.close()
    return df


def _classify_sheet(df: pd.DataFrame) -> str:
    cols = list(df.columns)

    has_name = _guess_column(cols, FIELD_ALIASES["driver_name"]) != NONE_OPTION
    id_col = _guess_column(cols, FIELD_ALIASES["driver_id"])
    has_id = id_col != NONE_OPTION

    has_vehicle_hint = _guess_column(cols, FIELD_ALIASES["vehicle_type"]) != NONE_OPTION
    if not has_vehicle_hint and has_name and has_id:
        name_col = _guess_column(cols, FIELD_ALIASES["driver_name"])
        detected_col, _mode = _detect_vehicle_source_column(df, {id_col, name_col})
        if detected_col is not None:
            has_vehicle_hint = True

    if has_name and has_id and has_vehicle_hint:
        return "roster"

    day_cols = _day_number_columns(cols)
    if len(day_cols) >= 15:
        sample = set()
        numeric_hits, checked = 0, 0
        for c in day_cols[:8]:
            vals = df[c].dropna().astype(str).str.strip()
            for v in vals.head(25):
                checked += 1
                vu = v.upper()
                sample.add(vu)
                if re.match(r"^-?\d+(\.\d+)?$", v):
                    numeric_hits += 1
        if sample & VALIDITY_TOKENS:
            return "validity"
        if numeric_hits and checked and (numeric_hits / checked) >= 0.4:
            return "orders"
        if sample & ATTENDANCE_TOKENS:
            return "attendance"
        return "unrecognized"

    if _guess_column(cols, FIELD_ALIASES["total_orders"]) != NONE_OPTION:
        lower_cols = [str(c).strip().lower() for c in cols]
        if any("need more" in c or "shift time" in c for c in lower_cols):
            return "unrecognized"
        return "orders"

    lower_cols = [str(c).strip().lower() for c in cols]
    if any("order id" in c for c in lower_cols) and any("name" in c for c in lower_cols) and not has_id:
        return "cancellation"

    return "unrecognized"


_OWN_VEHICLE_RE = re.compile(r"\bown\s*(?:car|bike|vehicle|moto|scooter)\b", re.I)
_COMPANY_CAR_RE = re.compile(r"\bcompany\s*car\b", re.I)


def _detect_vehicle_source_column(df: pd.DataFrame, exclude: set):
    """Find whichever column actually carries Own-vs-Company-vehicle
    info, in EITHER of two real shapes seen across different months'
    files:

      'label' -- the column literally spells out 'Company Car' / 'Own
      Car' for nearly every row (some months reuse a column called
      'STATUS' for this, even though it has nothing to do with driver
      employment status). Read each value directly.

      'plate' -- the column is really a plate-number field (header
      'PLATE NUMBER' in every file that uses this shape): a rider on
      their own vehicle gets 'OWN CAR'/'OWN BIKE' text, a rider on a
      company-assigned vehicle gets an actual plate string or a rental/
      owner note like '(RENT)', '(TAMKEEN)', '(SALEEM)' -- so ANY
      non-blank value that isn't an 'own' mention means Company Car, by
      elimination. Barely any literal 'company car' text is what tells
      this shape apart from 'label' columns -- relying on a fixed
      column NAME here breaks the moment a file renames it, which is
      why this is detected from cell content instead.

    Returns (column, mode), or (None, None) if nothing reliable enough
    is found."""
    best_plate = None  # (col, own_hits)
    for col in df.columns:
        if col in exclude:
            continue
        vals = df[col].dropna().astype(str).str.strip()
        vals = vals[vals != ""]
        total = len(vals)
        if total < 5:
            continue
        upper = vals.str.upper()
        own_hits = int(upper.str.contains(_OWN_VEHICLE_RE).sum())
        company_hits = int(upper.str.contains(_COMPANY_CAR_RE).sum())
        label_ratio = (own_hits + company_hits) / total

        if label_ratio >= 0.85 and company_hits >= max(3, total * 0.1):
            return col, "label"

        if own_hits >= 3 and company_hits == 0:
            non_own = vals[~upper.str.contains(_OWN_VEHICLE_RE)]
            if len(non_own) == 0:
                continue
            # The remainder should look like short plate/rental tokens,
            # not free-form prose (e.g. a termination-reason column also
            # has no 'company car' text, but its values are sentences).
            platish = (non_own.str.len() <= 30).sum()
            if platish / len(non_own) >= 0.7:
                if best_plate is None or own_hits > best_plate[1]:
                    best_plate = (col, own_hits)

    if best_plate:
        return best_plate[0], "plate"
    return None, None


def _resolve_vehicle_type(raw_value, mode: str):
    """Read one cell's vehicle type given the column's detected mode
    (see _detect_vehicle_source_column). Returns None (leave unchanged)
    when the cell is blank or, in 'label' mode, doesn't recognizably say
    either type."""
    if pd.isna(raw_value):
        return None
    s = str(raw_value).strip()
    if not s:
        return None
    if mode == "label":
        return _clean_vehicle_type(s)
    # mode == "plate": presence of a real plate/rental value (anything
    # that isn't an "own vehicle" mention) means a company-assigned car.
    return "Own Car" if _OWN_VEHICLE_RE.search(s.upper()) else "Company Car"


def _vehicle_column_confidence(df: pd.DataFrame, col) -> float:
    """Fraction of a column's values that are recognizable own/company
    vehicle text -- used only to sanity-check a column found by HEADER
    NAME (e.g. an explicit 'Vehicle Type' header) before trusting it."""
    if col == NONE_OPTION:
        return 0.0
    vals = df[col].dropna().astype(str).str.strip()
    vals = vals[vals != ""]
    if vals.empty:
        return 0.0
    upper = vals.str.upper()
    hits = (upper.str.contains(_OWN_VEHICLE_RE) | upper.str.contains(_COMPANY_CAR_RE)).sum()
    return hits / len(vals)


_REASON_TEXT_HINTS = (
    "terminat", "close date", "restrict", "violation", "sponsership",
    "sponsorship", "without permission", "not working", "changed spons",
)


def _looks_like_reason_text(s: str) -> bool:
    """True if a 'last name' cell is actually a termination-reason note
    (e.g. 'Terminate, without permission off ID / close date...') rather
    than a real surname -- these sheets sometimes reuse the Last Name
    column to record why someone was let go. Folding that straight into
    driver_name produces garbage like 'SANA ULLAH Terminate, without
    permission off ID / close ...', so such text is dropped instead of
    appended."""
    if not s:
        return False
    low = s.lower()
    if len(s) > 25:
        return True
    return any(hint in low for hint in _REASON_TEXT_HINTS)


def _extract_roster(df: pd.DataFrame, month_year: str = None) -> dict:
    """driver_id -> roster fields dict, plus a name->id lookup for later
    name-based matching (used by the cancellation sheet, which has no ID).

    Rows are walked IN ORDER so that free-text section-header rows (e.g.
    "TERMINATE FOR THIS MONTH", "ON VOCATION") can be detected and applied
    to every rider row that follows, until the next marker -- see
    _detect_status_section_marker(). An explicit status/ending-date value
    on a row always wins over the section it happens to sit in.

    month_year (the payroll month this whole sheet is being imported for)
    is used ONLY to backfill termination_date when a section marker says
    "Terminated" but the row's own Ending Date cell is blank -- otherwise
    that rider would be correctly marked Terminated yet never counted in
    the "Terminated (this month)" tile, which keys off termination_date."""
    cols = list(df.columns)
    id_col = _guess_column(cols, FIELD_ALIASES["driver_id"])
    name_col = _guess_column(cols, FIELD_ALIASES["driver_name"])
    first_col = _guess_column(cols, FIELD_ALIASES["first_name"])
    last_col = _guess_column(cols, FIELD_ALIASES["last_name"])
    phone_col = _guess_column(cols, FIELD_ALIASES["phone"])
    sup_col = _guess_column(cols, FIELD_ALIASES["supervisor_name"])
    sponsor_col = _guess_column(cols, FIELD_ALIASES["sponsor_name"])
    iqama_col = _guess_column(cols, FIELD_ALIASES["iqama_number"])
    veh_col = _guess_column(cols, FIELD_ALIASES["vehicle_type"])
    status_col = _guess_column(cols, FIELD_ALIASES["status"])
    join_col = _guess_column(cols, FIELD_ALIASES["join_date"])
    end_col = _guess_column(cols, FIELD_ALIASES["termination_date"])
    orders_col = _guess_column(cols, FIELD_ALIASES["total_orders"])

    exclude = {id_col, name_col, first_col, last_col, phone_col, sup_col,
               sponsor_col, iqama_col, join_col, end_col}
    header_veh_col = veh_col  # an explicit "Vehicle Type"-named header, if any
    veh_mode = "label"
    detected_col, detected_mode = _detect_vehicle_source_column(df, exclude)
    if detected_col is not None and detected_col != header_veh_col:
        # A column that actually LOOKS like real vehicle info by its
        # values wins over a header-name guess whose own content doesn't
        # back it up (e.g. a "Vehicle" column that turned out blank/junk).
        header_confidence = _vehicle_column_confidence(df, header_veh_col) if header_veh_col != NONE_OPTION else 0.0
        if header_confidence < 0.5:
            veh_col, veh_mode = detected_col, detected_mode
    if veh_col != NONE_OPTION and status_col == veh_col:
        # Whatever we ended up using for vehicle type isn't really a
        # driver-status column -- don't feed vehicle-type text into status.
        status_col = NONE_OPTION

    records = {}
    current_section_status = None  # set by a section-header row, applies until the next one
    for _, raw in df.iterrows():
        marker = _detect_status_section_marker(raw)
        if marker:
            current_section_status = marker
            continue

        if _row_is_summary(raw):
            continue
        if name_col != NONE_OPTION:
            name = str(raw[name_col]).strip()
        elif first_col != NONE_OPTION or last_col != NONE_OPTION:
            f = "" if first_col == NONE_OPTION or pd.isna(raw[first_col]) else str(raw[first_col]).strip()
            l = "" if last_col == NONE_OPTION or pd.isna(raw[last_col]) else str(raw[last_col]).strip()
            if _looks_like_reason_text(l):
                l = ""
            name = f"{f} {l}".strip()
        else:
            name = ""
        if not name or name.lower() == "nan":
            continue
        if re.fullmatch(r"[\d.\-eE]+", name):
            # A malformed/shifted sheet (columns off by one, headers on
            # the wrong row) can make the NAME column resolve to what's
            # actually the Courier ID -- e.g. name ends up as
            # '1767086760257579.0'. A real person's name is never just
            # digits, so this row is unusable rather than a real rider.
            continue

        driver_id = _clean_id_value(raw[id_col]) if id_col != NONE_OPTION else ""
        if not driver_id:
            driver_id = _slugify_name_to_id(name)

        termination_date = _clean_date_value(raw[end_col]) if end_col != NONE_OPTION else None
        status_val = str(raw[status_col]).strip() if status_col != NONE_OPTION and not pd.isna(raw[status_col]) else ""
        if status_val in DRIVER_STATUSES:
            status = status_val
        elif _is_real_parsed_date(termination_date):
            status = "Terminated"
        elif current_section_status:
            status = current_section_status
        else:
            status = "Active"

        if status == "Terminated" and not _is_real_parsed_date(termination_date) and month_year:
            # The section said "Terminated" but this row's own Ending Date
            # cell was blank -- fall back to the 1st of the upload month so
            # this rider still counts in that month's "Terminated" tile
            # instead of silently vanishing from it.
            termination_date = f"{month_year}-01"

        records[driver_id] = {
            "driver_id": driver_id,
            "driver_name": name,
            "phone": None if phone_col == NONE_OPTION or pd.isna(raw[phone_col]) else str(raw[phone_col]).strip(),
            "supervisor_name": None if sup_col == NONE_OPTION or pd.isna(raw[sup_col]) else str(raw[sup_col]).strip(),
            "sponsor_name": None if sponsor_col == NONE_OPTION or pd.isna(raw[sponsor_col]) else str(raw[sponsor_col]).strip(),
            "iqama_number": None if iqama_col == NONE_OPTION else (_clean_id_value(raw[iqama_col]) or None),
            "vehicle_type": _resolve_vehicle_type(raw[veh_col], veh_mode) if veh_col != NONE_OPTION else None,
            "status": status,
            "join_date": _clean_date_value(raw[join_col]) if join_col != NONE_OPTION else None,
            "termination_date": termination_date,
            # If the roster tab ITSELF carries a per-rider order total
            # (some sheets do -- a 'Total Orders' column right there
            # alongside the roster), that figure is captured here so it
            # can take priority over whatever a separate day-by-day
            # orders tab computes for the same rider. Not written to
            # the drivers table (upsert_driver ignores unknown keys) --
            # only used by the orders-merging step below.
            "roster_total_orders": (
                _clean_number_value(raw[orders_col], as_int=True)
                if orders_col != NONE_OPTION and not pd.isna(raw[orders_col])
                else None
            ),
        }
    return records


def _extract_orders(df: pd.DataFrame):
    cols = list(df.columns)
    id_col = _guess_column(cols, FIELD_ALIASES["driver_id"])
    orders_col = _guess_column(cols, FIELD_ALIASES["total_orders"])
    name_col = _guess_column(cols, FIELD_ALIASES["driver_name"])
    if id_col == NONE_OPTION:
        return {}, {}

    day_cols = _day_number_columns(cols) if orders_col == NONE_OPTION else []
    if orders_col == NONE_OPTION and not day_cols:
        return {}, {}

    out = {}
    id_to_name = {}
    for _, raw in df.iterrows():
        if _row_is_summary(raw):
            continue
        driver_id = _clean_id_value(raw[id_col])
        if not driver_id:
            continue
        if orders_col != NONE_OPTION:
            row_total = _clean_number_value(raw[orders_col], as_int=True)
        else:
            row_total = sum(_clean_number_value(raw[c], as_int=True) for c in day_cols)
        out[driver_id] = out.get(driver_id, 0) + row_total
        if name_col != NONE_OPTION and not pd.isna(raw[name_col]):
            nm = str(raw[name_col]).strip()
            if nm:
                id_to_name[driver_id] = nm
    return out, id_to_name


def _diagnose_orders_sheet(df: pd.DataFrame) -> dict:
    """For the 'why is Orders wrong' diagnostic: shows whether this
    orders-tab has its OWN explicit total-orders column (which
    _extract_orders always prefers when present) versus what summing
    its day-by-day columns instead would give -- so a mismatch between
    the two is immediately visible without guessing. Mirrors
    _extract_orders() row-by-row EXACTLY (skipping grand-total/summary
    rows AND rows with no usable Driver ID) so these numbers are
    directly, honestly comparable to what the real extraction produces
    -- not inflated by rows that wouldn't actually count.

    Also flags a DIFFERENT, sneakier problem: two different day-column
    HEADERS that both resolve to the same calendar day (e.g. a bare
    '8' column AND some other column -- a stray weekly-total, a typo'd
    duplicate -- that also happens to parse as day 8). When that
    happens, that one day's orders get counted twice (once per
    matching column) for every rider, inflating just that single day
    far above its neighbors."""
    cols = list(df.columns)
    id_col = _guess_column(cols, FIELD_ALIASES["driver_id"])
    orders_col = _guess_column(cols, FIELD_ALIASES["total_orders"])
    day_cols = _day_number_columns(cols)
    found_explicit = orders_col != NONE_OPTION

    explicit_sum = 0
    day_sum = 0
    rows_with_id = 0
    rows_without_id = 0
    for _, raw in df.iterrows():
        if _row_is_summary(raw):
            continue
        has_id = id_col != NONE_OPTION and bool(_clean_id_value(raw[id_col]))
        if has_id:
            rows_with_id += 1
        else:
            rows_without_id += 1
            continue  # matches _extract_orders: a row with no usable ID contributes nothing
        if found_explicit:
            explicit_sum += _clean_number_value(raw[orders_col], as_int=True)
        if day_cols:
            day_sum += sum(_clean_number_value(raw[c], as_int=True) for c in day_cols)

    day_column_map = []
    day_num_counts = {}
    for c in day_cols:
        dn = _day_column_to_daynum(c)
        day_column_map.append((str(c), dn))
        if dn is not None:
            day_num_counts[dn] = day_num_counts.get(dn, 0) + 1
    colliding_days = sorted(dn for dn, cnt in day_num_counts.items() if cnt > 1)

    return {
        "orders_col_found": found_explicit,
        "orders_col_name": orders_col if found_explicit else None,
        "orders_col_sum": explicit_sum if found_explicit else None,
        "day_cols_count": len(day_cols),
        "day_cols_sum": day_sum,
        "used": "its own Total Orders column" if found_explicit else "summing its day-by-day columns",
        "rows_with_id": rows_with_id,
        "rows_without_id": rows_without_id,
        "day_column_map": day_column_map,
        "colliding_days": colliding_days,
    }


def _orders_dicts_look_like_duplicates(a: dict, b: dict) -> bool:
    common = set(a) & set(b)
    if not common or len(common) < max(1, min(len(a), len(b)) // 2):
        return False
    matches = sum(1 for k in common if a[k] == b[k])
    return (matches / len(common)) >= 0.8


def _extract_validity(df: pd.DataFrame):
    cols = list(df.columns)
    id_col = _guess_column(cols, FIELD_ALIASES["driver_id"])
    name_col = _guess_column(cols, FIELD_ALIASES["driver_name"])
    day_cols = _day_number_columns(cols)
    if id_col == NONE_OPTION or not day_cols:
        return {}, {}
    out = {}
    id_to_name = {}
    for _, raw in df.iterrows():
        if _row_is_summary(raw):
            continue
        driver_id = _clean_id_value(raw[id_col])
        if not driver_id:
            continue
        vals = [str(raw[c]).strip().upper() for c in day_cols if not pd.isna(raw[c])]
        valid_days = sum(1 for v in vals if v == "VALID")
        invalid_days = sum(1 for v in vals if v == "INVALID")
        out[driver_id] = {
            "valid_days": valid_days,
            "invalid_days": invalid_days,
            "status": "Invalid" if invalid_days > 0 else "Valid",
        }
        if name_col != NONE_OPTION and not pd.isna(raw[name_col]):
            nm = str(raw[name_col]).strip()
            if nm:
                id_to_name[driver_id] = nm
    return out, id_to_name


def _extract_attendance(df: pd.DataFrame):
    cols = list(df.columns)
    id_col = _guess_column(cols, FIELD_ALIASES["driver_id"])
    name_col = _guess_column(cols, FIELD_ALIASES["driver_name"])
    day_cols = _day_number_columns(cols)
    if id_col == NONE_OPTION or not day_cols:
        return {}, {}
    out = {}
    id_to_name = {}
    for _, raw in df.iterrows():
        if _row_is_summary(raw):
            continue
        driver_id = _clean_id_value(raw[id_col])
        if not driver_id:
            continue
        vals = [str(raw[c]).strip().upper() for c in day_cols if not pd.isna(raw[c])]
        out[driver_id] = sum(1 for v in vals if v == "P" or v == "PRESENT")
        if name_col != NONE_OPTION and not pd.isna(raw[name_col]):
            nm = str(raw[name_col]).strip()
            if nm:
                id_to_name[driver_id] = nm
    return out, id_to_name


def _extract_cancellations_by_name(df: pd.DataFrame) -> dict:
    cols = list(df.columns)
    name_col = None
    for c in cols:
        if "name" in str(c).strip().lower():
            name_col = c
            break
    if name_col is None:
        return {}
    out = {}
    for _, raw in df.iterrows():
        if _row_is_summary(raw):
            continue
        if pd.isna(raw[name_col]):
            continue
        key = str(raw[name_col]).strip().upper()
        if not key:
            continue
        out[key] = out.get(key, 0) + 1
    return out


_COMMON_NAME_TOKENS = {
    "MD", "MOHAMMAD", "MOHAMMED", "MUHAMMAD", "MUHAMMED", "SYED", "SK",
    "SHEIKH", "MST", "MISTER", "MOHD", "MD.", "SYD",
}


def _fuzzy_match_name_to_id(name_to_id: dict, target_name: str):
    target_upper = target_name.strip().upper()
    if target_upper in name_to_id:
        return name_to_id[target_upper]
    target_words = set(w for w in target_upper.split() if len(w) > 1 and w not in _COMMON_NAME_TOKENS)
    if len(target_words) < 2:
        return None
    best_id, best_overlap, tie = None, 0, False
    for roster_name, did in name_to_id.items():
        roster_words = set(w for w in roster_name.split() if len(w) > 1 and w not in _COMMON_NAME_TOKENS)
        overlap = len(target_words & roster_words)
        if overlap >= 2 and overlap > best_overlap:
            best_overlap = overlap
            best_id = did
            tie = False
        elif overlap == best_overlap and overlap >= 2 and did != best_id:
            tie = True
    return None if tie else best_id


def _remap_ids_by_name(id_dict: dict, id_to_name: dict, name_to_id: dict, known_ids: set) -> dict:
    if not id_dict:
        return id_dict
    remapped = {}
    for key, value in id_dict.items():
        target = key
        if key not in known_ids:
            name = id_to_name.get(key)
            if name:
                matched = _fuzzy_match_name_to_id(name_to_id, name)
                if matched:
                    target = matched
        if isinstance(value, dict):
            remapped[target] = value
        else:
            remapped[target] = remapped.get(target, 0) + value
    return remapped


_ROSTER_NAME_HINTS = ("active", "main", "employee", "rider", "driver")


def _score_roster_candidate(sheet_name: str, df: pd.DataFrame) -> float:
    """Rank how likely a sheet is to be the REAL master roster, when a
    workbook has more than one sheet that looks roster-shaped (e.g. a
    proper 'Active Rider' sheet plus an old backup/working copy like
    'RUF' that's mostly blank/'Unnamed' columns, or two near-duplicate
    snapshots like 'SHIFT TIMING' vs 'MIAN DATA'). Higher score wins.
    Real, filled-in Courier IDs matter most; a wall of 'Unnamed: N'
    columns (a stray/legend sheet, not a real export) is penalized hard;
    a sheet name that reads like an actual roster gets a small nudge."""
    cols = list(df.columns)
    id_col = _guess_column(cols, FIELD_ALIASES["driver_id"])
    valid_ids = int(df[id_col].apply(lambda v: bool(_clean_id_value(v))).sum()) if id_col != NONE_OPTION else 0
    unnamed_penalty = _unnamed_ratio(cols) * 100
    name_bonus = 15 if any(h in sheet_name.lower() for h in _ROSTER_NAME_HINTS) else 0
    return valid_ids - unnamed_penalty + name_bonus


def process_workbook_all_sheets(uploaded_file, month_year: str):
    """Read every sheet of an uploaded Excel workbook and hand them off
    to _process_sheet_frames(), which contains the actual classify/
    extract/merge logic. Kept separate so the exact same logic can be
    reused for a Google Sheet (see process_workbook_from_gsheet below)
    without re-reading through pandas.ExcelFile at all."""
    xls = pd.ExcelFile(uploaded_file)
    sheet_frames = []
    sheet_read_errors = []
    for sheet_name in xls.sheet_names:
        try:
            df, _hdr = _read_excel_smart(uploaded_file, sheet_name)
            sheet_frames.append((sheet_name, df))
        except Exception as exc:  # noqa: BLE001
            sheet_read_errors.append((sheet_name, f"error reading sheet: {exc}"))
    return _process_sheet_frames(sheet_frames, month_year, sheet_read_errors)


def _process_sheet_frames(sheet_frames: list, month_year: str, sheet_read_errors: list = None, roster_tab_override: str = None) -> dict:
    """The actual whole-workbook auto-import engine: classify each
    already-loaded sheet (roster / orders / validity / attendance /
    cancellation / unrecognized), extract and merge everything into the
    database for month_year. sheet_frames is a list of (sheet_name, df)
    pairs -- it doesn't matter whether those DataFrames came from an
    uploaded Excel file or a live Google Sheet, which is exactly what
    lets both sources share this one code path.

    roster_tab_override: when a workbook has more than one roster-
    shaped tab, auto-scoring (see _score_roster_candidate) can pick the
    wrong one. If given, this tab name (matched case-insensitively,
    whitespace-trimmed) is used as the roster directly, bypassing
    scoring entirely -- the admin's explicit choice always wins."""
    roster_candidates = []  # [(sheet_name, df)] -- resolved to ONE roster after the loop
    orders_sheets = []
    validity_by_id = {}
    attendance_by_id = {}
    cancellations_by_name = {}
    id_to_name = {}
    sheet_report = list(sheet_read_errors or [])
    sheet_report = [(name, err, 0) for name, err in sheet_report]
    override_df = None
    override_matched_name = None
    override_key = roster_tab_override.strip().lower() if roster_tab_override else None

    # Day-by-day detail (driver_id -> {day: value}), merged across
    # every orders/validity/attendance sheet seen -- captured here
    # purely for the Rider Lookup day-by-day drilldown/chart; none of
    # this changes the existing monthly totals logic below.
    daily_orders_all = {}
    daily_validity_all = {}
    daily_attendance_all = {}

    for sheet_name, df in sheet_frames:
        df = df.dropna(axis=0, how="all").reset_index(drop=True)
        df.columns = _dedupe_headers([str(c).strip() for c in df.columns])
        if override_key and sheet_name.strip().lower() == override_key:
            override_df = df
            override_matched_name = sheet_name
        if df.empty:
            sheet_report.append((sheet_name, "empty", 0))
            continue

        kind = _classify_sheet(df)
        sheet_report.append((sheet_name, kind, len(df)))

        if kind == "roster":
            roster_candidates.append((sheet_name, df))
        elif kind == "orders":
            orders_dict, names = _extract_orders(df)
            orders_sheets.append((sheet_name, orders_dict))
            id_to_name.update(names)
            for did, day_map in _extract_daily_orders(df).items():
                merged = daily_orders_all.setdefault(did, {})
                for day, val in day_map.items():
                    merged[day] = merged.get(day, 0) + val
        elif kind == "validity":
            validity_dict, names = _extract_validity(df)
            validity_by_id.update(validity_dict)
            id_to_name.update(names)
            for did, day_map in _extract_daily_validity(df).items():
                daily_validity_all.setdefault(did, {}).update(day_map)
        elif kind == "attendance":
            attendance_dict, names = _extract_attendance(df)
            attendance_by_id.update(attendance_dict)
            id_to_name.update(names)
            for did, day_map in _extract_daily_attendance(df).items():
                daily_attendance_all.setdefault(did, {}).update(day_map)
        elif kind == "cancellation":
            for k, v in _extract_cancellations_by_name(df).items():
                cancellations_by_name[k] = cancellations_by_name.get(k, 0) + v

    # Resolve to ONE primary roster when a workbook has several
    # roster-shaped sheets (a real 'Active Rider' sheet plus a stray
    # backup/working copy, or two overlapping snapshots). Merging every
    # one blindly used to inflate headcount by however many extra riders
    # the lower-quality sheet(s) added -- e.g. a mostly-blank 427-row
    # leftover sheet nearly doubling a real 89-rider roster. The
    # highest-scoring sheet (see _score_roster_candidate) is used as-is;
    # any OTHER roster-shaped sheet only contributes riders whose ID
    # isn't already in the primary sheet, so a genuinely unique rider
    # sitting only in a secondary sheet still isn't lost.
    #
    # If the admin explicitly named a tab (roster_tab_override), that
    # tab wins outright -- no scoring, no merging from other roster-
    # shaped tabs. Trust the human over the heuristic.
    roster_records = {}
    roster_sheet_used = None
    roster_sheets_skipped = []
    extra_riders_from_skipped = 0
    roster_override_requested_but_not_found = bool(roster_tab_override) and override_df is None

    if override_df is not None:
        roster_sheet_used = override_matched_name
        roster_records = _extract_roster(override_df, month_year)
        roster_sheets_skipped = [name for name, _ in roster_candidates if name != override_matched_name]
    elif roster_candidates:
        scored = sorted(
            ((_score_roster_candidate(name, df), name, df) for name, df in roster_candidates),
            key=lambda t: -t[0],
        )
        roster_sheet_used = scored[0][1]
        roster_records = _extract_roster(scored[0][2], month_year)
        for _score, name, df in scored[1:]:
            roster_sheets_skipped.append(name)
            # Only trust a lower-ranked sheet enough to pull EXTRA riders
            # from it when it's itself a clean, real export (mostly named
            # columns) -- a mangled/shifted sheet like a stray working
            # copy with a wall of 'Unnamed: N' columns produces garbage
            # rows (Courier ID misread as the rider's name, etc.) that
            # would just add phantom "riders" rather than real ones.
            if _unnamed_ratio(list(df.columns)) >= 0.3:
                continue
            secondary = _extract_roster(df, month_year)
            for did, rec in secondary.items():
                if did not in roster_records:
                    roster_records[did] = rec
                    extra_riders_from_skipped += 1

    orders_by_id = {}
    duplicate_orders_sheets = []
    kept_orders_sheets = []
    for sheet_name, sheet_orders in orders_sheets:
        is_dup = any(
            _orders_dicts_look_like_duplicates(sheet_orders, kept_dict)
            for _, kept_dict in kept_orders_sheets
        )
        if is_dup:
            duplicate_orders_sheets.append(sheet_name)
            continue
        kept_orders_sheets.append((sheet_name, sheet_orders))
        for k, v in sheet_orders.items():
            orders_by_id[k] = orders_by_id.get(k, 0) + v

    conn = get_connection()
    for record in roster_records.values():
        upsert_driver(conn, record)
    conn.commit()

    all_drivers_df = load_drivers()
    known_driver_ids = set(all_drivers_df["driver_id"])
    name_to_id = dict(zip(all_drivers_df["driver_name"].str.strip().str.upper(), all_drivers_df["driver_id"]))

    orders_by_id = _remap_ids_by_name(orders_by_id, id_to_name, name_to_id, known_driver_ids)
    validity_by_id = _remap_ids_by_name(validity_by_id, id_to_name, name_to_id, known_driver_ids)
    attendance_by_id = _remap_ids_by_name(attendance_by_id, id_to_name, name_to_id, known_driver_ids)
    daily_orders_all = _remap_ids_by_name(daily_orders_all, id_to_name, name_to_id, known_driver_ids)
    daily_validity_all = _remap_ids_by_name(daily_validity_all, id_to_name, name_to_id, known_driver_ids)
    daily_attendance_all = _remap_ids_by_name(daily_attendance_all, id_to_name, name_to_id, known_driver_ids)

    # If the roster tab itself carries a per-rider order total (see
    # _extract_roster), that figure wins over whatever a separate
    # day-by-day orders tab computed for the same rider -- the roster
    # sheet is the one the admin has confirmed is authoritative (via
    # Roster Tab Override), so its own numbers should be trusted over a
    # different tab that may be stale, duplicated, or miscounted.
    for did, rec in roster_records.items():
        roster_total = rec.get("roster_total_orders")
        if roster_total is not None:
            orders_by_id[did] = roster_total

    cancellations_by_id = {}
    unmatched_names = 0
    for name_key, count in cancellations_by_name.items():
        did = _fuzzy_match_name_to_id(name_to_id, name_key)
        if did:
            cancellations_by_id[did] = cancellations_by_id.get(did, 0) + count
        else:
            unmatched_names += 1

    all_driver_ids = (
        set(orders_by_id) | set(validity_by_id) | set(attendance_by_id)
        | set(cancellations_by_id) | set(roster_records)
    )

    known_ids = set(
        r[0] for r in conn.execute("SELECT driver_id FROM drivers").fetchall()
    )
    placeholders_created = 0
    for driver_id in all_driver_ids:
        if driver_id in roster_records or driver_id in known_ids:
            continue
        upsert_driver(
            conn,
            {
                "driver_id": driver_id,
                "driver_name": f"Unknown Rider ({driver_id})",
                "status": "Active",
            },
        )
        placeholders_created += 1
    conn.commit()

    logs_written = 0
    for driver_id in all_driver_ids:
        total_orders = orders_by_id.get(driver_id)
        cancelled_orders = cancellations_by_id.get(driver_id)
        val = validity_by_id.get(driver_id)
        att_days = attendance_by_id.get(driver_id)

        if val:
            valid_days_in_month = val["valid_days"]
            validity_status = val["status"]
        elif att_days is not None:
            valid_days_in_month = att_days
            validity_status = None
        else:
            valid_days_in_month = None
            validity_status = None

        merge_monthly_log(
            conn,
            driver_id,
            month_year,
            {
                "total_orders": total_orders,
                "cancelled_orders": cancelled_orders,
                "gross_salary": None,
                "total_deductions": None,
                "pending_salary": None,
                "net_salary": None,
                "validity_status": validity_status,
                "valid_days_in_month": valid_days_in_month,
                # Only ever SET this true (this rider is confirmed on
                # the roster sheet for this month) -- never pass 0/False
                # here, or a rider correctly marked in_roster by an
                # earlier upload for the same month would get wiped out
                # by a later upload (e.g. a salary file) that doesn't
                # carry its own roster sheet.
                "in_roster": 1 if driver_id in roster_records else None,
            },
        )
        logs_written += 1

    if roster_records:
        # This sync DID find a definitive roster (a real roster tab was
        # read, not just a payroll-only upload with no roster of its
        # own) -- so this is the one safe place to CLEAR a stale
        # in_roster flag: any driver_id still marked in_roster=1 for
        # this month but NOT in this sync's roster_records is left over
        # from an earlier, less accurate sync (e.g. before a Roster Tab
        # Override was set to fix the wrong sheet being auto-picked).
        # Without this cleanup, that stale flag sticks forever, because
        # the merge above only ever ADDS in_roster=1 and never removes
        # it -- by design, so a payroll-only upload never wipes out a
        # rider's roster membership. A sync that legitimately resolved
        # a roster is the one case allowed to correct it.
        placeholders = ",".join("?" for _ in roster_records)
        conn.execute(
            f"UPDATE monthly_logs SET in_roster = 0 "
            f"WHERE month_year = ? AND in_roster = 1 AND driver_id NOT IN ({placeholders})",
            (month_year, *roster_records.keys()),
        )

    daily_rows_written = upsert_daily_logs(conn, month_year, daily_orders_all, daily_validity_all, daily_attendance_all)

    conn.commit()
    conn.close()

    summary = {
        "sheet_report": sheet_report,
        "roster_count": len(roster_records),
        "orders_drivers": len(orders_by_id),
        "orders_total": sum(orders_by_id.values()),
        "validity_drivers": len(validity_by_id),
        "attendance_drivers": len(attendance_by_id),
        "cancellations_matched": len(cancellations_by_id),
        "cancellations_unmatched": unmatched_names,
        "logs_written": logs_written,
        "placeholders_created": placeholders_created,
        "duplicate_orders_sheets": duplicate_orders_sheets,
        "roster_sheet_used": roster_sheet_used,
        "roster_sheets_skipped": roster_sheets_skipped,
        "extra_riders_from_skipped": extra_riders_from_skipped,
        "roster_override_requested_but_not_found": roster_override_requested_but_not_found,
        "daily_rows_written": daily_rows_written,
    }
    return summary


# ==============================================================================
# GOOGLE SHEETS SYNC  -- daily entries by supervisors, live + monthly import
# ==============================================================================
#
# How this works, end to end:
#
#   1. Supervisors keep filling in a Google Sheet every day -- same
#      shape as an uploaded Excel workbook (a roster tab, an ORDER
#      REPORT tab with one column per date, an ATTENDANCE/VALIDITY tab
#      the same way, etc.). Nothing about how they work day-to-day
#      changes.
#
#   2. "Live This Month" (see render_live_month_panel) reads the sheet
#      directly, live, and shows month-to-date totals WITHOUT writing
#      anything to the database -- so the dashboard can always show
#      "how's this month looking so far" even while the month is still
#      in progress and numbers are still changing every day.
#
#   3. "Sync Now" (admin action, in this tab) pulls the whole sheet and
#      runs it through the EXACT SAME import engine used for a manual
#      Excel upload (_process_sheet_frames) -- same roster resolution,
#      same duplicate-sheet detection, same placeholder-rider handling.
#      This is what actually saves a month into permanent history.
#      Can be run any time (once a week to checkpoint, or right at
#      month-end) and re-running it for the same month just updates the
#      numbers in place, so it's always safe to click again.
#
#   4. For FULLY unattended month-end saving (nobody has to open the
#      dashboard and click a button), see the separate gsheet_sync.py
#      script -- it reuses these exact same functions but runs from an
#      OS-level scheduler (cron / Task Scheduler / cloud scheduler)
#      instead of from a Streamlit button click.

GSHEET_CONFIG_PATH = "hq_gsheet.json"


def _load_gsheet_config():
    if not os.path.exists(GSHEET_CONFIG_PATH):
        return None
    try:
        with open(GSHEET_CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:  # noqa: BLE001
        return None


def _save_gsheet_config(sheet_id: str) -> None:
    _update_gsheet_config(sheet_id=sheet_id.strip())


def _update_gsheet_config(**kwargs) -> dict:
    """Merge new fields into hq_gsheet.json without wiping out fields
    already set (sheet_id, auto_sync toggle, last-sync bookkeeping)."""
    config = _load_gsheet_config() or {}
    config.update(kwargs)
    with open(GSHEET_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f)
    return config


def _extract_sheet_id(url_or_id: str) -> str:
    """Accept either a raw Sheet ID or a full Google Sheets URL and
    return just the ID -- so the admin can paste whatever they have
    copied without needing to know which part matters."""
    s = url_or_id.strip()
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", s)
    if match:
        return match.group(1)
    return s


def _get_gspread_client():
    """Authenticate using a Google Cloud service account. The service
    account's credentials (a JSON key) must be stored in Streamlit's
    secrets as [gcp_service_account] -- see the setup instructions in
    render_gsheet_sync_tab(). Never hard-code credentials in this file."""
    if not GSPREAD_AVAILABLE:
        raise RuntimeError(
            "The 'gspread' and 'google-auth' packages aren't installed. "
            "Run: pip install gspread google-auth"
        )
    if "gcp_service_account" not in st.secrets:
        raise RuntimeError(
            "No Google service account configured. Add a [gcp_service_account] "
            "section to your Streamlit secrets (see setup instructions below)."
        )
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds = _GCredentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=scopes
    )
    return gspread.authorize(creds)


def _dedupe_headers(header: list) -> list:
    """Mimic what pandas.read_excel() already does automatically for a
    duplicated column name -- rename the 2nd, 3rd, etc. occurrence to
    'Name.1', 'Name.2', and so on. Google Sheets happily lets a
    supervisor's tab have two columns with the identical header text
    (a typo'd re-entry, a copy-pasted section, etc.); without this, a
    duplicate column label turns df[that_column] into a DataFrame
    instead of a single column of values everywhere downstream expects
    a plain Series -- which is what was crashing the sync with
    "'DataFrame' object has no attribute 'str'"."""
    seen = {}
    deduped = []
    for h in header:
        if h not in seen:
            seen[h] = 0
            deduped.append(h)
        else:
            seen[h] += 1
            deduped.append(f"{h}.{seen[h]}")
    return deduped


def _gsheet_worksheet_to_df(ws) -> pd.DataFrame:
    """Pull one Google Sheet tab into a DataFrame the same way
    _read_excel_smart() reads an Excel sheet: try row 1 as the header,
    try row 2 as the header (some tabs -- like ORDER REPORT -- have a
    title/section row first and the REAL column headers, e.g. actual
    dates, one row down), and keep whichever version has fewer blank/
    'Unnamed' columns."""
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()

    def _build(header_idx):
        if header_idx >= len(values):
            return None
        header = [h.strip() if h.strip() else f"Unnamed: {i}" for i, h in enumerate(values[header_idx])]
        header = _dedupe_headers(header)
        body = values[header_idx + 1:]
        if not body:
            return pd.DataFrame(columns=header)
        # Pad/trim each row to the header's width -- Google Sheets
        # rows can come back shorter than the header if trailing
        # cells are empty.
        width = len(header)
        fixed_body = [
            (row + [""] * width)[:width] if len(row) < width else row[:width]
            for row in body
        ]
        return pd.DataFrame(fixed_body, columns=header)

    df0 = _build(0)
    df1 = _build(1)

    if df1 is not None and not df1.empty and _unnamed_ratio(df1.columns) < _unnamed_ratio(df0.columns if df0 is not None else []):
        return df1
    return df0 if df0 is not None else pd.DataFrame()


def fetch_gsheet_frames(sheet_id: str) -> list:
    """Open the Google Sheet by ID and return every tab as a
    (sheet_name, DataFrame) pair -- the exact same shape
    process_workbook_all_sheets() builds from an uploaded Excel file,
    so both sources can share one importer (_process_sheet_frames)."""
    client = _get_gspread_client()
    sh = client.open_by_key(sheet_id)
    frames = []
    for ws in sh.worksheets():
        try:
            df = _gsheet_worksheet_to_df(ws)
        except Exception as exc:  # noqa: BLE001
            frames.append((ws.title, pd.DataFrame()))
            continue
        frames.append((ws.title, df))
    return frames


def _filter_ignored_tabs(frames: list, ignored_tabs) -> list:
    """Drop any (sheet_name, df) pair the Admin has explicitly marked
    to ignore (case-insensitive, whitespace-trimmed) -- e.g. a
    'PT ORDERS REPOER' tab covering a different order type that
    shouldn't count toward the main totals. Applied identically before
    BOTH the real Sync and the live view, so the two always agree."""
    if not ignored_tabs:
        return frames
    ignored_set = {t.strip().lower() for t in ignored_tabs if t.strip()}
    if not ignored_set:
        return frames
    return [(name, df) for name, df in frames if name.strip().lower() not in ignored_set]


def sync_month_from_gsheet(sheet_id: str, month_year: str, roster_tab_override: str = None) -> dict:
    """Pull the whole Google Sheet and import it into the database for
    month_year, using the exact same engine as a manual Excel upload.
    Safe to re-run any time -- existing drivers/months are updated in
    place, never duplicated. roster_tab_override: see
    _process_sheet_frames -- forces a specific tab name to be used as
    the roster instead of auto-scoring."""
    frames = fetch_gsheet_frames(sheet_id)
    config = _load_gsheet_config() or {}
    frames = _filter_ignored_tabs(frames, config.get("ignored_tabs"))
    return _process_sheet_frames(frames, month_year, roster_tab_override=roster_tab_override)


def fetch_live_month_to_date(sheet_id: str, roster_tab_override: str = None) -> dict:
    """A read-only peek at the Google Sheet's current numbers -- does
    NOT write anything to the database. Used to show "this month so
    far" on the live dashboard while the month is still in progress
    and supervisors are still filling in today's column. Reuses the
    same sheet-classification and name/ID-matching logic as the real
    importer (roster resolution, cancellation name-matching, etc.) so
    the live view matches what a real Sync would produce -- it just
    skips every database-writing step and additionally builds a
    per-rider detail table for display. roster_tab_override: see
    _process_sheet_frames -- forces a specific tab name to be used as
    the roster instead of auto-scoring."""
    frames = fetch_gsheet_frames(sheet_id)
    config = _load_gsheet_config() or {}
    frames = _filter_ignored_tabs(frames, config.get("ignored_tabs"))

    roster_candidates = []
    orders_sheets = []
    validity_by_id = {}
    attendance_by_id = {}
    cancellations_by_name = {}
    id_to_name = {}
    sheet_report = []  # (tab name, detected kind, row count) -- for the debugging expander
    override_df = None
    override_matched_name = None
    override_key = roster_tab_override.strip().lower() if roster_tab_override else None
    orders_sheet_columns = {}  # {sheet_name: [columns]} -- for the "why is orders wrong" diagnostic
    orders_sheet_diagnostics = {}  # {sheet_name: {...}} -- explicit-column vs day-sum comparison

    for sheet_name, df in frames:
        df = df.dropna(axis=0, how="all").reset_index(drop=True)
        if df.empty:
            sheet_report.append((sheet_name, "empty", 0))
            continue
        df.columns = _dedupe_headers([str(c).strip() for c in df.columns])
        if override_key and sheet_name.strip().lower() == override_key:
            override_df = df
            override_matched_name = sheet_name
        kind = _classify_sheet(df)
        sheet_report.append((sheet_name, kind, len(df)))
        if kind == "roster":
            roster_candidates.append((sheet_name, df))
        elif kind == "orders":
            orders_sheet_columns[sheet_name] = list(df.columns)
            orders_sheet_diagnostics[sheet_name] = _diagnose_orders_sheet(df)
            orders_dict, names = _extract_orders(df)
            orders_sheets.append((sheet_name, orders_dict))
            id_to_name.update(names)
        elif kind == "validity":
            validity_dict, names = _extract_validity(df)
            validity_by_id.update(validity_dict)
            id_to_name.update(names)
        elif kind == "attendance":
            attendance_dict, names = _extract_attendance(df)
            attendance_by_id.update(attendance_dict)
            id_to_name.update(names)
        elif kind == "cancellation":
            for k, v in _extract_cancellations_by_name(df).items():
                cancellations_by_name[k] = cancellations_by_name.get(k, 0) + v

    roster_records = {}
    roster_sheet_used = None
    roster_sheet_columns = []
    roster_sheets_seen = [name for name, _ in roster_candidates]
    roster_override_requested_but_not_found = bool(roster_tab_override) and override_df is None

    if override_df is not None:
        roster_sheet_used = override_matched_name
        roster_sheet_columns = list(override_df.columns)
        roster_records = _extract_roster(override_df)
    elif roster_candidates:
        scored = sorted(
            ((_score_roster_candidate(name, df), name, df) for name, df in roster_candidates),
            key=lambda t: -t[0],
        )
        roster_sheet_used = scored[0][1]
        roster_sheet_columns = list(scored[0][2].columns)
        roster_records = _extract_roster(scored[0][2])

    orders_by_id = {}
    duplicate_orders_sheets = []
    kept_orders_sheets = []
    for sheet_name, sheet_orders in orders_sheets:
        # Same safeguard as the real Sync: if two orders tabs largely
        # overlap on the same riders with matching numbers, they're
        # almost certainly the same data under two names -- counting
        # both would double a rider's orders. Only the first copy is
        # kept; skipped ones are surfaced in the diagnostics below.
        is_dup = any(
            _orders_dicts_look_like_duplicates(sheet_orders, kept_dict)
            for _, kept_dict in kept_orders_sheets
        )
        if is_dup:
            duplicate_orders_sheets.append(sheet_name)
            continue
        kept_orders_sheets.append((sheet_name, sheet_orders))
        for k, v in sheet_orders.items():
            orders_by_id[k] = orders_by_id.get(k, 0) + v

    # Match orders/validity/attendance/cancellation rows to a real
    # roster driver_id by name when their own ID wasn't on the roster
    # (typo, different numbering, etc.) -- same fuzzy-matching used by
    # the real Sync, just against the live roster instead of the DB.
    name_to_id = {r["driver_name"].strip().upper(): did for did, r in roster_records.items()}
    known_ids = set(roster_records.keys())
    orders_by_id = _remap_ids_by_name(orders_by_id, id_to_name, name_to_id, known_ids)
    validity_by_id = _remap_ids_by_name(validity_by_id, id_to_name, name_to_id, known_ids)
    attendance_by_id = _remap_ids_by_name(attendance_by_id, id_to_name, name_to_id, known_ids)

    # Same rule as the real Sync: if the roster tab itself carries a
    # per-rider order total, that wins over a separate day-by-day
    # orders tab for the same rider.
    for did, rec in roster_records.items():
        roster_total = rec.get("roster_total_orders")
        if roster_total is not None:
            orders_by_id[did] = roster_total

    cancellations_by_id = {}
    for name_key, count in cancellations_by_name.items():
        did = _fuzzy_match_name_to_id(name_to_id, name_key)
        if did:
            cancellations_by_id[did] = cancellations_by_id.get(did, 0) + count

    active_count = 0
    terminated_count = 0
    suspended_count = 0
    company_cars = 0
    own_cars = 0
    for did, r in roster_records.items():
        if r["status"] == "Terminated":
            terminated_count += 1
        elif r["status"] == "Suspended":
            suspended_count += 1
        else:
            worked = orders_by_id.get(did, 0) > 0 or (
                validity_by_id[did]["valid_days"] if did in validity_by_id else attendance_by_id.get(did, 0)
            ) > 0
            if worked:
                active_count += 1
        if r["vehicle_type"] == "Company Car":
            company_cars += 1
        elif r["vehicle_type"] == "Own Car":
            own_cars += 1

    # "Valid" / "Invalid" is the same whole-month PERFORMANCE target used
    # everywhere else -- both a minimum order count AND a minimum
    # days-worked count for the month, not the raw day-by-day marks a
    # Validity Report sheet might carry. Days-worked here comes from the
    # validity sheet's own valid-day count when available, else the
    # attendance sheet's present-day count -- same fallback the real
    # Sync uses when writing valid_days_in_month.
    targets = _load_validity_targets()
    all_known_ids = set(roster_records) | set(orders_by_id) | set(validity_by_id) | set(attendance_by_id)
    valid_count = 0
    invalid_count = 0
    performance_validity_by_id = {}
    for did in all_known_ids:
        days_worked = validity_by_id[did]["valid_days"] if did in validity_by_id else attendance_by_id.get(did, 0)
        status = compute_performance_validity(orders_by_id.get(did, 0), days_worked, targets["min_orders"], targets["min_days"])
        performance_validity_by_id[did] = status
        if status == "Valid":
            valid_count += 1
        else:
            invalid_count += 1

    all_ids = set(roster_records) | set(orders_by_id) | set(validity_by_id) | set(attendance_by_id) | set(cancellations_by_id)
    rows = []
    for did in all_ids:
        r = roster_records.get(did, {})
        val = validity_by_id.get(did)
        rows.append({
            "Driver ID": did,
            "Name": r.get("driver_name") or id_to_name.get(did, f"Unknown ({did})"),
            "Supervisor": r.get("supervisor_name") or "",
            "Status": r.get("status") or "(not on roster)",
            "Vehicle": r.get("vehicle_type") or "",
            "Orders So Far": orders_by_id.get(did, 0),
            "Cancelled So Far": cancellations_by_id.get(did, 0),
            "Days Worked": val["valid_days"] if val else attendance_by_id.get(did, 0),
            "Validity": performance_validity_by_id.get(did, ""),
        })
    detail_df = pd.DataFrame(rows).sort_values("Name").reset_index(drop=True) if rows else pd.DataFrame()

    roster_orders_riders = sum(1 for rec in roster_records.values() if rec.get("roster_total_orders") is not None)
    roster_orders_sum = sum(rec.get("roster_total_orders") or 0 for rec in roster_records.values())

    return {
        "roster_count": len(roster_records),
        "active_count": active_count,
        "terminated_count": terminated_count,
        "suspended_count": suspended_count,
        "company_cars": company_cars,
        "own_cars": own_cars,
        "riders_with_orders": len(orders_by_id),
        "total_orders_so_far": sum(orders_by_id.values()),
        "total_cancelled_so_far": sum(cancellations_by_id.values()),
        "valid_count": valid_count,
        "invalid_count": invalid_count,
        "attendance_riders": len(attendance_by_id),
        "detail_df": detail_df,
        "sheet_report": sheet_report,
        "roster_sheet_used": roster_sheet_used,
        "roster_sheets_seen": roster_sheets_seen,
        "roster_override_requested_but_not_found": roster_override_requested_but_not_found,
        "roster_orders_riders": roster_orders_riders,
        "roster_orders_sum": roster_orders_sum,
        "roster_sheet_columns": roster_sheet_columns,
        "orders_sheet_columns": orders_sheet_columns,
        "orders_sheet_diagnostics": orders_sheet_diagnostics,
        "duplicate_orders_sheets": duplicate_orders_sheets,
        "fetched_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


def maybe_auto_sync_gsheet() -> None:
    """If the Admin has turned on auto-sync, checkpoint the current
    month once per day -- and, the first time this runs after the
    calendar rolls into a new month, do one FINAL sync of the month
    that just ended first, so nothing entered on its last day is
    missed. This is what makes month-end saving happen without anyone
    needing to click "Sync Now": it simply runs automatically the next
    time the Admin opens the dashboard. Silently does nothing if it's
    not configured, not the Admin, or the Sheet can't be reached right
    now -- a Google API hiccup should never block the dashboard from
    loading."""
    if not is_admin() or not GSPREAD_AVAILABLE:
        return
    config = _load_gsheet_config()
    if not config or not config.get("sheet_id") or not config.get("auto_sync"):
        return

    today_str = datetime.today().strftime("%Y-%m-%d")
    if config.get("last_auto_sync_date") == today_str:
        return  # already checked today -- avoid hitting the Sheet on every rerun

    current_month = datetime.today().strftime("%Y-%m")
    last_synced_month = config.get("last_synced_month")

    try:
        if last_synced_month and last_synced_month != current_month:
            sync_month_from_gsheet(config["sheet_id"], last_synced_month, roster_tab_override=config.get("roster_tab_override"))
        sync_month_from_gsheet(config["sheet_id"], current_month, roster_tab_override=config.get("roster_tab_override"))
        _update_gsheet_config(last_auto_sync_date=today_str, last_synced_month=current_month)
        st.toast(f"\U0001F4E1 Auto-synced {month_display(current_month)} from Google Sheet", icon="\u2705")
    except Exception:  # noqa: BLE001
        pass  # will simply retry next time the Admin opens the app


def render_live_tracker_tab():
    st.subheader("\U0001F4E1 Live Tracker (Google Sheets)")
    st.write(
        "A real-time read straight from your connected Google Sheet -- separate from "
        "the Operations Dashboard, which only ever shows data that's been **saved** "
        "into permanent history. Use this tab to see today's numbers as supervisors "
        "enter them, and to save the current month whenever you're ready."
    )
    config = _load_gsheet_config()
    if not config or not config.get("sheet_id"):
        st.info(
            "No Google Sheet connected yet. Go to **Upload Monthly Data \u2192 Google "
            "Sheet Sync** to connect one."
        )
        return
    if not GSPREAD_AVAILABLE:
        st.error("The 'gspread' and 'google-auth' packages aren't installed on this server.")
        return
    render_live_month_panel()


def render_live_month_panel():
    """The full 'This Month So Far' picture, sourced directly from the
    connected Google Sheet -- not from the database. Nothing here is
    saved; it's purely a live read so the team can see today's numbers
    without waiting for month-end or a manual sync. Mirrors everything
    the real Sync would capture: headcount by status, vehicle types,
    orders, cancellations, validity, attendance, and a full per-rider
    table."""
    config = _load_gsheet_config()
    if not config or not config.get("sheet_id"):
        return
    if not GSPREAD_AVAILABLE:
        return

    with st.expander("\U0001F4E1 Live This Month (from connected Google Sheet)", expanded=True):
        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            if st.button("\U0001F504 Refresh live numbers", key="refresh_live_month", use_container_width=True):
                st.session_state.pop("live_month_cache", None)
        with btn_col2:
            save_now = st.button(
                "\U0001F4BE Save This Month to History Now", key="live_panel_save_now",
                type="primary", use_container_width=True,
                help="Saves exactly what's shown below into permanent history for this month -- same as Sync Now.",
            )

        if "live_month_cache" not in st.session_state:
            try:
                with st.spinner("Reading every tab of the Google Sheet..."):
                    st.session_state["live_month_cache"] = fetch_live_month_to_date(config["sheet_id"], roster_tab_override=config.get("roster_tab_override"))
            except Exception as exc:  # noqa: BLE001
                st.warning(f"Couldn't read the Google Sheet right now: {exc}")
                return

        if save_now:
            current_month = datetime.today().strftime("%Y-%m")
            try:
                with st.spinner(f"Saving {month_display(current_month)} into history..."):
                    summary = sync_month_from_gsheet(config["sheet_id"], current_month, roster_tab_override=config.get("roster_tab_override"))
            except Exception as exc:  # noqa: BLE001
                st.error(f"Save failed: {exc}")
            else:
                st.success(
                    f"Saved {month_display(current_month)} -- {summary['roster_count']} roster "
                    f"record(s), {summary['logs_written']} monthly log row(s) written."
                )
                _update_gsheet_config(
                    last_auto_sync_date=datetime.today().strftime("%Y-%m-%d"),
                    last_synced_month=current_month,
                )
                st.session_state.pop("live_month_cache", None)
                st.rerun()

        live = st.session_state["live_month_cache"]
        st.caption(f"As of {live['fetched_at']} -- not yet saved to history until you click Save.")

        stat_cards([
            {"icon": "\U0001F465", "label": "Riders on Sheet", "value": live["roster_count"],
             "tip": "From the roster tab, right now", "variant": "a"},
            {"icon": "\U0001F7E2", "label": "Active", "value": live["active_count"],
             "tip": "Not terminated/suspended, shows activity so far", "variant": "a"},
            {"icon": "\U0001F6D1", "label": "Terminated", "value": live["terminated_count"],
             "tip": "Marked Terminated on the roster tab", "variant": "c"},
            {"icon": "\u23F8\uFE0F", "label": "Suspended", "value": live["suspended_count"],
             "tip": "Temporarily suspended / on leave", "variant": "d"},
        ])
        _targets = _load_validity_targets()
        stat_cards([
            {"icon": "\U0001F697", "label": "Company Cars", "value": live["company_cars"],
             "tip": "Riders using a company-provided vehicle", "variant": "b"},
            {"icon": "\U0001F699", "label": "Own Cars", "value": live["own_cars"],
             "tip": "Riders using their own vehicle", "variant": "b"},
            {"icon": "\u2705", "label": "Valid (so far)", "value": live["valid_count"],
             "tip": f"\u2265{_targets['min_orders']} orders AND \u2265{_targets['min_days']} days worked so far", "variant": "a"},
            {"icon": "\u274C", "label": "Invalid (so far)", "value": live["invalid_count"],
             "tip": f"Below the \u2265{_targets['min_orders']}-order or \u2265{_targets['min_days']}-day target so far", "variant": "c"},
        ])
        stat_cards([
            {"icon": "\U0001F4E6", "label": "Orders So Far", "value": f"{live['total_orders_so_far']:,}",
             "tip": "Sum of every day filled in so far this month", "variant": "a"},
            {"icon": "\U0001F6AB", "label": "Cancelled So Far", "value": f"{live['total_cancelled_so_far']:,}",
             "tip": "From the cancellation tab, matched by name", "variant": "c"},
            {"icon": "\U0001F4C5", "label": "Attendance Riders", "value": live["attendance_riders"],
             "tip": "Riders with attendance data so far", "variant": "b"},
            {"icon": "\U0001F4CB", "label": "Riders w/ Orders", "value": live["riders_with_orders"],
             "tip": "Distinct riders with at least one order logged", "variant": "d"},
        ])

        st.markdown("---")
        st.markdown("##### Every rider, live from the Sheet")
        if live["detail_df"].empty:
            st.caption("No rider-level data found yet.")
        else:
            st.dataframe(live["detail_df"], use_container_width=True, hide_index=True)

        with st.expander("\U0001F50D Tab-by-tab breakdown (use this if a number looks wrong)"):
            st.caption(
                "Shows exactly how each tab in your Sheet was read and classified. "
                "If 'Riders on Sheet' doesn't match your actual roster count, check "
                "which tab is listed as **Roster sheet used** below -- if it's the "
                "wrong tab (or the row count for it looks too low), that's the tab "
                "to fix (e.g. rename it more clearly, remove a stray near-duplicate "
                "tab, or check that its ID/Name columns have proper headers)."
            )
            if live.get("roster_sheet_used"):
                st.markdown(f"**Roster sheet used:** `{live['roster_sheet_used']}`")
            if len(live.get("roster_sheets_seen", [])) > 1:
                st.warning(
                    "\u26A0\uFE0F More than one tab looked like a roster: "
                    + ", ".join(f"`{n}`" for n in live["roster_sheets_seen"])
                    + " -- only the best match above was used. If the wrong one was "
                    "picked, that's very likely why the count is off."
                )
            if live.get("roster_orders_riders", 0) > 0:
                st.success(
                    f"\u2705 The roster sheet's own **Total Orders** column WAS found and "
                    f"used -- {live['roster_orders_riders']} rider(s), summing to "
                    f"**{live['roster_orders_sum']:,}**. This is what 'Orders So Far' above "
                    f"is based on."
                )
            else:
                st.info(
                    "\u2139\uFE0F No **Total Orders**-style column was found on the roster "
                    "sheet itself, so 'Orders So Far' above comes entirely from a separate "
                    "orders/day-by-day tab instead. If your roster sheet DOES have such a "
                    "column, check its exact header text -- it needs to contain words like "
                    "'orders', 'total orders', or 'trips'/'deliveries' for it to be detected."
                )
                if live.get("roster_sheet_columns"):
                    st.caption("Column headers Streamlit actually read from the roster sheet:")
                    st.code(" | ".join(str(c) for c in live["roster_sheet_columns"]), language=None)

            if live.get("orders_sheet_diagnostics"):
                st.markdown("###### Orders tab(s) -- explicit column vs. day-by-day sum")
                for sheet_name, diag in live["orders_sheet_diagnostics"].items():
                    st.markdown(f"**`{sheet_name}`** -- currently using: *{diag['used']}*")
                    d1, d2 = st.columns(2)
                    with d1:
                        if diag["orders_col_found"]:
                            st.metric(f"Own '{diag['orders_col_name']}' column sum", f"{diag['orders_col_sum']:,}")
                        else:
                            st.caption("No explicit Total Orders-style column found on this tab.")
                    with d2:
                        st.metric(f"Sum of {diag['day_cols_count']} day-by-day column(s)", f"{diag['day_cols_sum']:,}")
                    if diag["orders_col_found"] and diag["day_cols_count"] > 0 and diag["orders_col_sum"] != diag["day_cols_sum"]:
                        st.caption(
                            "\u2139\uFE0F These two numbers differ -- since an explicit column "
                            "was found, it's the one being used (day-by-day sum shown only "
                            "for comparison)."
                        )
                    if diag.get("rows_without_id", 0) > 0:
                        st.warning(
                            f"\u26A0\uFE0F **{diag['rows_without_id']} row(s)** on this tab had no "
                            f"usable Driver/Courier ID and were skipped entirely -- their orders "
                            f"contributed **nothing** to any total. If your sheet's real total is "
                            f"higher than what's shown here, this is very likely why: check the ID "
                            f"column on those specific rows for blanks or a format our importer "
                            f"can't read."
                        )
                    if diag.get("colliding_days"):
                        days_str = ", ".join(f"Day {d}" for d in diag["colliding_days"])
                        st.error(
                            f"\u26A0\uFE0F **Two or more columns on this tab both resolve to the "
                            f"SAME calendar day** ({days_str}) -- every rider's orders for that day "
                            f"are being counted once per matching column, inflating just that one "
                            f"day far above its neighbors. Check the full column list below for a "
                            f"stray duplicate, a weekly-total column, or a typo'd header that "
                            f"accidentally looks like a date."
                        )
                    st.caption(
                        f"Rows counted: {diag.get('rows_with_id', 0)} "
                        f"(skipped for missing ID: {diag.get('rows_without_id', 0)})"
                    )
                    st.caption("Column headers on this tab:")
                    st.code(" | ".join(str(c) for c in live["orders_sheet_columns"].get(sheet_name, [])), language=None)
                    if diag.get("day_column_map"):
                        st.caption("Every day-column header \u2192 the calendar day it resolved to:")
                        st.code(
                            "\n".join(
                                f"{header!r:>30}  ->  Day {dn}" if dn is not None else f"{header!r:>30}  ->  (not resolved)"
                                for header, dn in diag["day_column_map"]
                            ),
                            language=None,
                        )

            if live.get("duplicate_orders_sheets"):
                st.warning(
                    "\u26A0\uFE0F These orders tab(s) looked like a near-duplicate of another "
                    "one already counted (same riders, matching numbers) and were **excluded** "
                    "to avoid double-counting: " + ", ".join(f"`{n}`" for n in live["duplicate_orders_sheets"])
                )


            report_df = pd.DataFrame(live["sheet_report"], columns=["Tab", "Detected as", "Rows"])
            st.dataframe(report_df, use_container_width=True, hide_index=True)


def render_gsheet_sync_tab():
    st.subheader("\U0001F517 Google Sheet Sync")
    st.write(
        "Connect the same Google Sheet your supervisors update every day. Once "
        "connected: the Operations Dashboard can show a **live, read-only "
        "month-to-date view** straight from the sheet, and you can **Sync Now** "
        "at any point (weekly, or at month-end) to save that month permanently "
        "into this dashboard's history -- exactly like uploading an Excel file, "
        "just pulled directly from Google instead."
    )

    if not GSPREAD_AVAILABLE:
        st.error(
            "The required packages aren't installed on this server. Ask whoever "
            "manages the deployment to run:\n\n`pip install gspread google-auth`"
        )
        return

    with st.expander("\u2699\uFE0F One-time setup (do this before connecting)", expanded="gcp_service_account" not in st.secrets):
        st.markdown(
            """
1. **Create a Google Cloud service account** (free): in the
   [Google Cloud Console](https://console.cloud.google.com/), create a
   project (or use an existing one) → **APIs & Services → Credentials**
   → **Create Credentials → Service Account**.
2. Enable the **Google Sheets API** and **Google Drive API** for that
   project (**APIs & Services → Library**, search each, click Enable).
3. Open the service account you created → **Keys** tab → **Add Key →
   Create new key → JSON**. This downloads a `.json` file -- keep it
   private, never commit it anywhere public.
4. In that JSON file, find the `"client_email"` value (looks like
   `something@your-project.iam.gserviceaccount.com`). **Share your
   Google Sheet with this email address** (Share button in Google
   Sheets, give it Viewer access) -- exactly like sharing with a
   colleague.
5. Add the JSON file's contents to this app's **Streamlit secrets**
   (`.streamlit/secrets.toml` if running locally, or the "Secrets"
   panel if deployed on Streamlit Community Cloud) under a
   `[gcp_service_account]` section, e.g.:

```toml
[gcp_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
client_email = "something@your-project.iam.gserviceaccount.com"
client_id = "..."
token_uri = "https://oauth2.googleapis.com/token"
```

   (Copy each field straight from the downloaded JSON -- just wrap the
   whole thing under `[gcp_service_account]` as shown.)
6. Restart the app after saving secrets.
            """
        )

    config = _load_gsheet_config()
    current_id = config["sheet_id"] if config else ""

    sheet_input = st.text_input(
        "Google Sheet URL or ID",
        value=current_id,
        placeholder="https://docs.google.com/spreadsheets/d/XXXXXXXX/edit",
        help="Paste the full sheet link or just the ID -- either works.",
    )
    if st.button("\U0001F4BE Save Connection", use_container_width=True):
        if not sheet_input.strip():
            st.error("Please paste a Google Sheet URL or ID.")
        else:
            _save_gsheet_config(_extract_sheet_id(sheet_input))
            st.success("Google Sheet connection saved.")
            st.rerun()

    if not config or not config.get("sheet_id"):
        st.info("No Google Sheet connected yet -- paste a link above to get started.")
        return

    st.success(f"\u2705 Connected to Sheet ID: `{config['sheet_id']}`")

    if st.button("\U0001F441\uFE0F Test Connection", use_container_width=True):
        try:
            with st.spinner("Connecting..."):
                frames = fetch_gsheet_frames(config["sheet_id"])
            st.success(f"Connected successfully -- found {len(frames)} tab(s): " + ", ".join(n for n, _ in frames))
        except Exception as exc:  # noqa: BLE001
            st.error(f"Couldn't connect: {exc}")

    st.markdown("---")
    st.markdown("#### \U0001F3AF Roster Tab Override")
    st.caption(
        "If your Sheet has more than one tab that looks like a roster (e.g. "
        "'Main Data' and 'Registered ID'), auto-detection picks whichever "
        "looks most complete -- which can be the wrong one. Type the EXACT "
        "tab name here to force that tab to always be used as the roster. "
        "Leave blank to auto-detect."
    )
    roster_override_input = st.text_input(
        "Roster tab name (exact, case doesn't matter)",
        value=config.get("roster_tab_override") or "",
        placeholder="e.g. Main Data",
        key="roster_tab_override_input",
    )
    if st.button("\U0001F4BE Save Roster Tab Choice", use_container_width=True):
        _update_gsheet_config(roster_tab_override=roster_override_input.strip() or None)
        st.session_state.pop("live_month_cache", None)
        st.success("Saved.")
        st.rerun()

    st.markdown("---")
    st.markdown("#### \U0001F6AB Tabs to Ignore")
    st.caption(
        "List any tab(s) that should be completely skipped -- never counted "
        "toward orders, validity, headcount, or anything else -- even if they'd "
        "otherwise be auto-detected as an orders/validity/roster sheet. Useful "
        "for a tab covering a different order type or category you don't want "
        "mixed into the main totals (e.g. 'PT ORDERS REPOER'). One tab name "
        "per line, exact spelling (case doesn't matter)."
    )
    ignored_tabs_input = st.text_area(
        "Tab names to ignore",
        value="\n".join(config.get("ignored_tabs") or []),
        placeholder="e.g.\nPT ORDERS REPOER\nScheduling",
        key="ignored_tabs_input",
        height=100,
    )
    if st.button("\U0001F4BE Save Ignored Tabs", use_container_width=True):
        tabs_list = [line.strip() for line in ignored_tabs_input.splitlines() if line.strip()]
        _update_gsheet_config(ignored_tabs=tabs_list)
        st.session_state.pop("live_month_cache", None)
        st.success(f"Saved -- {len(tabs_list)} tab(s) will now be skipped entirely." if tabs_list else "Saved -- no tabs ignored.")
        st.rerun()

    st.markdown("---")
    st.markdown("#### \U0001F4E5 Sync a Month into History")
    st.caption(
        "Pulls the sheet right now and saves it as one month's record -- exactly "
        "like the Excel auto-import. Safe to run again for the same month; it "
        "updates numbers in place rather than duplicating them. Run this at "
        "month-end (or any time you want to checkpoint the current numbers)."
    )
    inferred_month = datetime.today().strftime("%Y-%m")
    sync_month = st.text_input("Month to sync (YYYY-MM)", value=inferred_month, key="gsheet_sync_month")

    if st.button("\U0001F680 Sync Now", type="primary"):
        if not sync_month.strip():
            st.error("Please provide a month (YYYY-MM).")
        else:
            try:
                with st.spinner("Reading the Google Sheet and importing..."):
                    summary = sync_month_from_gsheet(config["sheet_id"], sync_month.strip(), roster_tab_override=config.get("roster_tab_override"))
            except Exception as exc:  # noqa: BLE001
                st.error(f"Sync failed: {exc}")
            else:
                st.success(
                    f"Imported {summary['roster_count']} roster record(s) and wrote "
                    f"{summary['logs_written']} monthly log row(s) for {sync_month.strip()}."
                )
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Orders found for", f"{summary['orders_drivers']} riders")
                m2.metric("Total Orders", f"{summary['orders_total']:,}")
                m3.metric("Validity data for", f"{summary['validity_drivers']} riders")
                m4.metric("Attendance data for", f"{summary['attendance_drivers']} riders")
                if summary["placeholders_created"]:
                    st.warning(
                        f"\u26A0\uFE0F {summary['placeholders_created']} rider(s) appeared in "
                        f"orders/validity/attendance but weren't in the roster tab -- "
                        f"placeholder profiles were created. Find them in **Rider Lookup** "
                        f"(search 'Unknown Rider')."
                    )
                st.markdown("##### What each tab was used for")
                report_df = pd.DataFrame(summary["sheet_report"], columns=["Tab", "Detected as", "Rows"])
                st.dataframe(report_df, use_container_width=True, hide_index=True)
                st.session_state.pop("live_month_cache", None)
                st.rerun()

    st.markdown("---")
    st.markdown("#### \U0001F916 Automatic Month-End Saving")
    st.caption(
        "Turn this on and the dashboard checkpoints the current month "
        "automatically the next time you (Admin) open it -- once a day is "
        "enough. The moment the calendar rolls into a new month, it also "
        "does one FINAL sync of the month that just ended first, so nothing "
        "entered on the last day gets missed. No button-clicking needed."
    )
    auto_sync_on = st.checkbox(
        "\U0001F501 Auto-save automatically when I open the dashboard",
        value=bool(config.get("auto_sync", False)),
        key="gsheet_auto_sync_toggle",
    )
    if auto_sync_on != bool(config.get("auto_sync", False)):
        _update_gsheet_config(auto_sync=auto_sync_on)
        st.rerun()

    if config.get("last_auto_sync_date"):
        st.caption(
            f"Last auto-sync: **{config['last_auto_sync_date']}** "
            f"(month checkpointed: {month_display(config.get('last_synced_month', ''))})."
        )

    st.markdown("---")
    st.markdown("#### \U0001F5A5\uFE0F Fully Unattended (even if nobody ever opens the dashboard)")
    st.caption(
        "The toggle above still needs an Admin to open the dashboard at least "
        "once a day. If you want saving to happen even if NOBODY ever opens "
        "it, use the separate `gsheet_sync.py` script (provided alongside "
        "this app) with an OS-level scheduler instead:"
    )
    st.code(
        "# Runs the same sync used above, but from the command line --\n"
        "# schedule this however your server supports:\n"
        "#   Linux/macOS (cron), daily at 11:59 PM:\n"
        "#     59 23 * * * cd /path/to/app && python3 gsheet_sync.py\n"
        "#   Windows: use Task Scheduler to run the same command daily.\n"
        "python3 gsheet_sync.py",
        language="bash",
    )
    st.caption(
        "Running it daily (not just on the last day) means even if a server "
        "restart or a missed run happens right at month-end, you're never more "
        "than a day out of date -- each run just re-syncs the current month in "
        "place."
    )


def render_upload_tab():
    if not is_admin():
        st.info(
            "\U0001F512 Only the HQ Admin can upload or import data. "
            "You're viewing this dashboard as a read-only Viewer."
        )
        return

    tab_ops, tab_salary, tab_gsheet = st.tabs(
        ["\U0001F4E4 Roster / Orders / Validity", "\U0001F4B0 Salary Data", "\U0001F517 Google Sheet Sync"]
    )
    with tab_ops:
        _render_operations_upload()
    with tab_salary:
        _render_salary_upload()
    with tab_gsheet:
        render_gsheet_sync_tab()


def _render_operations_upload():
    st.subheader("\U0001F4E4 Upload Monthly Data")
    st.write(
        "Drag & drop **any** Excel/CSV export of your driver roster and/or monthly "
        "payroll -- real files rarely match a fixed template, so you'll map your "
        "actual columns to our fields below before importing. Existing drivers/months "
        "are updated in place, so re-uploading the same file is always safe."
    )

    st.download_button(
        "\u2B07\uFE0F Download a Reference CSV Template",
        data=_build_template_csv(),
        file_name="monthly_upload_template.csv",
        mime="text/csv",
        help="Not required -- just shows the field names we understand.",
    )

    uploaded_file = st.file_uploader(
        "Choose a CSV or Excel file", type=["csv", "xlsx", "xls"]
    )

    if uploaded_file is None:
        return

    is_excel = uploaded_file.name.lower().endswith((".xlsx", ".xls"))

    if is_excel:
        xls = pd.ExcelFile(uploaded_file)
        if len(xls.sheet_names) > 1:
            st.success(
                f"\U0001F4C4 This workbook has **{len(xls.sheet_names)} sheets**: "
                f"{', '.join(xls.sheet_names)}"
            )
            st.markdown("### \U0001F680 Auto-Import Everything")
            st.write(
                "Every sheet is inspected automatically -- roster, orders, validity, "
                "attendance, and cancellation sheets are all recognized by their columns "
                "and merged into one record per rider, regardless of how many sheets there are. "
                "Sheets that don't match a known pattern (schedules, accident logs, etc.) are "
                "safely skipped and listed below so nothing is silently lost. If two sheets "
                "contain the same orders data under different names, only one is counted. "
                "Section labels inside a roster sheet (e.g. 'TERMINATE FOR THIS MONTH', "
                "'ON VOCATION') are also recognized -- every rider listed under one is "
                "tagged with that status automatically."
            )
            inferred_month = _infer_month_from_filename(uploaded_file.name) or datetime.today().strftime("%Y-%m")
            month_year = st.text_input(
                "Which month does this whole file cover? (YYYY-MM)",
                value=inferred_month,
                key="all_sheets_month",
            )
            existing_months = distinct_months()
            if month_year.strip() in existing_months:
                existing_count = len(load_monthly_logs()[load_monthly_logs()["month_year"] == month_year.strip()])
                st.info(
                    f"\u2139\uFE0F **{existing_count} log entries** already exist for "
                    f"**{month_year.strip()}**. Importing will update matching drivers' rows "
                    f"for that month rather than duplicate them."
                )

            if st.button("\U0001F680 Auto-Import All Sheets", type="primary"):
                if not month_year.strip():
                    st.error("Please provide a month (YYYY-MM).")
                else:
                    with st.spinner("Reading and classifying every sheet..."):
                        summary = process_workbook_all_sheets(uploaded_file, month_year.strip())

                    st.success(
                        f"Imported {summary['roster_count']} roster record(s) and wrote "
                        f"{summary['logs_written']} monthly log row(s) for {month_year.strip()}."
                    )
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Orders found for", f"{summary['orders_drivers']} riders")
                    m2.metric("Total Orders", f"{summary['orders_total']:,}")
                    m3.metric("Validity data for", f"{summary['validity_drivers']} riders")
                    m4.metric("Attendance data for", f"{summary['attendance_drivers']} riders")
                    if summary["cancellations_matched"] or summary["cancellations_unmatched"]:
                        st.caption(
                            f"Cancellations matched to {summary['cancellations_matched']} riders by name "
                            f"({summary['cancellations_unmatched']} names in the cancellation log couldn't "
                            f"be matched to a roster rider and were skipped)."
                        )
                    if summary["roster_sheets_skipped"]:
                        extra_note = (
                            f" {summary['extra_riders_from_skipped']} rider(s) unique to the skipped "
                            f"sheet(s) were still merged in."
                            if summary["extra_riders_from_skipped"] else " No riders were lost."
                        )
                        st.info(
                            f"\u2139\uFE0F This workbook had more than one roster-shaped sheet -- used "
                            f"**{summary['roster_sheet_used']}** as the main roster (most complete match) "
                            f"and treated {', '.join(summary['roster_sheets_skipped'])} as a duplicate/backup "
                            f"copy instead of merging it in wholesale (that used to inflate headcount)."
                            + extra_note
                        )
                    if summary["placeholders_created"]:
                        st.warning(
                            f"\u26A0\uFE0F {summary['placeholders_created']} rider(s) appeared in the orders/"
                            f"validity/attendance/cancellation sheets but weren't found in the roster sheet -- "
                            f"placeholder profiles were created for them so their payroll data wasn't lost. "
                            f"You can find and fill in their real names/details in **Rider Lookup** "
                            f"(search 'Unknown Rider')."
                        )
                    if summary["duplicate_orders_sheets"]:
                        st.warning(
                            "\u26A0\uFE0F These sheet(s) looked like a duplicate of another orders sheet "
                            "already counted (same riders, same order numbers), so they were **skipped** "
                            "to avoid doubling totals: " + ", ".join(summary["duplicate_orders_sheets"])
                        )

                    st.markdown("##### What each sheet was used for")
                    report_df = pd.DataFrame(
                        summary["sheet_report"], columns=["Sheet", "Detected as", "Rows"]
                    )
                    st.dataframe(report_df, use_container_width=True, hide_index=True)
                    st.rerun()

            st.markdown("---")
            with st.expander("\U0001F527 Advanced: manually map a single sheet instead"):
                sheet_name = st.selectbox("Sheet to import", xls.sheet_names, index=0, key="manual_sheet_pick")
                _render_single_sheet_upload(uploaded_file, sheet_name)
            return

        sheet_name = xls.sheet_names[0]
        _render_single_sheet_upload(uploaded_file, sheet_name)
        return

    _render_single_sheet_upload(uploaded_file, None)


def _render_salary_upload():
    st.subheader("\U0001F4B0 Upload Salary Data")
    st.write(
        "Separate from the roster/orders/validity upload -- use this for "
        "your monthly salary workbook (Courier ID, Total payable amount, "
        "Total Deduction, Final Salary, Pending, plus the company-level "
        "summary sheet with Tax Amount). This fills in the Gross Salary, "
        "Deductions, Net Salary, and Pending Salary figures shown on the "
        "**Financial & Payroll** tab -- existing riders/months are "
        "updated in place, so re-uploading the same file is always safe."
    )
    st.caption(
        "Every sheet is inspected automatically: the per-rider sheet "
        "(however it's named -- 'JAN SALARY', 'riderDetail', etc.) and "
        "the one-row company summary sheet are both recognized by their "
        "columns, not their sheet name."
    )

    salary_file = st.file_uploader(
        "Choose the salary Excel file", type=["xlsx", "xls"], key="salary_uploader"
    )
    if salary_file is None:
        return

    inferred_month = _infer_month_from_filename(salary_file.name) or datetime.today().strftime("%Y-%m")
    salary_month = st.text_input(
        "Which month does this file cover, if a sheet doesn't already say? (YYYY-MM)",
        value=inferred_month,
        key="salary_month_input",
        help="Used only as a fallback -- if a sheet has its own 'Billing Cycle' "
             "column (e.g. 'Feb 2026'), that value wins for each row.",
    )
    existing_months = distinct_months()
    if salary_month.strip() in existing_months:
        st.info(
            f"\u2139\uFE0F There's already payroll data on file for "
            f"**{salary_month.strip()}**. Importing will update matching "
            f"riders' salary figures for that month rather than duplicate them."
        )

    if st.button("\U0001F4B0 Import Salary Data", type="primary"):
        if not salary_month.strip():
            st.error("Please provide a fallback month (YYYY-MM).")
        else:
            with st.spinner("Reading and matching salary data..."):
                result = process_salary_workbook(salary_file, salary_month.strip())

            st.success(f"Updated salary figures for {result['riders_updated']} rider(s).")
            if result["summary_written"]:
                st.caption(
                    f"Also recorded {result['summary_written']} company-level "
                    f"summary row(s) (total money received, tax amount)."
                )
            if result["placeholders_created"]:
                st.warning(
                    f"\u26A0\uFE0F {result['placeholders_created']} rider(s) in this file "
                    f"weren't found in your roster (by ID or name) -- placeholder "
                    f"profiles were created so their salary data wasn't lost. "
                    f"Find and fix them in **Rider Lookup** (search 'Unknown Rider')."
                )

            st.markdown("##### What each sheet was used for")
            report_df = pd.DataFrame(
                result["sheet_report"], columns=["Sheet", "Detected as", "Rows"]
            )
            st.dataframe(report_df, use_container_width=True, hide_index=True)
            st.rerun()


def _render_single_sheet_upload(uploaded_file, sheet_name):
    try:
        if sheet_name is not None:
            raw_df, header_row_used = _read_excel_smart(uploaded_file, sheet_name)
            if header_row_used == 2:
                st.caption(
                    "\u2139\uFE0F Detected that this sheet's real column headers are on "
                    "row 2 (row 1 looked like a title/section row) -- used row 2 automatically."
                )
        else:
            raw_df = pd.read_csv(uploaded_file)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not read file: {exc}")
        return

    raw_df = raw_df.dropna(axis=0, how="all").reset_index(drop=True)
    raw_df.columns = [str(c).strip() for c in raw_df.columns]

    st.markdown(f"**Preview** -- {len(raw_df):,} rows, {len(raw_df.columns)} columns detected")
    st.dataframe(raw_df.head(15), use_container_width=True, hide_index=True)

    file_columns = list(raw_df.columns)
    column_options = [NONE_OPTION] + file_columns

    st.markdown("---")
    st.markdown("#### \U0001F517 Map Your Columns")
    st.caption(
        "We've auto-guessed matches below where possible -- double-check them and "
        "adjust any dropdown that isn't right. Fields left as "
        f"**'{NONE_OPTION}'** are simply skipped."
    )

    mapping = {}
    with st.expander("\U0001F464 Driver / Roster Fields", expanded=True):
        c1, c2 = st.columns(2)
        for i, field in enumerate(["driver_id", "driver_name", "first_name", "last_name", "phone", "supervisor_name", "sponsor_name", "iqama_number"]):
            guess = _guess_column(file_columns, FIELD_ALIASES.get(field, []))
            target_col = c1 if i % 2 == 0 else c2
            mapping[field] = target_col.selectbox(
                FIELD_LABELS[field], column_options,
                index=column_options.index(guess) if guess in column_options else 0,
                key=f"map_{field}",
            )
        c3, c4 = st.columns(2)
        for i, field in enumerate(["vehicle_type", "status", "join_date", "termination_date"]):
            guess = _guess_column(file_columns, FIELD_ALIASES.get(field, []))
            target_col = c3 if i % 2 == 0 else c4
            mapping[field] = target_col.selectbox(
                FIELD_LABELS[field], column_options,
                index=column_options.index(guess) if guess in column_options else 0,
                key=f"map_{field}",
            )

    guessed_payroll = any(
        _guess_column(file_columns, FIELD_ALIASES[f]) != NONE_OPTION
        for f in ["total_orders", "cancelled_orders", "gross_salary", "total_deductions",
                  "pending_salary", "net_salary", "validity_status", "valid_days_in_month"]
    )
    include_payroll = st.checkbox(
        "\U0001F4CA This file also includes monthly payroll / performance figures",
        value=guessed_payroll,
    )

    default_month = ""
    if include_payroll:
        with st.expander("\U0001F4B0 Payroll Fields", expanded=True):
            c5, c6 = st.columns(2)
            for i, field in enumerate(PAYROLL_FIELDS):
                guess = _guess_column(file_columns, FIELD_ALIASES.get(field, []))
                target_col = c5 if i % 2 == 0 else c6
                mapping[field] = target_col.selectbox(
                    FIELD_LABELS[field], column_options,
                    index=column_options.index(guess) if guess in column_options else 0,
                    key=f"map_{field}",
                )
            if mapping.get("month_year", NONE_OPTION) == NONE_OPTION:
                inferred_month = _infer_month_from_filename(uploaded_file.name) or datetime.today().strftime("%Y-%m")
                st.warning(
                    "\u26A0\uFE0F **This file has no month column -- double-check the month below "
                    "before importing.** Every upload without a month column reuses whatever you "
                    "type here, so if you forget to update it for each new file, the new data will "
                    "silently overwrite the previous month instead of adding to it."
                )
                default_month = st.text_input(
                    "Payroll month this file applies to (YYYY-MM)",
                    value=inferred_month,
                )
                existing_months = distinct_months()
                if default_month.strip() in existing_months:
                    existing_count = len(load_monthly_logs()[load_monthly_logs()["month_year"] == default_month.strip()])
                    st.info(
                        f"\u2139\uFE0F There are already **{existing_count} log entries** stored for "
                        f"**{default_month.strip()}**. Importing now will update those rows in place "
                        f"(matching drivers) rather than create a duplicate month. If this file is for "
                        f"a *different* month, please correct the field above first."
                    )

    st.markdown("---")
    if st.button("\U0001F680 Process & Import into Database", type="primary"):
        if include_payroll and mapping.get("month_year", NONE_OPTION) == NONE_OPTION and not default_month.strip():
            st.error("Please provide a payroll month (YYYY-MM) since the file has no month column.")
            return

        with st.spinner("Importing rows into logistics.db..."):
            success_count, errors = _process_upload_mapped(raw_df, mapping, default_month.strip(), include_payroll)

        if success_count:
            st.success(f"Successfully imported/updated {success_count:,} row(s).")
        if errors:
            st.error(f"{len(errors)} row(s) failed validation and were skipped:")
            st.dataframe(pd.DataFrame({"Error": errors}), use_container_width=True, hide_index=True)

        if success_count:
            st.rerun()


# ==============================================================================
# TAB 4: RIDER LOOKUP  (animated per-driver profile + month-by-month history)
# ==============================================================================


def render_rider_drilldown_panel(state_key: str, drill_defs: dict) -> None:
    """Like render_drilldown_panel (Operations Dashboard), but with its
    own independent session_state key AND richer content than a plain
    table: a day-by-day orders chart+table, a day-by-day validity/
    attendance table, or a simple explanatory message for figures that
    have no daily breakdown at all (salary, deductions -- recorded
    once a month, not per day). drill_defs is
    {card_id: {"title", "kind", ...}} where kind is "daily_orders",
    "daily_status", or "message"."""
    active = st.session_state.get(state_key)
    if not active or active not in drill_defs:
        return
    spec = drill_defs[active]
    with st.container(key=f"drill_panel_box_{state_key}"):
        c1, c2 = st.columns([6, 1])
        c1.markdown(f"##### \U0001F50E {spec['title']}")
        if c2.button("\u2715 Close", key=f"close_drill_btn_{state_key}", use_container_width=True):
            st.session_state[state_key] = None
            st.rerun()
        if spec.get("note"):
            st.caption(spec["note"])

        kind = spec.get("kind")
        if kind == "daily_orders":
            df = spec["df"]
            if df.empty:
                st.caption("No day-by-day data on file for this month yet.")
            else:
                st.bar_chart(df.set_index("Date")["Orders"], height=180)
                st.dataframe(df, use_container_width=True, hide_index=True)
        elif kind == "daily_status":
            df = spec["df"]
            if df.empty:
                st.caption("No day-by-day data on file for this month yet.")
            else:
                st.dataframe(df, use_container_width=True, hide_index=True)
        elif kind == "message":
            st.info(spec.get("message", "No further detail available."))


def _rider_month_trend(rider_logs: pd.DataFrame, sel_month: str, column: str):
    """Compares sel_month's value in `column` to the most recent
    EARLIER month on file for this rider. Returns (delta, pct_change,
    previous_month) or None if there's no earlier month to compare
    against yet."""
    earlier = rider_logs[rider_logs["month_year"] < sel_month].sort_values("month_year")
    if earlier.empty:
        return None
    prev_row = earlier.iloc[-1]
    prev_val = float(prev_row[column] or 0)
    cur_rows = rider_logs[rider_logs["month_year"] == sel_month]
    if cur_rows.empty:
        return None
    cur_val = float(cur_rows[column].iloc[0] or 0)
    delta = cur_val - prev_val
    pct = (delta / prev_val * 100) if prev_val else (100.0 if cur_val > 0 else 0.0)
    return delta, pct, prev_row["month_year"]


def render_rider_performance_trend(rider_logs: pd.DataFrame, sel_month: str) -> None:
    """A small, honest 'is this rider trending up or down' strip,
    comparing sel_month to whichever earlier month is on file for
    them -- orders trending up is good, cancellations trending up is
    bad, so each is judged on its own terms rather than one combined
    score. Says nothing at all if there's no earlier month yet (first
    month on file has nothing to compare against)."""
    orders_trend = _rider_month_trend(rider_logs, sel_month, "total_orders")
    cancel_trend = _rider_month_trend(rider_logs, sel_month, "cancelled_orders")
    if not orders_trend and not cancel_trend:
        return

    parts = []
    if orders_trend:
        delta, pct, prev_month = orders_trend
        if delta > 0:
            parts.append(f"\U0001F4C8 Orders up {pct:.0f}% vs {month_display(prev_month)}")
        elif delta < 0:
            parts.append(f"\U0001F4C9 Orders down {abs(pct):.0f}% vs {month_display(prev_month)}")
        else:
            parts.append(f"\u27A1\uFE0F Orders unchanged vs {month_display(prev_month)}")
    if cancel_trend:
        delta, pct, prev_month = cancel_trend
        # Lower cancellations is the improvement here, so the arrow
        # direction is intentionally the OPPOSITE of the orders one.
        if delta < 0:
            parts.append(f"\U0001F4C9 Cancellations down {abs(pct):.0f}% vs {month_display(prev_month)}")
        elif delta > 0:
            parts.append(f"\U0001F4C8 Cancellations up {pct:.0f}% vs {month_display(prev_month)}")
        else:
            parts.append(f"\u27A1\uFE0F Cancellations unchanged vs {month_display(prev_month)}")

    st.markdown(
        f"""
        <div style="border-radius:12px; padding:10px 16px; margin:4px 0 14px 0;
                    background:linear-gradient(135deg, rgba(59,130,246,0.10), rgba(34,197,94,0.08));
                    border:1px solid rgba(120,120,120,0.18); font-size:13.5px;">
          {" &nbsp;\u2022&nbsp; ".join(parts)}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_month_reveal_cards(rows: pd.DataFrame, join_date: str, vehicle_type: str):
    css = """
    <style>
      .reveal-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
        gap: 14px;
        margin-top: 6px;
      }
      .reveal-card {
        border-radius: 14px;
        padding: 14px 16px;
        background: linear-gradient(160deg, rgba(34,197,94,0.08), rgba(59,130,246,0.06));
        border: 1px solid rgba(120,120,120,0.18);
        box-shadow: 0 4px 14px rgba(0,0,0,0.08);
        opacity: 0;
        animation: cardReveal 0.55s ease forwards;
        transition: transform 0.18s ease, box-shadow 0.18s ease;
      }
      .reveal-card:hover {
        transform: translateY(-6px) scale(1.02);
        box-shadow: 0 14px 28px rgba(0,0,0,0.18);
      }
      @keyframes cardReveal {
        from { opacity: 0; transform: translateY(16px) scale(0.97); }
        to   { opacity: 1; transform: translateY(0) scale(1); }
      }
      .reveal-month { font-weight: 800; font-size: 15px; margin-bottom: 6px; }
      .reveal-row { display: flex; justify-content: space-between; font-size: 13px; padding: 2px 0; }
      .reveal-row span { opacity: 0.7; }
      .reveal-badge {
        display: inline-block; margin-top: 8px; padding: 2px 10px;
        border-radius: 999px; font-size: 11px; font-weight: 700;
      }
      .reveal-badge.valid   { background: rgba(34,197,94,0.18); color: #16803c; }
      .reveal-badge.invalid { background: rgba(239,68,68,0.18); color: #b91c1c; }
    </style>
    <div class="reveal-grid">
    """
    cards = []
    for i, row in enumerate(rows.itertuples()):
        delay = round(i * 0.08, 2)
        badge_class = "valid" if row.validity_status == "Valid" else "invalid"
        vehicle_icon = "\U0001F697" if vehicle_type == "Company Car" else "\U0001F699"
        cards.append(
            f"""
            <div class="reveal-card" style="animation-delay:{delay}s">
              <div class="reveal-month">{month_display(row.month_year)}</div>
              <div class="reveal-row"><span>Orders</span><b>{row.total_orders}</b></div>
              <div class="reveal-row"><span>Cancelled</span><b>{row.cancelled_orders}</b></div>
              <div class="reveal-row"><span>Days Worked</span><b>{row.valid_days_in_month}</b></div>
              <div class="reveal-row"><span>Gross</span><b>SAR {row.gross_salary:,.0f}</b></div>
              <div class="reveal-row"><span>Deductions</span><b>SAR {row.total_deductions:,.0f}</b></div>
              <div class="reveal-row"><span>Pending</span><b>SAR {row.pending_salary:,.0f}</b></div>
              <div class="reveal-row"><span>Net</span><b>SAR {row.net_salary:,.0f}</b></div>
              <div class="reveal-row"><span>Joined</span><b>{join_date or 'N/A'}</b></div>
              <div class="reveal-row"><span>Vehicle</span><b>{vehicle_icon} {vehicle_type}</b></div>
              <div class="reveal-badge {badge_class}">{row.validity_status}</div>
            </div>
            """
        )
    st.markdown(css + "".join(cards) + "</div>", unsafe_allow_html=True)


def render_rider_lookup(filters: dict):
    st.subheader("\U0001F50D Rider Lookup")
    st.write("Search by name, ID, IQAMA, or phone -- matching riders update instantly below.")

    drivers = load_drivers()
    if drivers.empty:
        st.info("No drivers yet. Seed sample data or upload a file.")
        return

    search_query = st.text_input(
        "Search",
        placeholder="\U0001F50D  Type a name, ID, IQAMA, or phone number...",
        label_visibility="collapsed",
        key="rider_search_box",
    )

    searchable = (
        drivers["driver_name"].fillna("") + " " + drivers["driver_id"].fillna("") + " "
        + drivers["iqama_number"].fillna("") + " " + drivers["phone"].fillna("")
    ).str.lower()

    query = search_query.strip().lower()
    if query:
        mask = searchable.str.contains(query, na=False, regex=False)
        matched = drivers[mask].copy()
        matched["_rank"] = ~matched["driver_name"].str.lower().str.startswith(query)
        matched = matched.sort_values(["_rank", "driver_name"]).drop(columns="_rank")
    else:
        matched = drivers.sort_values("driver_name")

    if matched.empty:
        st.warning(f"No riders match '{search_query}'.")
        return

    if query and len(matched) == 1:
        driver_id = matched.iloc[0]["driver_id"]
        st.caption(f"1 rider matched -- showing {matched.iloc[0]['driver_name']} directly.")
    else:
        name_to_id = dict(zip(matched["driver_name"] + "  (" + matched["driver_id"] + ")", matched["driver_id"]))
        st.caption(f"{len(matched)} rider(s) matched")
        choice = st.selectbox("Select a rider", list(name_to_id.keys()))
        driver_id = name_to_id[choice]

    profile = drivers[drivers["driver_id"] == driver_id].iloc[0]
    logs = load_monthly_logs()
    rider_logs = logs[logs["driver_id"] == driver_id].sort_values("month_year")

    st.markdown(
        f"""
        <div style="border-radius:16px; padding:16px 20px; margin-bottom:10px;
                    background:linear-gradient(135deg, rgba(34,197,94,0.10), rgba(168,85,247,0.08));
                    border:1px solid rgba(120,120,120,0.18); animation: cardReveal 0.5s ease both;">
          <div style="font-size:20px; font-weight:800;">{profile['driver_name']}</div>
          <div style="opacity:0.75; font-size:13px; margin-top:4px;">
            {profile['status']} &nbsp;|&nbsp; {profile['vehicle_type']} &nbsp;|&nbsp;
            Supervisor: {profile['supervisor_name'] or 'N/A'} &nbsp;|&nbsp;
            Sponsor: {profile['sponsor_name'] or 'N/A'}
          </div>
          <div style="opacity:0.75; font-size:13px; margin-top:2px;">
            IQAMA: {profile['iqama_number'] or 'N/A'} &nbsp;|&nbsp;
            Joined: {profile['join_date'] or 'N/A'} &nbsp;|&nbsp;
            Ended: {profile['termination_date'] or 'Still active'}
          </div>
        </div>
        <style>
          @keyframes cardReveal {{
            from {{ opacity: 0; transform: translateY(10px); }}
            to   {{ opacity: 1; transform: translateY(0); }}
          }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    if rider_logs.empty:
        st.info("No monthly logs on file for this rider yet.")
        return

    # Scoped to ONE month at a time -- defaults to whatever month is
    # picked in the sidebar filter, falling back to this rider's most
    # recent month if they have no log for the sidebar's month. Every
    # figure below (orders, cancelled, salary, validity, days worked)
    # comes from that single month's row -- nothing here is summed
    # across months, so it can't be mistaken for an all-time total.
    available_months = sorted(rider_logs["month_year"].dropna().unique(), reverse=True)
    default_month = filters.get("month") if filters.get("month") in available_months else available_months[0]
    sel_month = st.selectbox(
        "Viewing month",
        available_months,
        index=available_months.index(default_month),
        format_func=month_display,
        key="rider_lookup_month",
    )

    month_rows = rider_logs[rider_logs["month_year"] == sel_month]
    st.markdown(f"#### {month_display(sel_month)}")
    if month_rows.empty:
        st.info(f"No log on file for this rider in {month_display(sel_month)}.")
    else:
        row = month_rows.iloc[0]
        render_rider_performance_trend(rider_logs, sel_month)
        st.caption("\U0001F446 Tap any card to see the day-by-day detail behind that number.")

        daily_df = load_daily_logs(driver_id, sel_month)
        if not daily_df.empty:
            daily_df = daily_df.copy()
            daily_df["Date"] = daily_df["day"].apply(lambda d: f"{sel_month}-{int(d):02d}")

        orders_daily_df = pd.DataFrame()
        status_daily_df = pd.DataFrame()
        if not daily_df.empty:
            od = daily_df[daily_df["orders"].notna()][["Date", "orders"]].rename(columns={"orders": "Orders"})
            orders_daily_df = od.sort_values("Date").reset_index(drop=True)
            sd = daily_df[daily_df["validity"].notna() | daily_df["attendance"].notna()][
                ["Date", "validity", "attendance"]
            ].rename(columns={"validity": "Validity", "attendance": "Attendance"})
            status_daily_df = sd.sort_values("Date").reset_index(drop=True)

        drill_defs = {
            "rider_row1_0": {
                "title": "Orders -- day by day", "kind": "daily_orders", "df": orders_daily_df,
                "note": "From the day-by-day orders sheet, when available." if not orders_daily_df.empty else None,
            },
            "rider_row1_1": {
                "title": "Cancelled Orders", "kind": "message",
                "message": (
                    f"{int(row['cancelled_orders'] or 0)} cancelled order(s) this month. Only a monthly "
                    f"total is tracked for cancellations -- day-by-day detail isn't available."
                ),
            },
            "rider_row1_2": {
                "title": "Gross Salary", "kind": "message",
                "message": f"SAR {(row['gross_salary'] or 0):,.0f} for {month_display(sel_month)}. Salary figures are recorded once per month, not per day.",
            },
            "rider_row1_3": {
                "title": "Pending Salary", "kind": "message",
                "message": f"SAR {(row['pending_salary'] or 0):,.0f} still owed for {month_display(sel_month)}.",
            },
            "rider_row2_0": {
                "title": "Deductions", "kind": "message",
                "message": f"SAR {(row['total_deductions'] or 0):,.0f} deducted for {month_display(sel_month)}.",
            },
            "rider_row2_1": {
                "title": "Net Salary", "kind": "message",
                "message": f"SAR {(row['net_salary'] or 0):,.0f} net for {month_display(sel_month)} (gross minus deductions).",
            },
            "rider_row2_2": {
                "title": "Days Worked / Attendance -- day by day", "kind": "daily_status", "df": status_daily_df,
                "note": "Validity and/or attendance, whichever this month's sheets recorded." if not status_daily_df.empty else None,
            },
            "rider_row2_3": {
                "title": "Validity -- day by day", "kind": "daily_status", "df": status_daily_df,
                "note": "Validity and/or attendance, whichever this month's sheets recorded." if not status_daily_df.empty else None,
            },
        }

        render_clickable_stat_row([
            {"icon": "\U0001F4E6", "label": "Orders", "value": f"{int(row['total_orders'] or 0):,}",
             "tip": "Completed orders this month", "variant": "a"},
            {"icon": "\u274C", "label": "Cancelled", "value": f"{int(row['cancelled_orders'] or 0):,}",
             "tip": "Cancelled orders this month", "variant": "c"},
            {"icon": "\U0001F4B5", "label": "Gross Salary", "value": f"SAR {(row['gross_salary'] or 0):,.0f}",
             "tip": "Before deductions, this month", "variant": "b"},
            {"icon": "\u23F3", "label": "Pending", "value": f"SAR {(row['pending_salary'] or 0):,.0f}",
             "tip": "Still owed for this month", "variant": "d"},
        ], row_key="rider_row1", state_key="rider_drill")
        if (st.session_state.get("rider_drill") or "").startswith("rider_row1_"):
            render_rider_drilldown_panel("rider_drill", drill_defs)

        render_clickable_stat_row([
            {"icon": "\u2796", "label": "Deductions", "value": f"SAR {(row['total_deductions'] or 0):,.0f}",
             "tip": "Deducted this month", "variant": "c"},
            {"icon": "\u2705", "label": "Net Salary", "value": f"SAR {(row['net_salary'] or 0):,.0f}",
             "tip": "Gross minus deductions, this month", "variant": "a"},
            {"icon": "\U0001F4C5", "label": "Days Worked", "value": int(row['valid_days_in_month'] or 0),
             "tip": "Attendance this month", "variant": "b"},
            {"icon": "\u2139\uFE0F", "label": "Validity", "value": row['validity_status'] or "N/A",
             "tip": "Validity status for this month", "variant": "d"},
        ], row_key="rider_row2", state_key="rider_drill")
        if (st.session_state.get("rider_drill") or "").startswith("rider_row2_"):
            render_rider_drilldown_panel("rider_drill", drill_defs)

    with st.expander("\U0001F4CA Lifetime Totals (all months combined)"):
        stat_cards([
            {"icon": "\U0001F4E6", "label": "Total Orders", "value": f"{int(rider_logs['total_orders'].sum()):,}",
             "tip": "Across every month on file", "variant": "a"},
            {"icon": "\u274C", "label": "Total Cancelled", "value": f"{int(rider_logs['cancelled_orders'].sum()):,}",
             "tip": "Cancelled orders, all-time", "variant": "c"},
            {"icon": "\U0001F4B5", "label": "Total Gross Earned", "value": f"SAR {rider_logs['gross_salary'].sum():,.0f}",
             "tip": "Before deductions, all-time", "variant": "b"},
            {"icon": "\u23F3", "label": "Total Pending", "value": f"SAR {rider_logs['pending_salary'].sum():,.0f}",
             "tip": "Still owed to this rider", "variant": "d"},
        ])
        stat_cards([
            {"icon": "\u2705", "label": "Months Valid", "value": int((rider_logs["validity_status"] == "Valid").sum()),
             "tip": "Months marked Valid", "variant": "a"},
            {"icon": "\u274C", "label": "Months Invalid", "value": int((rider_logs["validity_status"] == "Invalid").sum()),
             "tip": "Months marked Invalid", "variant": "c"},
            {"icon": "\U0001F4C5", "label": "Total Days Worked", "value": int(rider_logs["valid_days_in_month"].sum()),
             "tip": "Summed attendance across all months", "variant": "b"},
        ])

    st.markdown("---")
    st.markdown("#### Month-by-Month History")
    render_month_reveal_cards(rider_logs, profile["join_date"], profile["vehicle_type"])


# ==============================================================================
# TAB 5: SUPERVISOR ACTION ALERT GENERATOR
# ==============================================================================


def render_supervisor_alerts():
    st.subheader("\U0001F4E3 Supervisor Action Alert Generator")

    supervisors = distinct_supervisors()
    months = distinct_months()

    if not supervisors or not months:
        st.info("No data available yet. Seed sample data or upload a file.")
        return

    c1, c2, c3 = st.columns([2, 2, 2])
    supervisor = c1.selectbox("Supervisor", supervisors)
    month = c2.selectbox("Month", months, format_func=month_display)
    low_order_threshold = c3.slider("Low-order threshold", 20, 300, 80, step=10)

    merged = load_merged()
    team = merged[
        (merged["supervisor_name"] == supervisor)
        & (merged["month_year"] == month)
        & (merged["log_id"].notna())
    ]

    if team.empty:
        st.warning(f"No logs found for **{supervisor}** in **{month}**.")
        return

    total_riders = len(team)
    invalid_count = int((team["validity_status"] == "Invalid").sum())
    low_order_count = int((team["total_orders"] < low_order_threshold).sum())
    total_deductions = float(team["total_deductions"].sum())
    total_cancelled = int(team["cancelled_orders"].sum())
    total_orders = int(team["total_orders"].sum())
    avg_orders = total_orders / total_riders if total_riders else 0

    st.markdown("#### Team Snapshot")
    stat_cards([
        {"icon": "\U0001F465", "label": "Total Riders", "value": total_riders,
         "tip": "Riders on this supervisor's team this month", "variant": "a"},
        {"icon": "\u274C", "label": "Invalid Riders", "value": invalid_count,
         "tip": "Marked Invalid -- needs follow-up", "variant": "c"},
        {"icon": "\U0001F4C9", "label": "Low-Order Riders", "value": low_order_count,
         "tip": f"Below the {low_order_threshold}-order threshold", "variant": "d"},
        {"icon": "\U0001F6AB", "label": "Cancelled Orders", "value": total_cancelled,
         "tip": "Cancelled across the whole team", "variant": "c"},
        {"icon": "\U0001F4B8", "label": "Total Deductions", "value": f"SAR {total_deductions:,.0f}",
         "tip": "Summed across the team this month", "variant": "b"},
    ])

    line1 = (
        f"\u26A0\uFE0F *{supervisor} Team Alert \u2013 {month}*: {total_riders} riders tracked, "
        f"{invalid_count} invalid aur {low_order_count} riders {low_order_threshold} se kam orders par hain."
    )
    line2 = (
        f"\U0001F4B0 Total deductions is month SAR {total_deductions:,.0f} rahi hain, {total_cancelled} orders cancel huay, "
        f"average orders/rider = {avg_orders:.0f}. Kindly low performers ko review karein."
    )
    line3 = (
        "\U0001F527 Action: Invalid riders ko immediately valid karwayein, low-order riders ko coach karein, "
        "aur updated status EOD tak share karein. Shukriya!"
    )
    alert_text = f"{line1}\n{line2}\n{line3}"

    st.markdown("#### Generated Alert (copy-ready for WhatsApp / Teams)")
    st.code(alert_text, language=None)

    st.markdown("#### Underlying Team Data")
    st.dataframe(
        team[
            [
                "driver_name", "total_orders", "cancelled_orders", "total_deductions",
                "validity_status", "vehicle_type", "status",
            ]
        ].rename(
            columns={
                "driver_name": "Rider Name",
                "total_orders": "Orders",
                "cancelled_orders": "Cancelled",
                "total_deductions": "Deductions",
                "validity_status": "Validity",
                "vehicle_type": "Vehicle",
                "status": "Status",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )


# ==============================================================================
# MAIN APP ENTRY POINT
# ==============================================================================


def main():
    init_db()
    inject_custom_css()

    if not render_auth_gate():
        return

    render_hq_banner()
    maybe_auto_sync_gsheet()

    filters = render_sidebar()
    render_header(filters)

    st.markdown(
        "<h2 style='text-align:center; margin-top:0;'>Food Delivery Logistics, Operations & Payroll Dashboard</h2>"
        "<p style='text-align:center; opacity:0.7; margin-top:-8px;'>Driver operations, workforce KPIs, payroll, "
        "bulk uploads, rider lookup & supervisor alerts -- all in one place.</p>",
        unsafe_allow_html=True,
    )

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
        [
            "\U0001F4CA Operations Dashboard",
            "\U0001F4E1 Live Tracker (Google Sheets)",
            "\U0001F4B0 Financial & Payroll",
            "\U0001F4E4 Upload Monthly Data",
            "\U0001F50D Rider Lookup",
            "\U0001F4E3 Supervisor Alerts",
        ]
    )

    with tab1:
        render_dashboard(filters)
    with tab2:
        render_live_tracker_tab()
    with tab3:
        render_financials(filters)
    with tab4:
        render_upload_tab()
    with tab5:
        render_rider_lookup(filters)
    with tab6:
        render_supervisor_alerts()


if __name__ == "__main__":
    main()
