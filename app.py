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

# ==============================================================================
# CONFIG / CONSTANTS
# ==============================================================================
DB_PATH = "logistics.db"
OWNER_CONFIG_PATH = "hq_owner.json"

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


# ==============================================================================
# ANIMATED 3D BRAND LOGO (Tamkeen)
# ==============================================================================
# Drawn entirely as inline SVG + CSS (no embedded image file) so the whole
# app stays a small, easy-to-copy single .py file. A layered "3D extrusion"
# (several offset, darkened copies of the same shapes stacked behind the
# crisp full-color mark) is what actually reads as solid depth, combined
# with a bounded tilt oscillation -- a full 360-degree spin makes a flat
# SVG look like a broken sliver at 90/270 degrees, so we avoid that.


# ==============================================================================
# HQ ACCESS CONTROL  -- one Admin (you, email+password) + tracked Viewers
# ==============================================================================
# The first person to open the app sets up the one Admin account (email +
# password). That password is never stored in plain text -- only a salted
# hash, saved to a local file (hq_owner.json) next to the database.
#
# Everyone else who opens the app is asked for just a NAME and EMAIL (no
# password) to "Continue as Viewer" -- this is what lets you see exactly
# who has looked at the dashboard: every viewer who has ever identified
# themselves is logged with first-seen/last-seen times and whether they
# currently look "online" (active in the last few minutes). You can
# revoke ("Remove") any viewer's access at any time from the sidebar --
# a removed person is blocked from continuing as a viewer again until you
# restore them.
#
# HONESTY NOTE (also shown in the app): a name/email here is
# self-reported, not verified -- anyone could type someone else's name.
# This is lightweight tracking for a trusted team, not strong identity
# verification or enterprise-grade security.


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
    # Migrate an older version of this table that only had a boolean
    # "blocked" column -- anyone already granted access under the old
    # auto-approve system counts as already 'approved', not 'pending'.
    cols = [r[1] for r in conn.execute("PRAGMA table_info(viewer_log)").fetchall()]
    if "blocked" in cols and "status" not in cols:
        conn.execute("ALTER TABLE viewer_log ADD COLUMN status TEXT NOT NULL DEFAULT 'pending'")
        conn.execute("UPDATE viewer_log SET status = CASE WHEN blocked = 1 THEN 'revoked' ELSE 'approved' END")
    elif "status" not in cols:
        conn.execute("ALTER TABLE viewer_log ADD COLUMN status TEXT NOT NULL DEFAULT 'pending'")


def _get_viewer_status(email: str):
    """Returns 'pending' / 'approved' / 'revoked', or None if this email
    has never requested access before."""
    conn = get_connection()
    _ensure_viewer_table(conn)
    row = conn.execute("SELECT status FROM viewer_log WHERE email = ?", (email.strip().lower(),)).fetchone()
    conn.close()
    return row[0] if row else None


def _record_viewer_request(email: str, name: str) -> bool:
    """Log a viewer's access request (or a returning approved viewer's
    visit). Returns True if this is a BRAND NEW request (so the caller
    knows whether to email the Admin) -- a returning viewer's status is
    left untouched, only their name/last_seen/visits are refreshed."""
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
    """Refresh last_seen/visits for an ALREADY-approved viewer, without
    touching their status."""
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
    """Emails the HQ Admin when someone new requests access. Reads Gmail
    sender credentials from Streamlit secrets ([email] address /
    app_password) -- if those aren't configured, this silently does
    nothing (the request still shows up in-app either way, so nothing is
    lost, just no email alert). Returns True only if the email was
    actually sent."""
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
    """Shows first-run Admin setup (once), then a sign-in / continue-as-
    viewer screen on every subsequent run until someone picks one. Returns
    True once it's OK to show the actual dashboard."""
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
        # status is "revoked", or the record vanished somehow
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
    """HQ branding + the signed-in person's email, right-aligned at the
    top of the dashboard (not centered)."""
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
    """Admin-only sidebar panel: pending access requests needing a
    decision, then everyone already approved (with online status and a
    Remove button), then anyone previously revoked (with a Restore
    button)."""
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


def render_header(filters: dict):
    """One unified header, built as a single HTML component so everything
    lines up in a real single row: big roster stat cards on the left, a
    large layered-3D animated Tamkeen mark dead center, big money/orders
    stat cards on the right. Every card pops up (lift + tooltip) on hover."""
    merged = load_merged()
    # Placeholder drivers (created only because some report mentioned an
    # ID/name that didn't match anyone on the actual roster) should not
    # inflate headcount/active counts -- but their order/salary figures
    # still belong in the money totals, so those stay unfiltered.
    roster_only = merged[merged["is_placeholder"] == 0] if not merged.empty else merged

    headcount = active = 0
    orders_m = 0
    gross_m = 0.0

    if not merged.empty:
        if filters["month"]:
            # Scope headcount/active to drivers who actually have a record
            # for THIS month -- not the entire all-time roster -- so
            # switching the month filter shows that month's own numbers
            # instead of an ever-growing all-time total.
            month_df = roster_only[roster_only["month_year"] == filters["month"]]
            month_df = month_df[
                month_df["supervisor_name"].isin(filters["supervisors"])
                & month_df["vehicle_type"].isin(filters["vehicle_types"])
            ]
            headcount = month_df["driver_id"].nunique()
            active = int(month_df[month_df["status"] == "Active"]["driver_id"].nunique())
            money_df = merged[merged["month_year"] == filters["month"]]
            orders_m = int(money_df["total_orders"].sum())
            gross_m = float(money_df["gross_salary"].sum())
        else:
            roster = roster_only.drop_duplicates(subset="driver_id")
            roster_f = roster[
                roster["supervisor_name"].isin(filters["supervisors"])
                & roster["vehicle_type"].isin(filters["vehicle_types"])
            ]
            headcount = len(roster_f)
            active = int((roster_f["status"] == "Active").sum())

    # Build the layered "3D extrusion" for the logo mark: the same three
    # shapes drawn 7 times with a growing offset and darkening shade,
    # like a stack of cards, then the crisp full-color shapes on top.
    # This is what actually reads as "solid 3D", independent of rotation.
    extrusion_layers = ""
    layer_count = 7
    for i in range(layer_count, 0, -1):
        shade = 10 + i * 4  # darker for deeper (further-back) layers
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
        background: linear-gradient(135deg, rgba(34,197,94,0.12), rgba(59,130,246,0.09));
        border: 1px solid rgba(120,120,120,0.20);
        box-shadow: 0 6px 18px rgba(0,0,0,0.10);
        animation: statPop 0.5s ease both;
        transition: transform 0.22s cubic-bezier(.2,.8,.3,1.3), box-shadow 0.22s ease;
      }
      .stat-card:hover {
        transform: translateY(-8px) scale(1.05);
        box-shadow: 0 18px 34px rgba(0,0,0,0.24);
        z-index: 6;
      }
      .stat-card.right { background: linear-gradient(135deg, rgba(168,85,247,0.12), rgba(59,130,246,0.09)); }
      .stat-card .label { font-size: 13px; letter-spacing: 1.4px; text-transform: uppercase; opacity: 0.62; font-weight: 600; }
      .stat-card .value { font-size: 32px; font-weight: 900; margin-top: 4px; }
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
        <div class="stat-card">
          <div class="label">Headcount</div>
          <div class="value">__HEADCOUNT__</div>
          <div class="htip">Total drivers matching your current filters</div>
        </div>
        <div class="stat-card">
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
        <div class="stat-card right">
          <div class="label">Orders (month)</div>
          <div class="value">__ORDERS__</div>
          <div class="htip">Total orders logged for the selected month</div>
        </div>
        <div class="stat-card right">
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
        .replace("__GROSS__", f"Rs {gross_m:,.0f}")
    )
    components.html(html, height=320)


# ==============================================================================
# DATABASE LAYER
# ==============================================================================


def get_connection() -> sqlite3.Connection:
    """Open a fresh SQLite connection with foreign keys enabled."""
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def _add_column_if_missing(conn: sqlite3.Connection, table: str, column_def: str) -> None:
    """Auto-upgrade older database files: adds a column to an existing table
    if it is not already there. Safe to call every launch."""
    col_name = column_def.split()[0]
    existing = [r[1] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]
    if col_name not in existing:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column_def}")


def init_db() -> None:
    """Create tables on first launch if they do not already exist, and
    migrate older database files forward with any newly-added columns."""
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

    # --- Forward migrations for fields added after the initial release ---
    _add_column_if_missing(conn, "drivers", "iqama_number TEXT")
    _add_column_if_missing(conn, "drivers", "sponsor_name TEXT")
    _add_column_if_missing(conn, "drivers", "is_placeholder INTEGER NOT NULL DEFAULT 0")
    _add_column_if_missing(conn, "monthly_logs", "cancelled_orders INTEGER NOT NULL DEFAULT 0")
    _add_column_if_missing(conn, "monthly_logs", "pending_salary REAL NOT NULL DEFAULT 0")

    conn.commit()
    conn.close()


def upsert_driver(conn: sqlite3.Connection, row: dict) -> None:
    """Insert a driver, or update it in place if the driver_id already
    exists. Vehicle type is only overwritten when THIS row actually
    carries a real value for it -- a report sheet that has no
    vehicle-type column at all (e.g. an Orders, Validity, or Attendance
    report) must never silently reset an existing driver's 'Company Car'
    back to the default 'Own Car'. That silent reset on every later
    upload is what was inflating the Own Car count.

    Pass is_placeholder=True only when this row exists SOLELY because some
    OTHER report (orders/validity/salary/etc) mentioned an ID or name that
    couldn't be matched to anyone on the actual roster -- these are kept
    so the associated data isn't lost, but excluded from headcount/status/
    vehicle KPI counts elsewhere so those numbers reflect the real roster.
    A real roster row always clears this flag, even for a driver_id that
    was previously placeholder-only."""
    existing_vehicle = conn.execute(
        "SELECT vehicle_type FROM drivers WHERE driver_id = ?", (row["driver_id"],)
    ).fetchone()
    vehicle_type = row.get("vehicle_type") or (existing_vehicle[0] if existing_vehicle else None) or "Own Car"
    is_placeholder = 1 if row.get("is_placeholder") else 0

    conn.execute(
        """
        INSERT INTO drivers
            (driver_id, driver_name, phone, supervisor_name, status,
             vehicle_type, join_date, termination_date, iqama_number,
             sponsor_name, is_placeholder)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(driver_id) DO UPDATE SET
            driver_name       = excluded.driver_name,
            phone             = COALESCE(NULLIF(excluded.phone, ''), drivers.phone),
            supervisor_name   = COALESCE(NULLIF(excluded.supervisor_name, ''), drivers.supervisor_name),
            status            = excluded.status,
            vehicle_type      = excluded.vehicle_type,
            join_date         = COALESCE(NULLIF(excluded.join_date, ''), drivers.join_date),
            termination_date  = excluded.termination_date,
            iqama_number      = COALESCE(NULLIF(excluded.iqama_number, ''), drivers.iqama_number),
            sponsor_name      = COALESCE(NULLIF(excluded.sponsor_name, ''), drivers.sponsor_name),
            is_placeholder    = CASE WHEN excluded.is_placeholder = 0 THEN 0 ELSE drivers.is_placeholder END
        """,
        (
            row["driver_id"], row["driver_name"], row.get("phone"),
            row.get("supervisor_name"), row.get("status", "Active"),
            vehicle_type, row.get("join_date"),
            row.get("termination_date"), row.get("iqama_number"), row.get("sponsor_name"),
            is_placeholder,
        ),
    )


def upsert_monthly_log(conn: sqlite3.Connection, row: dict) -> None:
    """Insert a monthly log, or update it if one already exists for that
    driver + month (idempotent re-upload of the same file is safe). This is
    the low-level writer -- it always writes every field. For any upload
    that only carries SOME of the fields (e.g. an orders-only sheet), use
    merge_monthly_log() instead so untouched fields aren't wiped to zero."""
    conn.execute(
        """
        INSERT INTO monthly_logs
            (driver_id, month_year, total_orders, gross_salary,
             total_deductions, net_salary, validity_status, valid_days_in_month,
             cancelled_orders, pending_salary)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(driver_id, month_year) DO UPDATE SET
            total_orders        = excluded.total_orders,
            gross_salary         = excluded.gross_salary,
            total_deductions      = excluded.total_deductions,
            net_salary             = excluded.net_salary,
            validity_status         = excluded.validity_status,
            valid_days_in_month      = excluded.valid_days_in_month,
            cancelled_orders          = excluded.cancelled_orders,
            pending_salary             = excluded.pending_salary
        """,
        (
            row["driver_id"], row["month_year"], row["total_orders"],
            row["gross_salary"], row["total_deductions"], row["net_salary"],
            row["validity_status"], row["valid_days_in_month"],
            row.get("cancelled_orders", 0), row.get("pending_salary", 0.0),
        ),
    )


_MONTHLY_LOG_FIELDS = [
    "total_orders", "cancelled_orders", "gross_salary", "total_deductions",
    "pending_salary", "net_salary", "validity_status", "valid_days_in_month",
]
_MONTHLY_LOG_DEFAULTS = {
    "total_orders": 0, "cancelled_orders": 0, "gross_salary": 0.0,
    "total_deductions": 0.0, "pending_salary": 0.0, "net_salary": 0.0,
    "validity_status": "Valid", "valid_days_in_month": 0,
}


def get_monthly_log(conn: sqlite3.Connection, driver_id: str, month_year: str):
    """Fetch the current stored values for one driver+month, or None."""
    row = conn.execute(
        f"SELECT {', '.join(_MONTHLY_LOG_FIELDS)} FROM monthly_logs "
        f"WHERE driver_id = ? AND month_year = ?",
        (driver_id, month_year),
    ).fetchone()
    if row is None:
        return None
    return dict(zip(_MONTHLY_LOG_FIELDS, row))


def merge_monthly_log(conn: sqlite3.Connection, driver_id: str, month_year: str, updates: dict) -> None:
    """Apply a PARTIAL set of updates to a driver's monthly log, preserving
    every field the caller didn't mention. Pass None (or simply omit a key)
    for any field this particular upload doesn't know about -- e.g. an
    orders-report upload should pass cancelled_orders=None, not 0, so a
    later cancellations-report upload for the same month doesn't get wiped
    out, and vice versa."""
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
    """Generate n unique realistic driver names."""
    pool = set()
    while len(pool) < n:
        pool.add(f"{random.choice(FIRST_NAMES)} {random.choice(LAST_NAMES)}")
    return list(pool)


def seed_sample_data() -> None:
    """Wipe existing data and populate the DB with realistic demo records:
    a mix of Active/Terminated/Suspended drivers, Company/Own cars, IQAMA
    numbers, sponsors, and 4 months of payroll history including cancelled
    orders and pending-salary amounts."""
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
    """Danger-zone helper: wipes both tables completely."""
    conn = get_connection()
    conn.execute("DELETE FROM monthly_logs;")
    conn.execute("DELETE FROM drivers;")
    conn.commit()
    conn.close()


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
    """Drivers joined to their monthly logs (left join, so drivers with no
    logs yet for any month still appear). A blank supervisor is normalized
    to 'Unassigned' so those drivers stay visible in every filter instead
    of silently vanishing (a real gap in uploaded roster files)."""
    conn = get_connection()
    df = pd.read_sql_query(
        """
        SELECT d.driver_id, d.driver_name, d.phone, d.supervisor_name,
               d.status, d.vehicle_type, d.join_date, d.termination_date,
               d.iqama_number, d.sponsor_name, d.is_placeholder,
               m.log_id, m.month_year, m.total_orders, m.gross_salary,
               m.total_deductions, m.net_salary, m.validity_status,
               m.valid_days_in_month, m.cancelled_orders, m.pending_salary
        FROM drivers d
        LEFT JOIN monthly_logs m ON d.driver_id = m.driver_id
        """,
        conn,
    )
    conn.close()
    df["supervisor_name"] = df["supervisor_name"].fillna(UNASSIGNED).replace("", UNASSIGNED)
    return df


def distinct_months() -> list:
    conn = get_connection()
    rows = conn.execute(
        "SELECT DISTINCT month_year FROM monthly_logs ORDER BY month_year DESC"
    ).fetchall()
    conn.close()
    return [r[0] for r in rows]


def month_display(ym: str) -> str:
    """Turn 'YYYY-MM' into a human-readable 'Month YYYY' label (e.g.
    '2026-04' -> 'April 2026') for anywhere a month is shown to the person."""
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
            st.caption("Permanently wipes all drivers and monthly logs.")
            if st.button("Clear All Data", use_container_width=True):
                clear_all_data()
                st.success("Database cleared.")
                st.rerun()

    st.sidebar.markdown("---")
    st.sidebar.markdown("### \U0001F50E Global Filters")

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
    """Apply the shared sidebar filters to a merged drivers+logs dataframe."""
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

    # Placeholder drivers (created only because some report -- orders,
    # validity, salary, etc -- mentioned an ID/name that didn't match
    # anyone on the actual roster) must not inflate headcount/active/
    # vehicle-type counts. Their order/salary data is still usable
    # elsewhere (Rider Lookup, Financial totals); this only affects "how
    # many riders are there" style KPI cards.
    roster_source = merged[merged["is_placeholder"] == 0]

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

    roster = roster_source.drop_duplicates(subset="driver_id")
    roster_filtered = roster[
        roster["supervisor_name"].isin(filters["supervisors"])
        & roster["vehicle_type"].isin(filters["vehicle_types"])
    ]

    month_df = roster_source[roster_source["month_year"] == filters["month"]] if filters["month"] else roster_source.iloc[0:0]
    month_df = month_df[
        month_df["supervisor_name"].isin(filters["supervisors"])
        & month_df["vehicle_type"].isin(filters["vehicle_types"])
    ]
    money_df = merged[merged["month_year"] == filters["month"]] if filters["month"] else merged.iloc[0:0]
    money_df = money_df[
        money_df["supervisor_name"].isin(filters["supervisors"])
        & money_df["vehicle_type"].isin(filters["vehicle_types"])
    ]

    if filters["month"]:
        # Scoped to THIS month's own records -- not the all-time roster --
        # so picking a different month shows that month's own headcount
        # instead of a number that only ever grows across uploads.
        month_roster = month_df.drop_duplicates(subset="driver_id")
        total_headcount = month_roster["driver_id"].nunique()
        active_drivers = (month_roster["status"] == "Active").sum()
        suspended_drivers = int((month_roster["status"] == "Suspended").sum())
    else:
        total_headcount = len(roster_filtered)
        active_drivers = (roster_filtered["status"] == "Active").sum()
        suspended_drivers = int((roster_filtered["status"] == "Suspended").sum())

    terminated_this_month = 0
    if filters["month"]:
        term_mask = (
            roster_filtered["status"].eq("Terminated")
            & roster_filtered["termination_date"].fillna("").str.startswith(filters["month"])
        )
        terminated_this_month = int(term_mask.sum())

    company_cars = (month_df["vehicle_type"] == "Company Car").sum()
    own_cars = (month_df["vehicle_type"] == "Own Car").sum()
    valid_drivers = (month_df["validity_status"] == "Valid").sum()
    invalid_drivers = (month_df["validity_status"] == "Invalid").sum()

    st.markdown(f"**Selected Month:** `{month_display(filters['month'])}`")

    stat_cards([
        {"icon": "\U0001F465", "label": "Total Headcount", "value": total_headcount,
         "tip": "Drivers with a record for the selected month", "variant": "a"},
        {"icon": "\U0001F7E2", "label": "Active Drivers", "value": int(active_drivers),
         "tip": "Currently active, not terminated or suspended", "variant": "a"},
        {"icon": "\U0001F6D1", "label": "Terminated (this month)", "value": terminated_this_month,
         "tip": "Ending date falls within the selected month", "variant": "c"},
        {"icon": "\u23F8\uFE0F", "label": "Suspended Drivers", "value": suspended_drivers,
         "tip": "Temporarily suspended, not terminated", "variant": "d"},
    ])

    stat_cards([
        {"icon": "\U0001F697", "label": "Company Cars (month)", "value": int(company_cars),
         "tip": "Riders using a company-provided vehicle", "variant": "b"},
        {"icon": "\U0001F699", "label": "Own Cars (month)", "value": int(own_cars),
         "tip": "Riders using their own vehicle", "variant": "b"},
        {"icon": "\u2705", "label": "Valid Drivers (month)", "value": int(valid_drivers),
         "tip": "Marked Valid for this month's payroll", "variant": "a"},
        {"icon": "\u274C", "label": "Invalid Drivers (month)", "value": int(invalid_drivers),
         "tip": "Marked Invalid -- needs supervisor follow-up", "variant": "c"},
    ])

    stat_cards([
        {"icon": "\U0001F4E6", "label": "Total Orders (month)", "value": f"{int(money_df['total_orders'].sum()):,}",
         "tip": "Sum of all completed orders this month", "variant": "a"},
        {"icon": "\u274C", "label": "Cancelled Orders (month)", "value": f"{int(money_df['cancelled_orders'].sum()):,}",
         "tip": "Orders cancelled across the whole team", "variant": "c"},
        {"icon": "\U0001F4B5", "label": "Gross Salary (month)", "value": f"Rs {money_df['gross_salary'].sum():,.0f}",
         "tip": "Total pay before deductions", "variant": "b"},
        {"icon": "\u23F3", "label": "Pending Salary (month)", "value": f"Rs {money_df['pending_salary'].sum():,.0f}",
         "tip": "Amount still owed, not yet paid out", "variant": "d"},
    ])

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
            {"icon": "\U0001F4B0", "label": "Total Money Received", "value": f"Rs {(company_total or 0):,.2f}",
             "tip": "Company-level total payable amount for this billing cycle", "variant": "a"},
            {"icon": "\U0001F9FE", "label": "Tax Amount", "value": f"Rs {(tax_amount or 0):,.2f}",
             "tip": "Tax amount for this billing cycle, from the salary summary sheet", "variant": "c"},
            {"icon": "\U0001F4C4", "label": "Invoice Amount", "value": f"Rs {(invoice_amount or 0):,.2f}",
             "tip": "Invoiced amount for this billing cycle", "variant": "b"},
        ])
    else:
        st.caption(
            "\u2139\uFE0F No company-level salary summary (Total Money Received / "
            "Tax Amount) found for this month yet -- upload it from "
            "**Upload Monthly Data \u2192 Salary Data**."
        )

    stat_cards([
        {"icon": "\U0001F4B5", "label": "Gross Salary", "value": f"Rs {total_gross:,.0f}",
         "tip": "Total pay before deductions", "variant": "b"},
        {"icon": "\u2796", "label": "Deductions", "value": f"Rs {total_deductions:,.0f}",
         "tip": "Total amounts deducted this month", "variant": "c"},
        {"icon": "\u2705", "label": "Net Salary", "value": f"Rs {total_net:,.0f}",
         "tip": "Gross minus deductions", "variant": "a"},
        {"icon": "\u23F3", "label": "Pending Salary", "value": f"Rs {total_pending:,.0f}",
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
#
# Real-world operations exports rarely match a fixed schema -- headers vary
# ("Courier ID" vs "driver_id"), some files are roster-only (no payroll
# columns at all), dates come in mixed formats, and numeric IDs can arrive
# as floats in scientific notation. Instead of forcing a rigid template,
# this tab lets the user MAP their file's actual columns onto our internal
# fields via dropdowns (with smart auto-guessed defaults), and gracefully
# skips anything that isn't provided.

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
    """Return the best-matching source column for a field, or NONE_OPTION.
    A bare, ultra-generic header like "Total" or "Amount" is excluded from
    the fuzzy substring fallback -- it would otherwise fuzzy-match almost
    any alias list (a deduction sheet's "Total", a validity sheet's
    "Total", an orders sheet's "Total" all look identical by header name
    alone), silently misclassifying entire sheets. Exact matches are still
    allowed even for these words."""
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
    """Robustly parse a numeric cell that may contain commas ('1,234'),
    currency symbols ('Rs 45,000', '$320'), stray whitespace, a blank, or
    placeholder text ('N/A', '-'). Returns 0 (never raises) for anything
    that isn't recognizably a number, so one messy cell can't silently
    drop an entire row's totals out of the count."""
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
    """Turn IDs that arrived as floats (e.g. 1.761062e+15 from Excel) into
    clean strings without scientific notation or trailing .0. Also
    normalizes case/whitespace (upper + strip) so the SAME rider's ID
    matches across different sheets/files even if one export wrote it as
    'drv-1001' and another as 'DRV-1001' -- a mismatch here used to
    silently create a duplicate 'Unknown Rider' placeholder and make the
    real driver's orders show up as zero."""
    if pd.isna(v):
        return ""
    if isinstance(v, float):
        s = str(int(v)) if v.is_integer() else str(v)
    else:
        s = str(v).strip()
    return s.strip().upper()


def _clean_date_value(v):
    """Normalize a date cell (Timestamp, datetime, or assorted string
    formats) into 'YYYY-MM-DD'. Returns None if empty/unparseable-safe."""
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


def _clean_month_value(v) -> str:
    """Normalize a payroll-month cell into 'YYYY-MM' regardless of how it
    arrived (a full date, 'July 2026', '07/2026', '2026-07-01', etc.) so
    uploaded months always line up with the sidebar's Month filter instead
    of silently creating a mismatched, invisible month."""
    if pd.isna(v):
        return ""
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.strftime("%Y-%m")
    s = str(v).strip()
    if not s or s.lower() in ("nan", "nat"):
        return ""
    if len(s) == 7 and s[4] == "-":
        return s  # already YYYY-MM
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%B %Y", "%b %Y", "%b-%y", "%m/%Y", "%Y/%m", "%B-%y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m")
        except ValueError:
            continue
    return s  # fall back to the raw text rather than dropping the row


_MONTH_NAME_TO_NUM = {
    "jan": 1, "january": 1, "feb": 2, "february": 2, "mar": 3, "march": 3,
    "apr": 4, "april": 4, "may": 5, "jun": 6, "june": 6, "jul": 7, "july": 7,
    "aug": 8, "august": 8, "sep": 9, "sept": 9, "september": 9, "oct": 10,
    "october": 10, "nov": 11, "november": 11, "dec": 12, "december": 12,
}


def _infer_month_from_filename(filename: str) -> str:
    """Best-effort guess at the payroll month from a filename like
    'JULY_TAMKEEN_ACTIVE_RIDERS.xlsx' or 'Payroll_2026-07.csv'. Returns
    'YYYY-MM' if a month name/number and (optionally) a year could be
    found, else an empty string so the caller can fall back sensibly."""
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
    """Returns 'Company Car' / 'Own Car' when the cell clearly says so, or
    None when we can't tell (blank cell or unrecognized text). Callers
    must NOT turn that None into a guessed 'Own Car' themselves -- let
    upsert_driver() decide, so a blank cell in one upload doesn't silently
    overwrite a driver's real vehicle type from an earlier upload."""
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
    """Import a dataframe using a user-confirmed column mapping.
    `mapping` maps our internal field name -> the source column name in df,
    or NONE_OPTION if that field wasn't provided in the file.
    Returns (success_count, errors)."""

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
                # This row's ID doesn't match anyone in the roster -- before
                # treating it as a brand-new rider, check if the NAME on
                # this row matches an existing roster rider (different ID
                # scheme, typo, etc.). Matching by name here is what stops
                # the same person being counted twice under two IDs.
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
    """Fraction of column names that look like pandas' auto-generated
    'Unnamed: N' placeholder -- a strong signal that the real header row
    was one row lower than where we read from."""
    if len(columns) == 0:
        return 1.0
    unnamed = sum(1 for c in columns if str(c).strip().lower().startswith("unnamed"))
    return unnamed / len(columns)


def _read_excel_smart(file_obj, sheet_name):
    """Read an Excel sheet, automatically picking whichever of header
    row 0 or row 1 produces cleaner column names. Many real-world report
    exports (like multi-week attendance/order sheets) have a merged title
    row above the actual column headers, which otherwise turns every
    column into a useless 'Unnamed: N'. Falls back gracefully if the sheet
    is too short to even have a second header row."""
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
        return df1, 2  # human-friendly: "row 2"
    df0.columns = [str(c).strip() for c in df0.columns]
    return df0, 1


# ==============================================================================
# WHOLE-WORKBOOK AUTO-IMPORT (every sheet, auto-classified)
# ==============================================================================
# Real operations exports often split one month's data across many sheets --
# a roster, a day-by-day order count, a day-by-day validity grid, a
# day-by-day attendance grid, a cancellation log, plus other sheets that
# aren't relevant to payroll at all (shift schedules, accident reports...).
# Rather than making the person pick one sheet at a time, this pipeline
# reads every sheet, figures out what KIND of data it holds by inspecting
# its columns and sample values, and merges everything it recognizes into
# one row per driver per month -- skipped sheets are reported clearly so
# nothing silently vanishes.

VALIDITY_TOKENS = {"VALID", "INVALID"}
ATTENDANCE_TOKENS = {"P", "OFF", "A", "ABSENT", "PRESENT", "OFFDAY", "OFF DAY", "LEAVE"}

SUMMARY_ROW_TOKENS = {"total", "totals", "grand total", "grand totals", "sum", "overall", "subtotal"}


def _row_is_summary(raw_row) -> bool:
    """True if a spreadsheet row is a 'Total' / 'Grand Total' row rather
    than a real driver row. These rows sometimes leave the name/ID cells
    genuinely blank (already skipped elsewhere), but sometimes carry a
    leftover label or stray value in one of those cells -- if that
    happens, the row's already-summed number gets counted as if it were
    one MORE driver's orders on top of everyone already counted, inflating
    the month's total above the real figure."""
    for v in raw_row:
        if pd.isna(v):
            continue
        s = str(v).strip().lower()
        if s in SUMMARY_ROW_TOKENS or "grand total" in s:
            return True
    return False


# ==============================================================================
# SALARY WORKBOOK UPLOAD  (separate from the roster/orders/validity upload)
# ==============================================================================
# The monthly salary export (e.g. "Feb_salary_main_sheet.xlsx") has its own
# distinct layout: a per-rider sheet (headers change slightly month to
# month -- "JAN SALARY " vs "riderDetail") with columns like Courier ID,
# Total payable amount, TOTAL DEDUCTION, FINAL SALARY, PENDING; and a
# one-row company-summary sheet ("partnerDetail") with Tax Amount and
# Total payable amount at the company level. This uses EXACT column-name
# matching only (no fuzzy substring fallback) -- "Deduction" (one line
# item) and "TOTAL DEDUCTION" (the actual total) are both real columns in
# this file, and a substring match would confuse the two.

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
    """Like _guess_column, but EXACT normalized match only -- no substring
    fallback. Use this whenever similarly-named columns in the same sheet
    (e.g. 'Deduction' vs 'TOTAL DEDUCTION') could otherwise be confused."""
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
    """Returns 'salary_detail' (per-rider), 'salary_summary' (one-row
    company total), or 'unrecognized'."""
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
    """Returns (rows, id_to_name) -- rows is a list of per-rider salary
    update dicts ready for merge_monthly_log."""
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
    """Reads every sheet, classifies each as per-rider salary detail,
    company-level summary, or unrecognized, and merges everything found
    into monthly_logs (gross_salary/total_deductions/net_salary/
    pending_salary) plus the salary_summary table for company totals."""
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
        df.columns = [str(c).strip() for c in df.columns]
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
                        "is_placeholder": True,
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
    """Columns representing one calendar day each -- either literally
    named '1'..'31' (a plain day-of-month grid), OR an actual date value
    like '2026-05-01' or '2026-05-01 00:00:00' (common when the sheet's
    real header row uses full dates instead of bare day numbers). Both
    layouts are used interchangeably by validity/attendance/order-count
    report sheets, so both must be recognized the same way."""
    out = []
    for c in columns:
        s = str(c).strip()
        if s.isdigit() and 1 <= int(s) <= 31:
            out.append(c)
            continue
        date_part = s.split(" ")[0]
        matched = False
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                datetime.strptime(date_part, fmt)
                matched = True
                break
            except ValueError:
                continue
        if matched:
            out.append(c)
    return out


def _classify_sheet(df: pd.DataFrame) -> str:
    """Return one of: 'roster', 'orders', 'validity', 'attendance',
    'cancellation', or 'unrecognized'."""
    cols = list(df.columns)

    has_name = _guess_column(cols, FIELD_ALIASES["driver_name"]) != NONE_OPTION
    id_col = _guess_column(cols, FIELD_ALIASES["driver_id"])
    has_id = id_col != NONE_OPTION

    # A roster sheet is distinguished from OTHER Courier-ID+Name sheets
    # (accident logs, equipment/petrol-card logs, order reports -- all of
    # which legitimately carry the same ID/name columns as the roster,
    # since they're exports from the same rider database) by genuine
    # vehicle-assignment evidence: either a column literally headed
    # "Vehicle Type" etc, OR -- since many real exports bury this info in
    # a generically-named column like "Status" -- a column whose ACTUAL
    # VALUES are overwhelmingly literal "Own Car"/"Company Car" text.
    # A loose "any column with 'plate' in its name" or "any column named
    # something like 'status'" check used to be enough to trigger this,
    # but both of those show up on plenty of non-roster sheets too (an
    # accident log's "Plate_No", an equipment log's workflow "STATUS")
    # and were silently corrupting driver records -- most damagingly, an
    # accident/equipment sheet's own "Date"/"Ending Date" column (about
    # the incident or the equipment task, not the RIDER) getting read as
    # that rider's termination date and marking active people "Terminated".
    has_vehicle_hint = _guess_column(cols, FIELD_ALIASES["vehicle_type"]) != NONE_OPTION
    if not has_vehicle_hint and has_name and has_id:
        name_col = _guess_column(cols, FIELD_ALIASES["driver_name"])
        content_col = _find_vehicle_type_column_by_content(df, {id_col, name_col})
        if content_col is not None and _vehicle_type_content_ratio(df, content_col) >= 0.9:
            has_vehicle_hint = True

    # Roster is checked FIRST, before the orders-column shortcut below. A
    # roster export that also happens to carry an orders-like column (e.g.
    # a "Total Trips" summary column) must NOT be misclassified as a
    # dedicated orders sheet -- doing so used to add its numbers a second
    # time on top of the real orders sheet elsewhere in the same workbook,
    # which is what caused totals to come out doubled.
    if has_name and has_id and has_vehicle_hint:
        return "roster"

    # Day-by-day grid sheets (columns literally named '1'..'31') are
    # classified by their CELL CONTENT, not by any header name -- this is
    # checked before the generic total_orders header check below because a
    # day-by-day ORDER COUNT grid often has nothing but a vague "Total"
    # column for its header, which is deliberately excluded from fuzzy
    # header matching (see _GENERIC_HEADER_STOPWORDS) to stop it from
    # mis-tagging validity/deduction/other sheets as "orders". Content is
    # the reliable signal here: mostly VALID/INVALID -> validity; mostly
    # P/OFF/ABSENT-style attendance tokens -> attendance; mostly plain
    # numbers (daily order counts, with some "OFF"/"NO SHIFT" text on
    # non-working days) -> orders.
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
            # A target/coaching-tracker sheet (e.g. "TARGET B RIDERS") --
            # its ORDERS column is an interim snapshot for a subset of
            # riders, not the authoritative monthly total. Summing it in
            # alongside the real order report double-counts those riders.
            return "unrecognized"
        return "orders"

    lower_cols = [str(c).strip().lower() for c in cols]
    if any("order id" in c for c in lower_cols) and any("name" in c for c in lower_cols) and not has_id:
        return "cancellation"

    return "unrecognized"


def _vehicle_type_content_ratio(df: pd.DataFrame, col) -> float:
    """Fraction of a column's actual values that are literally 'Own Car' /
    'Company Car' text."""
    vals = df[col].dropna().astype(str).str.strip().str.upper()
    if vals.empty:
        return 0.0
    sample = vals.head(60)
    hits = sample.apply(lambda s: "OWN CAR" in s or "COMPANY CAR" in s).sum()
    return hits / len(sample)


def _column_looks_like_vehicle_type(df: pd.DataFrame, col) -> bool:
    """True if most of a column's actual values are literally 'Own Car' /
    'Company Car' -- REGARDLESS of what the column header says. Real
    exports sometimes reuse a header like 'Status' for the vehicle
    assignment instead of a proper 'Vehicle Type' column, so relying on
    the header name alone misses it and every driver silently defaults to
    'Own Car'."""
    if col == NONE_OPTION:
        return False
    return _vehicle_type_content_ratio(df, col) >= 0.9


def _find_vehicle_type_column_by_content(df: pd.DataFrame, exclude_cols: set):
    """Scan every column NOT already claimed by another field and pick the
    one with the HIGHEST own-car/company-car match ratio (not just the
    first one crossing a threshold) -- a 'Plate Number' column that only
    sometimes reads 'Own Car' (for riders with no plate) must lose out to
    a dedicated column that reads 'Own Car'/'Company Car' on every row."""
    best_col, best_ratio = None, 0.0
    for c in df.columns:
        if c in exclude_cols:
            continue
        ratio = _vehicle_type_content_ratio(df, c)
        if ratio > best_ratio:
            best_col, best_ratio = c, ratio
    return best_col if best_ratio >= 0.5 else None


_SECTION_STATUS_KEYWORDS = [
    (["terminat"], "Terminated"),
    (["vacation", "vocation", "on leave", "leave"], "Suspended"),
    (["suspend"], "Suspended"),
    (["active rider", "back to active", "reactivat"], "Active"),
]


def _section_status_label(raw_row):
    """Real-world roster sheets sometimes pack THREE lists into one sheet:
    the active roster, then a 'TERMINATE FOR THIS MONTH' label followed by
    terminated riders, then an 'ON VOCATION' label followed by suspended
    riders -- with no per-row status column at all, just these section
    headers. A section-header row has exactly ONE filled cell in the
    entire row (a label, not driver data); a real driver row always has
    several (at minimum an ID and a name). Returns the status that should
    apply to rows AFTER this one, or None if this isn't a section header."""
    texts = [str(v).strip() for v in raw_row if not (isinstance(v, float) and pd.isna(v)) and str(v).strip()]
    if len(texts) != 1:
        return None
    label = texts[0].lower()
    for keywords, status in _SECTION_STATUS_KEYWORDS:
        if any(k in label for k in keywords):
            return status
    return None


def _extract_roster(df: pd.DataFrame, default_month: str = None) -> dict:
    """driver_id -> roster fields dict, plus a name->id lookup for later
    name-based matching (used by the cancellation sheet, which has no ID)."""
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

    # Content beats header name for vehicle type: even when a column was
    # found BY HEADER NAME, also check every other column's actual values.
    # If some other column is (almost) entirely literal 'Own Car'/'Company
    # Car' text, that's unambiguous ground truth and wins -- a header-name
    # guess can point at the wrong column (e.g. a generic "Type" column
    # that isn't actually vehicle type at all), but real 'Own Car'/
    # 'Company Car' text in a column can't lie about what it is.
    exclude = {id_col, name_col, first_col, last_col, phone_col, sup_col,
               sponsor_col, iqama_col, join_col, end_col}
    content_veh_col = _find_vehicle_type_column_by_content(df, exclude)
    if content_veh_col is not None and content_veh_col != veh_col:
        content_ratio = _vehicle_type_content_ratio(df, content_veh_col)
        header_ratio = _vehicle_type_content_ratio(df, veh_col) if veh_col != NONE_OPTION else 0.0
        if content_ratio >= 0.9 and content_ratio > header_ratio:
            veh_col = content_veh_col
    if veh_col != NONE_OPTION and status_col == veh_col:
        # Whatever we ended up using for vehicle type isn't really a
        # driver-status column -- don't feed vehicle-type text into status.
        status_col = NONE_OPTION
    elif _column_looks_like_vehicle_type(df, status_col):
        # A dedicated vehicle-type column exists, but the "status" column
        # ALSO turned out to be vehicle-type text (not a real Active/
        # Terminated status) -- don't let it feed driver status either.
        status_col = NONE_OPTION

    records = {}
    section_status = "Active"
    for _, raw in df.iterrows():
        detected_section = _section_status_label(raw)
        if detected_section is not None:
            section_status = detected_section
            continue  # this row is a label ("TERMINATE FOR THIS MONTH"), not a rider
        if _row_is_summary(raw):
            continue
        if name_col != NONE_OPTION:
            name = str(raw[name_col]).strip()
        elif first_col != NONE_OPTION or last_col != NONE_OPTION:
            f = "" if first_col == NONE_OPTION or pd.isna(raw[first_col]) else str(raw[first_col]).strip()
            l = "" if last_col == NONE_OPTION or pd.isna(raw[last_col]) else str(raw[last_col]).strip()
            name = f"{f} {l}".strip()
        else:
            name = ""
        if not name or name.lower() == "nan":
            continue

        driver_id = _clean_id_value(raw[id_col]) if id_col != NONE_OPTION else ""
        if not driver_id:
            driver_id = _slugify_name_to_id(name)

        termination_date = _clean_date_value(raw[end_col]) if end_col != NONE_OPTION else None
        status_val = str(raw[status_col]).strip() if status_col != NONE_OPTION and not pd.isna(raw[status_col]) else ""
        if section_status != "Active":
            # A section header ("TERMINATE FOR THIS MONTH", "ON VOCATION")
            # is an explicit, deliberate signal from whoever built the
            # sheet -- it overrides the column-based guess below. If the
            # sheet didn't also give an explicit date, anchor it to the
            # month being imported so "Terminated (this month)" on the
            # dashboard (which keys off termination_date) still counts it.
            status = section_status
            if not termination_date and default_month:
                termination_date = f"{default_month}-01"
        else:
            status = status_val if status_val in DRIVER_STATUSES else ("Terminated" if termination_date else "Active")

        records[driver_id] = {
            "driver_id": driver_id,
            "driver_name": name,
            "phone": None if phone_col == NONE_OPTION or pd.isna(raw[phone_col]) else str(raw[phone_col]).strip(),
            "supervisor_name": None if sup_col == NONE_OPTION or pd.isna(raw[sup_col]) else str(raw[sup_col]).strip(),
            "sponsor_name": None if sponsor_col == NONE_OPTION or pd.isna(raw[sponsor_col]) else str(raw[sponsor_col]).strip(),
            "iqama_number": None if iqama_col == NONE_OPTION else (_clean_id_value(raw[iqama_col]) or None),
            "vehicle_type": _clean_vehicle_type(raw[veh_col]) if veh_col != NONE_OPTION else None,
            "status": status,
            "join_date": _clean_date_value(raw[join_col]) if join_col != NONE_OPTION else None,
            "termination_date": termination_date,
        }
    return records


def _extract_orders(df: pd.DataFrame):
    """Returns (orders_dict, id_to_name) -- driver_id -> total_orders
    (summed if a driver appears twice WITHIN this same sheet), plus
    whatever name was sitting next to each ID so a later ID mismatch can
    be resolved by NAME instead of guessing from the ID itself."""
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
            # No single "Total Orders" column -- this is a day-by-day grid
            # (columns named 1..31) holding a daily order count per day,
            # with text like "OFF"/"NO SHIFT"/"ABSENT" on non-working days.
            # Sum whichever cells are actually numbers; non-numeric cells
            # contribute 0 rather than breaking the row.
            row_total = sum(_clean_number_value(raw[c], as_int=True) for c in day_cols)
        out[driver_id] = out.get(driver_id, 0) + row_total
        if name_col != NONE_OPTION and not pd.isna(raw[name_col]):
            nm = str(raw[name_col]).strip()
            if nm:
                id_to_name[driver_id] = nm
    return out, id_to_name


def _orders_dicts_look_like_duplicates(a: dict, b: dict) -> bool:
    """True when two per-sheet {driver_id: total_orders} maps look like the
    SAME report saved under two different sheet names (e.g. a monthly
    total re-pasted into a second tab) rather than genuinely different data
    (e.g. two separate weeks) that should be added together. We only treat
    them as duplicates when they share a large chunk of drivers AND those
    shared drivers carry near-identical values -- a real second week of
    orders would share the driver list but NOT the values."""
    common = set(a) & set(b)
    if not common or len(common) < max(1, min(len(a), len(b)) // 2):
        return False
    matches = sum(1 for k in common if a[k] == b[k])
    return (matches / len(common)) >= 0.8


def _extract_validity(df: pd.DataFrame):
    """Returns (validity_dict, id_to_name): driver_id -> {'valid_days': n,
    'invalid_days': n, 'status': str}, plus each row's name for later
    name-based ID resolution."""
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
    """Returns (attendance_dict, id_to_name): driver_id -> days_worked
    (count of 'P' cells across the day columns), plus each row's name for
    later name-based ID resolution."""
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
    """normalized driver name -> cancelled order count (this report has no
    driver ID, only a name, so it's matched against the roster afterwards)."""
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
    """Match a transactional-log name (often abbreviated, e.g. 'MD UDDIN')
    against the roster's full names (e.g. 'MD MARUF UDDIN') by requiring at
    least 2 shared MEANINGFUL words -- catches real matches that an exact-
    string comparison misses. Extremely common South-Asian name prefixes
    ('MD', 'MOHAMMAD', 'SYED', ...) are excluded from that word count: two
    DIFFERENT riders both named 'MD KARIM ...' would otherwise share 'MD' +
    one more word and get incorrectly merged into a single rider, mixing
    their orders together. If fewer than 2 meaningful words remain after
    removing those prefixes, we deliberately do NOT guess -- a missed
    match (left as a separate 'Unknown Rider' you can fix by hand) is far
    safer than silently combining two different people's data."""
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
    """When a report sheet's ID column uses a different numbering scheme
    than the roster (or has a typo), matching by ID alone creates a
    duplicate 'Unknown Rider' entry for a person who already exists in the
    roster -- so that rider's orders/validity/attendance get counted
    twice: once under their real driver_id, once under the mismatched one.
    This resolves any key that ISN'T already a known driver_id by matching
    the NAME that sat next to it in the sheet against the roster (fuzzy,
    same logic used for the cancellation report) -- never by phone."""
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


def process_workbook_all_sheets(uploaded_file, month_year: str):
    """Read every sheet in the workbook, classify each one, and merge
    everything recognized into unified driver + monthly_log records.
    Returns a summary dict for display plus the usual (success_count, errors)."""
    xls = pd.ExcelFile(uploaded_file)

    roster_records = {}
    orders_sheets = []  # list of (sheet_name, {driver_id: total_orders}) -- combined AFTER the loop
    validity_by_id = {}
    attendance_by_id = {}
    cancellations_by_name = {}
    id_to_name = {}  # driver_id -> name, gathered from every report sheet, used to resolve ID mismatches
    sheet_report = []  # (sheet_name, classification, row_count)

    for sheet_name in xls.sheet_names:
        try:
            df, _hdr = _read_excel_smart(uploaded_file, sheet_name)
        except Exception as exc:  # noqa: BLE001
            sheet_report.append((sheet_name, f"error reading sheet: {exc}", 0))
            continue

        df = df.dropna(axis=0, how="all").reset_index(drop=True)
        df.columns = [str(c).strip() for c in df.columns]
        if df.empty:
            sheet_report.append((sheet_name, "empty", 0))
            continue

        kind = _classify_sheet(df)
        sheet_report.append((sheet_name, kind, len(df)))

        if kind == "roster":
            roster_records.update(_extract_roster(df, month_year))
        elif kind == "orders":
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

    # Combine the "orders"-classified sheets. Sheets whose driver->orders
    # values are near-identical to a sheet we already kept are treated as
    # the SAME report duplicated across tabs and skipped, instead of being
    # added a second time (the root cause of totals coming out doubled).
    # Genuinely different sheets (e.g. separate weeks) are still summed.
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

    # Resolve any report key that ISN'T already a known driver_id by
    # matching its NAME against the roster instead (fuzzy, same approach
    # as the cancellation report) -- this is what stops a rider from being
    # counted once under their real driver_id and AGAIN under a mismatched
    # ID from a report sheet that numbers riders differently.
    all_drivers_df = load_drivers()
    known_driver_ids = set(all_drivers_df["driver_id"])
    name_to_id = dict(zip(all_drivers_df["driver_name"].str.strip().str.upper(), all_drivers_df["driver_id"]))

    orders_by_id = _remap_ids_by_name(orders_by_id, id_to_name, name_to_id, known_driver_ids)
    validity_by_id = _remap_ids_by_name(validity_by_id, id_to_name, name_to_id, known_driver_ids)
    attendance_by_id = _remap_ids_by_name(attendance_by_id, id_to_name, name_to_id, known_driver_ids)

    # Resolve name-based cancellations against the FULL roster on file
    # (including drivers from earlier uploads), using fuzzy word matching.
    cancellations_by_id = {}
    unmatched_names = 0
    for name_key, count in cancellations_by_name.items():
        did = _fuzzy_match_name_to_id(name_to_id, name_key)
        if did:
            cancellations_by_id[did] = cancellations_by_id.get(did, 0) + count
        else:
            unmatched_names += 1

    # Every driver that appeared ANYWHERE in this workbook (roster included)
    # gets a monthly presence record for this month -- even an all-zero one
    # for a roster-only entry -- so "how many drivers this month" reflects
    # exactly who was in THIS file, not the all-time roster.
    all_driver_ids = (
        set(orders_by_id) | set(validity_by_id) | set(attendance_by_id)
        | set(cancellations_by_id) | set(roster_records)
    )

    # A driver_id can show up in ORDER REPORT / VALIDTY REPORT / etc. without
    # ever appearing in the roster sheet (different sheet, slightly different
    # export, a rider who left before the roster snapshot was taken...). The
    # monthly_logs table requires a matching drivers row to exist first, so
    # anyone not already known gets a minimal placeholder profile here --
    # better than losing their payroll data to a crash.
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
                "is_placeholder": True,
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
            validity_status = None  # attendance alone doesn't tell us pass/fail validity
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
            },
        )
        logs_written += 1

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
    }
    return summary


def render_upload_tab():
    if not is_admin():
        st.info(
            "\U0001F512 Only the HQ Admin can upload or import data. "
            "You're viewing this dashboard as a read-only Viewer."
        )
        return

    tab_ops, tab_salary = st.tabs(["\U0001F4E4 Roster / Orders / Validity", "\U0001F4B0 Salary Data"])
    with tab_ops:
        _render_operations_upload()
    with tab_salary:
        _render_salary_upload()


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
                "contain the same orders data under different names, only one is counted."
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

    # Plain CSV -- always the single-table flow.
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
    """The original column-mapping flow for a single table (one Excel sheet
    or a CSV) -- kept as the manual/advanced path, and still the only path
    for plain CSVs and single-sheet workbooks."""
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


def render_month_reveal_cards(rows: pd.DataFrame, join_date: str, vehicle_type: str):
    """Render one animated card per month of history for a rider -- each
    card fades/slides in with a small staggered delay so the reveal feels
    like an animation rather than a plain table dump. Join date and
    vehicle type are driver-level facts, repeated on every card so they're
    visible alongside each month's figures without having to scroll back
    up to the profile header."""
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
              <div class="reveal-row"><span>Gross</span><b>Rs {row.gross_salary:,.0f}</b></div>
              <div class="reveal-row"><span>Deductions</span><b>Rs {row.total_deductions:,.0f}</b></div>
              <div class="reveal-row"><span>Pending</span><b>Rs {row.pending_salary:,.0f}</b></div>
              <div class="reveal-row"><span>Net</span><b>Rs {row.net_salary:,.0f}</b></div>
              <div class="reveal-row"><span>Joined</span><b>{join_date or 'N/A'}</b></div>
              <div class="reveal-row"><span>Vehicle</span><b>{vehicle_icon} {vehicle_type}</b></div>
              <div class="reveal-badge {badge_class}">{row.validity_status}</div>
            </div>
            """
        )
    st.markdown(css + "".join(cards) + "</div>", unsafe_allow_html=True)


def render_rider_lookup():
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

    st.markdown("#### Lifetime Totals")
    stat_cards([
        {"icon": "\U0001F4E6", "label": "Total Orders", "value": f"{int(rider_logs['total_orders'].sum()):,}",
         "tip": "Across every month on file", "variant": "a"},
        {"icon": "\u274C", "label": "Total Cancelled", "value": f"{int(rider_logs['cancelled_orders'].sum()):,}",
         "tip": "Cancelled orders, all-time", "variant": "c"},
        {"icon": "\U0001F4B5", "label": "Total Gross Earned", "value": f"Rs {rider_logs['gross_salary'].sum():,.0f}",
         "tip": "Before deductions, all-time", "variant": "b"},
        {"icon": "\u23F3", "label": "Total Pending", "value": f"Rs {rider_logs['pending_salary'].sum():,.0f}",
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
        {"icon": "\U0001F4B8", "label": "Total Deductions", "value": f"Rs {total_deductions:,.0f}",
         "tip": "Summed across the team this month", "variant": "b"},
    ])

    line1 = (
        f"\u26A0\uFE0F *{supervisor} Team Alert \u2013 {month}*: {total_riders} riders tracked, "
        f"{invalid_count} invalid aur {low_order_count} riders {low_order_threshold} se kam orders par hain."
    )
    line2 = (
        f"\U0001F4B0 Total deductions is month Rs {total_deductions:,.0f} rahi hain, {total_cancelled} orders cancel huay, "
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

    filters = render_sidebar()
    render_header(filters)

    st.markdown(
        "<h2 style='text-align:center; margin-top:0;'>Food Delivery Logistics, Operations & Payroll Dashboard</h2>"
        "<p style='text-align:center; opacity:0.7; margin-top:-8px;'>Driver operations, workforce KPIs, payroll, "
        "bulk uploads, rider lookup & supervisor alerts -- all in one place.</p>",
        unsafe_allow_html=True,
    )

    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        [
            "\U0001F4CA Operations Dashboard",
            "\U0001F4B0 Financial & Payroll",
            "\U0001F4E4 Upload Monthly Data",
            "\U0001F50D Rider Lookup",
            "\U0001F4E3 Supervisor Alerts",
        ]
    )

    with tab1:
        render_dashboard(filters)
    with tab2:
        render_financials(filters)
    with tab3:
        render_upload_tab()
    with tab4:
        render_rider_lookup()
    with tab5:
        render_supervisor_alerts()


if __name__ == "__main__":
    main()
