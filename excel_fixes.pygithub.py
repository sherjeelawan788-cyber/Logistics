# =============================================================================
# excel_fixes.py  --  loaded by app.py (see the exec(...) line at the bottom of
# app.py). Every function below REPLACES the function with the same name in
# app.py, so nothing else in app.py needs to be edited.
#
# What this fixes for Excel uploads:
#   1. Header row detection: tries rows 1, 2 and 3 and keeps the one that
#      actually contains a Courier/Driver ID column and/or day columns.
#   2. Day columns: also understands "1.0", "1-Oct", "Oct 1", "01 Oct 2026",
#      real Excel dates, etc. (Google Sheets headers were already fine).
#   3. Orders: if a tab's own "Total Orders" cell is blank/0 for a rider, the
#      sum of that rider's day-by-day cells is used instead of 0.
#   4. A roster tab's own order total of 0 no longer overrides real orders.
#   5. After an Excel import, a report stays on screen (it used to vanish
#      because the page reloads) showing how every tab was read.
# =============================================================================
import re as _re
from datetime import datetime as _dt

_DAY_FORMATS = (
    "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y",
    "%d/%m/%y", "%d-%b-%Y", "%d-%b-%y", "%d %b %Y", "%d %B %Y",
    "%d-%b", "%d %b", "%b %d", "%B %d",
)


def _day_from_header(col):
    s = str(col).strip()
    if not s:
        return None
    if _re.fullmatch(r"\d{1,2}(\.0+)?", s):
        n = int(float(s))
        return n if 1 <= n <= 31 else None
    for candidate in (s, s.split(" ")[0]):
        for fmt in _DAY_FORMATS:
            try:
                return _dt.strptime(candidate, fmt).day
            except ValueError:
                continue
    return None


def _day_number_columns(columns) -> list:
    return [c for c in columns if _day_from_header(c) is not None]


def _day_column_to_daynum(col):
    return _day_from_header(col)


def _header_score(df) -> float:
    cols = list(df.columns)
    score = 0.0
    if _guess_column(cols, FIELD_ALIASES["driver_id"]) != NONE_OPTION:
        score += 3
    n_days = len(_day_number_columns(cols))
    score += 3 if n_days >= 15 else (n_days / 15) * 3
    score += 1 - _unnamed_ratio(cols)
    return score


def _read_excel_smart(file_obj, sheet_name):
    """Reads one sheet, trying row 1, 2 and 3 as the header row and keeping
    the best one (ties keep the earlier row). Returns (df, header_row_number)."""
    best_df, best_row, best_score = None, 1, None
    for hdr in (0, 1, 2):
        try:
            file_obj.seek(0)
            df = pd.read_excel(file_obj, sheet_name=sheet_name, header=hdr)
        except Exception:  # noqa: BLE001
            if hdr == 0:
                raise
            continue
        df.columns = [str(c).strip() for c in df.columns]
        score = _header_score(df)
        if best_score is None or score > best_score + 1e-9:
            best_df, best_row, best_score = df, hdr + 1, score
    file_obj.seek(0)
    return best_df, best_row


def _extract_orders(df):
    cols = list(df.columns)
    id_col = _guess_column(cols, FIELD_ALIASES["driver_id"])
    orders_col = _guess_column(cols, FIELD_ALIASES["total_orders"])
    name_col = _guess_column(cols, FIELD_ALIASES["driver_name"])
    if id_col == NONE_OPTION:
        return {}, {}

    day_cols = _day_number_columns(cols)
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
        explicit = _clean_number_value(raw[orders_col], as_int=True) if orders_col != NONE_OPTION else 0
        day_total = sum(_clean_number_value(raw[c], as_int=True) for c in day_cols) if day_cols else 0
        row_total = explicit if (orders_col != NONE_OPTION and explicit > 0) else day_total
        out[driver_id] = out.get(driver_id, 0) + row_total
        if name_col != NONE_OPTION and not pd.isna(raw[name_col]):
            nm = str(raw[name_col]).strip()
            if nm:
                id_to_name[driver_id] = nm
    return out, id_to_name


# ---- a roster tab's own order total of 0 must not override real orders -----
if not getattr(_extract_roster, "_zero_total_guard", False):
    _orig_extract_roster = _extract_roster

    def _extract_roster(df, month_year=None):
        records = _orig_extract_roster(df, month_year)
        for rec in records.values():
            if not rec.get("roster_total_orders"):
                rec["roster_total_orders"] = None
        return records

    _extract_roster._zero_total_guard = True


# ---- keep an Excel import report on screen after the page reloads ----------
if not getattr(process_workbook_all_sheets, "_report_patch", False):
    _orig_process_workbook_all_sheets = process_workbook_all_sheets

    def process_workbook_all_sheets(uploaded_file, month_year, roster_tab_override=None):
        report = {"month": month_year, "tabs": [], "orders": []}
        try:
            ignored = {t.strip().lower() for t in ((_load_gsheet_config() or {}).get("ignored_tabs") or [])}
            xls = pd.ExcelFile(uploaded_file)
            for sheet_name in xls.sheet_names:
                try:
                    df, hdr_row = _read_excel_smart(uploaded_file, sheet_name)
                except Exception as exc:  # noqa: BLE001
                    report["tabs"].append({"Tab": sheet_name, "Detected as": f"read error: {exc}",
                                           "Header row": "-", "Rows": 0, "Day columns": 0, "Ignored": ""})
                    continue
                df = df.dropna(axis=0, how="all").reset_index(drop=True)
                df.columns = _dedupe_headers([str(c).strip() for c in df.columns])
                if df.empty:
                    report["tabs"].append({"Tab": sheet_name, "Detected as": "empty", "Header row": hdr_row,
                                           "Rows": 0, "Day columns": 0, "Ignored": ""})
                    continue
                kind = _classify_sheet(df)
                is_ignored = sheet_name.strip().lower() in ignored
                report["tabs"].append({"Tab": sheet_name, "Detected as": kind, "Header row": hdr_row,
                                       "Rows": len(df), "Day columns": len(_day_number_columns(df.columns)),
                                       "Ignored": "yes" if is_ignored else ""})
                if kind == "orders" and not is_ignored:
                    diag = _diagnose_orders_sheet(df)
                    diag["tab"] = sheet_name
                    diag["columns"] = [str(c) for c in df.columns]
                    report["orders"].append(diag)
        except Exception as exc:  # noqa: BLE001
            report["error"] = str(exc)
        summary = _orig_process_workbook_all_sheets(uploaded_file, month_year, roster_tab_override)
        report["summary"] = {k: summary.get(k) for k in (
            "roster_count", "orders_total", "orders_drivers", "validity_drivers",
            "attendance_drivers", "roster_sheet_used", "duplicate_orders_sheets")}
        st.session_state["excel_import_report"] = report
        return summary

    process_workbook_all_sheets._report_patch = True


def _render_excel_import_report():
    report = st.session_state.get("excel_import_report")
    if not report:
        return
    title = f"\U0001F9FE Last Excel import report ({month_display(report.get('month'))}) -- open this if a number looks wrong"
    with st.expander(title, expanded=True):
        s = report.get("summary") or {}
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Roster riders", s.get("roster_count") or 0)
        c2.metric("Total Orders", f"{(s.get('orders_total') or 0):,}")
        c3.metric("Riders with orders", s.get("orders_drivers") or 0)
        c4.metric("Validity data for", s.get("validity_drivers") or 0)
        if s.get("roster_sheet_used"):
            st.caption(f"Roster tab used: **{s['roster_sheet_used']}**")
        if s.get("duplicate_orders_sheets"):
            st.warning("Skipped as duplicate orders tab(s): " + ", ".join(s["duplicate_orders_sheets"]))
        if report.get("error"):
            st.warning(f"Report could not read everything: {report['error']}")

        st.markdown("##### Every tab -- how it was read")
        st.dataframe(pd.DataFrame(report.get("tabs", [])), use_container_width=True, hide_index=True)

        for diag in report.get("orders", []):
            st.markdown(f"##### Orders tab `{diag['tab']}`")
            m1, m2 = st.columns(2)
            if diag.get("orders_col_found"):
                m1.metric(f"Own '{diag['orders_col_name']}' column", f"{diag['orders_col_sum']:,}")
            else:
                m1.caption("No Total-Orders-style column on this tab.")
            m2.metric(f"Sum of {diag['day_cols_count']} day column(s)", f"{diag['day_cols_sum']:,}")
            if diag.get("rows_without_id"):
                st.warning(f"{diag['rows_without_id']} row(s) had no usable Courier/Driver ID and were skipped.")
            if diag.get("colliding_days"):
                st.error("Two or more columns resolve to the SAME day: "
                         + ", ".join(f"Day {d}" for d in diag["colliding_days"]))
            st.caption("Column headers read from this tab:")
            st.code(" | ".join(diag.get("columns", [])), language=None)

        if st.button("Clear this report", key="clear_excel_import_report"):
            st.session_state.pop("excel_import_report", None)
            st.rerun()


def render_upload_tab():
    if not is_admin():
        st.info(
            "\U0001F512 Only the HQ Admin can upload or import data. "
            "You're viewing this dashboard as a read-only Viewer."
        )
        return

    _render_excel_import_report()

    tab_ops, tab_salary, tab_gsheet, tab_offdays = st.tabs(
        ["\U0001F4E4 Roster / Orders / Validity", "\U0001F4B0 Salary Data",
         "\U0001F517 Google Sheet Sync", "\U0001F4C5 Weekly Off-Days"]
    )
    with tab_ops:
        _render_operations_upload()
    with tab_salary:
        _render_salary_upload()
    with tab_gsheet:
        render_gsheet_sync_tab()
    with tab_offdays:
        _render_off_days_upload()
