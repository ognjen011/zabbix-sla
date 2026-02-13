#!/usr/bin/env python3
"""
Streamlit frontend for Zabbix SLA Report Generator.
Includes authentication, role-based access, and report history.

Run with: streamlit run app.py
Default login: admin / admin
"""

import io
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import yaml

import database as db
from zabbix_sla_report import DateRangeCalculator, ExcelReportGenerator, ZabbixAPI

# --- Page config (must be first Streamlit call) ---
st.set_page_config(
    page_title="Zabbix Reporter",
    page_icon=":bar_chart:",
    layout="wide",
    menu_items={
        "Get Help": None,
        "Report a bug": None,
        "About": "Zabbix SLA Report Generator",
    },
)

# Initialize database
db.init_db()

# --- Load config defaults ---
CONFIG_PATH = Path("config.yaml")
default_config = {}
if CONFIG_PATH.exists():
    with open(CONFIG_PATH) as f:
        default_config = yaml.safe_load(f) or {}


# ============================================================
# Authentication
# ============================================================

def is_logged_in() -> bool:
    return st.session_state.get("authenticated", False)


def current_user() -> dict:
    return st.session_state.get("user", {})


def is_admin() -> bool:
    return current_user().get("role") == "admin"


def logout():
    for key in ["authenticated", "user"]:
        st.session_state.pop(key, None)
    st.rerun()


def show_login_page():
    """Display the login form."""
    # Center the login form - narrow column
    _, col, _ = st.columns([2, 1, 2])
    with col:
        st.title("Zabbix Reporter")
        with st.form("login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Login", type="primary", use_container_width=True)

        if submitted:
            if not username or not password:
                st.error("Please enter username and password.")
                return
            user = db.authenticate(username, password)
            if user:
                st.session_state["authenticated"] = True
                st.session_state["user"] = {
                    "id": user["id"],
                    "username": user["username"],
                    "role": user["role"],
                    "display_name": user["display_name"],
                }
                st.rerun()
            else:
                st.error("Invalid username or password.")


# ============================================================
# Restrict Streamlit menu for normal users
# ============================================================

def apply_role_css():
    """Hide top-right Streamlit menu items for non-admin users and style sidebar."""
    css = """
    <style>
        /* Push sidebar content into a flex column so logout stays at bottom */
        [data-testid="stSidebar"] > div:first-child {
            display: flex;
            flex-direction: column;
            height: 100vh;
        }
        [data-testid="stSidebarContent"] {
            display: flex;
            flex-direction: column;
            flex-grow: 1;
        }
    </style>
    """
    if not is_admin():
        css = """
        <style>
            [data-testid="stSidebar"] > div:first-child {
                display: flex;
                flex-direction: column;
                height: 100vh;
            }
            [data-testid="stSidebarContent"] {
                display: flex;
                flex-direction: column;
                flex-grow: 1;
            }
            [data-testid="stMainMenu"] {display: none;}
            .stDeployButton {display: none;}
            [data-testid="manage-app-button"] {display: none;}
        </style>
        """
    st.markdown(css, unsafe_allow_html=True)


# ============================================================
# Helper: SLA cell coloring
# ============================================================

def color_sla(val, threshold, orange_thresh):
    if isinstance(val, str):
        colors = {
            "COMPLIANT": "background-color: #C6EFCE; color: #006100",
            "WARNING": "background-color: #FFEB9C; color: #9C5700",
            "BREACH": "background-color: #FFC7CE; color: #9C0006",
        }
        return colors.get(val, "")
    try:
        v = float(val)
    except (ValueError, TypeError):
        return ""
    if v >= threshold:
        return "background-color: #C6EFCE; color: #006100"
    elif v >= threshold - orange_thresh:
        return "background-color: #FFEB9C; color: #9C5700"
    else:
        return "background-color: #FFC7CE; color: #9C0006"


# ============================================================
# Helper: build Excel bytes
# ============================================================

def build_excel_bytes(all_group_data, all_group_summaries, selected_groups, sla_threshold, orange_threshold, report_mode):
    """Build Excel report and return (filename, bytes) list."""
    results = []
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if report_mode == "combined":
        report = ExcelReportGenerator(sla_threshold, orange_threshold)
        for group_name, host_list in all_group_data.items():
            gc = selected_groups.get(group_name, {})
            grp_sla = gc.get("sla_threshold", sla_threshold)
            grp_orange = gc.get("orange_threshold", orange_threshold)
            report.sla_threshold = grp_sla
            report.orange_threshold = grp_orange
            report.create_sheet(group_name, host_list, grp_sla)
        report.add_summary_sheet(all_group_summaries)
        buf = io.BytesIO()
        report.workbook.save(buf)
        results.append((f"SLA_Report_{period}_{timestamp}.xlsx", buf.getvalue()))
    else:
        for group_name, host_list in all_group_data.items():
            if not host_list:
                continue
            gc = selected_groups.get(group_name, {})
            grp_sla = gc.get("sla_threshold", sla_threshold)
            grp_orange = gc.get("orange_threshold", orange_threshold)
            report = ExcelReportGenerator(grp_sla, grp_orange)
            report.create_sheet(group_name, host_list, grp_sla)
            summary_for_group = [s for s in all_group_summaries if s["group_name"] == group_name]
            report.add_summary_sheet(summary_for_group)
            buf = io.BytesIO()
            report.workbook.save(buf)
            safe_name = group_name.replace(" ", "_").replace("/", "-")
            results.append((f"SLA_Report_{safe_name}_{period}_{timestamp}.xlsx", buf.getvalue()))

    return results


# ============================================================
# Login gate
# ============================================================

if not is_logged_in():
    show_login_page()
    st.stop()

apply_role_css()
user = current_user()

# ============================================================
# Sidebar: Navigation + User info
# ============================================================

with st.sidebar:
    st.markdown(f"**{user['display_name'] or user['username']}** ({user['role']})")
    st.divider()

    nav_options = ["Generate Report", "Report History", "My Account"]
    if is_admin():
        nav_options.append("User Management")

    page = st.radio("Navigation", nav_options, label_visibility="collapsed")

    st.divider()
    with st.expander("Zabbix Connection", expanded=False):
        zabbix_url = st.text_input(
            "Zabbix URL",
            value=default_config.get("zabbix", {}).get("url", ""),
        )
        zabbix_token = st.text_input(
            "API Token",
            value=default_config.get("zabbix", {}).get("token", ""),
            type="password",
        )
        if st.button("Test Connection"):
            if not zabbix_url or not zabbix_token:
                st.error("Please enter Zabbix URL and API token.")
            else:
                try:
                    api = ZabbixAPI(zabbix_url, zabbix_token)
                    ver = api._call("apiinfo.version", use_auth=False)
                    st.success(f"Connected to Zabbix API v{ver}")
                except Exception as e:
                    st.error(f"Connection failed: {e}")

    st.button("Logout", on_click=logout, use_container_width=True)


# ============================================================
# Page: Generate Report
# ============================================================

if page == "Generate Report":

    st.title("Generate SLA Report")

    # --- SLA settings ---
    default_sla = default_config.get("default_sla_threshold", 99.9)
    default_orange = default_config.get("default_orange_threshold", 5.0)

    col_sla1, col_sla2, col_sla3, col_sla4 = st.columns(4)
    with col_sla1:
        sla_threshold = st.number_input(
            "SLA Threshold (%)",
            value=default_sla, min_value=0.0, max_value=100.0, step=0.01, format="%.2f",
            disabled=not is_admin(),
            help="Only admins can change SLA thresholds" if not is_admin() else None,
        )
    with col_sla2:
        orange_threshold = st.number_input(
            "Warning Threshold (%)",
            value=default_orange, min_value=0.0, max_value=100.0, step=0.1, format="%.1f",
            disabled=not is_admin(),
            help="Only admins can change thresholds" if not is_admin() else None,
        )
    with col_sla3:
        period = st.selectbox(
            "SLA Period",
            options=["month", "week", "day"],
            format_func=lambda x: {"month": "Previous Month", "week": "Last 7 Days", "day": "Last 24 Hours"}[x],
        )
    with col_sla4:
        report_mode = st.selectbox(
            "Report Mode",
            options=["combined", "separate"],
            format_func=lambda x: {"combined": "Combined (one file)", "separate": "Separate (per group)"}[x],
            index=0 if default_config.get("report_mode", "combined") == "combined" else 1,
        )

    # --- Host group selection ---
    st.subheader("Host Groups")
    config_groups = list((default_config.get("host_groups", {}) or {}).keys())

    group_source = st.radio(
        "Select host groups from:",
        ["Config file", "Fetch from Zabbix", "Manual entry"],
        horizontal=True,
    )

    selected_groups = {}

    def get_zabbix_api():
        if not zabbix_url or not zabbix_token:
            return None
        try:
            api = ZabbixAPI(zabbix_url, zabbix_token)
            api._call("apiinfo.version", use_auth=False)
            return api
        except Exception:
            return None

    if group_source == "Config file":
        if config_groups:
            chosen = st.multiselect("Host Groups", options=config_groups, default=config_groups)
            host_groups_config = default_config.get("host_groups", {})
            for g in chosen:
                gc = host_groups_config.get(g, {}) or {}
                selected_groups[g] = {
                    "sla_threshold": gc.get("sla_threshold", sla_threshold),
                    "orange_threshold": gc.get("orange_threshold", orange_threshold),
                    "excluded_hosts": gc.get("excluded_hosts", []) or [],
                }
        else:
            st.warning("No host groups found in config.yaml")

    elif group_source == "Fetch from Zabbix":
        api = get_zabbix_api()
        if api:
            try:
                all_groups = api.get_host_groups()
                group_names = sorted([g["name"] for g in all_groups])
                chosen = st.multiselect(
                    "Host Groups", options=group_names,
                    default=[g for g in config_groups if g in group_names],
                )
                for g in chosen:
                    selected_groups[g] = {
                        "sla_threshold": sla_threshold,
                        "orange_threshold": orange_threshold,
                        "excluded_hosts": [],
                    }
            except Exception as e:
                st.error(f"Failed to fetch groups: {e}")
        else:
            st.warning("Connect to Zabbix first.")

    else:
        manual_groups = st.text_area("Enter host group names (one per line)", height=100)
        for g in manual_groups.strip().split("\n"):
            g = g.strip()
            if g:
                selected_groups[g] = {
                    "sla_threshold": sla_threshold,
                    "orange_threshold": orange_threshold,
                    "excluded_hosts": [],
                }

    # Per-group settings (admin only can edit)
    if selected_groups and is_admin():
        with st.expander("Per-group threshold overrides", expanded=False):
            for g in list(selected_groups.keys()):
                c1, c2, c3 = st.columns([2, 1, 1])
                with c1:
                    st.markdown(f"**{g}**")
                with c2:
                    selected_groups[g]["sla_threshold"] = st.number_input(
                        f"SLA % ({g})", value=selected_groups[g]["sla_threshold"],
                        min_value=0.0, max_value=100.0, step=0.01, format="%.2f",
                        key=f"sla_{g}", label_visibility="collapsed",
                    )
                with c3:
                    selected_groups[g]["orange_threshold"] = st.number_input(
                        f"Warn % ({g})", value=selected_groups[g]["orange_threshold"],
                        min_value=0.0, max_value=100.0, step=0.1, format="%.1f",
                        key=f"orange_{g}", label_visibility="collapsed",
                    )

    # Excluded hosts
    global_excluded_str = st.text_area(
        "Global Excluded Hosts (one per line)",
        value="\n".join(default_config.get("global_excluded_hosts", []) or []),
        height=68,
        disabled=not is_admin(),
    )
    global_excluded = [h.strip() for h in global_excluded_str.strip().split("\n") if h.strip()]

    # --- Generate button (right-aligned) ---
    st.divider()
    _, btn_col = st.columns([3, 1])
    with btn_col:
        generate_clicked = st.button("Generate Report", type="primary", use_container_width=True)
    if generate_clicked:
        st.session_state["trigger_generate"] = True
    generate_btn = st.session_state.pop("trigger_generate", False)

    if generate_btn and selected_groups:
        if not zabbix_url or not zabbix_token:
            st.error("Please configure Zabbix connection.")
            st.stop()

        try:
            api = ZabbixAPI(zabbix_url, zabbix_token)
            api._call("apiinfo.version", use_auth=False)
        except Exception as e:
            st.error(f"Cannot connect to Zabbix: {e}")
            st.stop()

        date_calc = DateRangeCalculator()
        availability_periods = date_calc.get_availability_periods()
        global_excluded_lower = [h.lower() for h in global_excluded]

        group_names_list = list(selected_groups.keys())
        groups = api.get_host_groups(group_names_list)

        if not groups:
            st.error(f"No matching host groups found in Zabbix: {group_names_list}")
            st.stop()

        found_names = {g["name"] for g in groups}
        missing = set(group_names_list) - found_names
        if missing:
            st.warning(f"Groups not found in Zabbix: {', '.join(missing)}")

        progress = st.progress(0, text="Starting...")
        all_group_summaries = []
        all_group_data = {}
        total_hosts_processed = 0

        total_hosts = 0
        group_hosts_map = {}
        for group in groups:
            hosts = api.get_hosts_in_group(group["groupid"])
            group_hosts_map[group["groupid"]] = hosts
            total_hosts += len(hosts)

        if total_hosts == 0:
            progress.empty()
            st.warning("No enabled hosts found in the selected groups.")
            st.stop()

        for group in groups:
            group_name = group["name"]
            group_id = group["groupid"]
            gc = selected_groups.get(group_name, {})
            grp_sla = gc.get("sla_threshold", sla_threshold)
            grp_orange = gc.get("orange_threshold", orange_threshold)
            grp_excluded = gc.get("excluded_hosts", []) or []
            grp_excluded_lower = [h.lower() for h in grp_excluded]
            all_excluded_lower = global_excluded_lower + grp_excluded_lower

            hosts = group_hosts_map[group_id]
            host_data_list = []
            summary = {
                "group_name": group_name,
                "sla_threshold": grp_sla,
                "total": 0, "compliant": 0, "warning": 0, "breach": 0,
            }

            for host in hosts:
                host_id = host["hostid"]
                host_name = host["name"]
                host_technical = host["host"]

                total_hosts_processed += 1
                pct = total_hosts_processed / total_hosts
                progress.progress(pct, text=f"Processing: {host_name} ({group_name})")

                if host_name.lower() in all_excluded_lower or host_technical.lower() in all_excluded_lower:
                    continue

                avail_1_day = api.get_host_availability(
                    host_id,
                    int(availability_periods["1_day"][0].timestamp()),
                    int(availability_periods["1_day"][1].timestamp()),
                )
                avail_7_days = api.get_host_availability(
                    host_id,
                    int(availability_periods["7_days"][0].timestamp()),
                    int(availability_periods["7_days"][1].timestamp()),
                )
                avail_prev_month = api.get_host_availability(
                    host_id,
                    int(availability_periods["prev_month"][0].timestamp()),
                    int(availability_periods["prev_month"][1].timestamp()),
                )

                if period == "day":
                    device_sla = avail_1_day["availability"]
                elif period == "week":
                    device_sla = avail_7_days["availability"]
                else:
                    device_sla = avail_prev_month["availability"]

                if device_sla >= grp_sla:
                    status = "COMPLIANT"
                    summary["compliant"] += 1
                elif device_sla >= grp_sla - grp_orange:
                    status = "WARNING"
                    summary["warning"] += 1
                else:
                    status = "BREACH"
                    summary["breach"] += 1
                summary["total"] += 1

                host_data_list.append({
                    "name": host_name,
                    "host": host_technical,
                    "avail_1_day": avail_1_day["availability"],
                    "avail_7_days": avail_7_days["availability"],
                    "avail_prev_month": avail_prev_month["availability"],
                    "device_sla": device_sla,
                    "sla_status": status,
                    "downtime_1_day": avail_1_day["downtime_seconds"],
                    "downtime_7_days": avail_7_days["downtime_seconds"],
                    "downtime_prev_month": avail_prev_month["downtime_seconds"],
                    "total_1_day": avail_1_day["total_seconds"],
                    "total_7_days": avail_7_days["total_seconds"],
                    "total_prev_month": avail_prev_month["total_seconds"],
                })

            if host_data_list:
                td1 = sum(h["downtime_1_day"] for h in host_data_list)
                td7 = sum(h["downtime_7_days"] for h in host_data_list)
                tdm = sum(h["downtime_prev_month"] for h in host_data_list)
                tp1 = sum(h["total_1_day"] for h in host_data_list)
                tp7 = sum(h["total_7_days"] for h in host_data_list)
                tpm = sum(h["total_prev_month"] for h in host_data_list)

                o1 = ((tp1 - td1) / tp1 * 100) if tp1 > 0 else 100.0
                o7 = ((tp7 - td7) / tp7 * 100) if tp7 > 0 else 100.0
                om = ((tpm - tdm) / tpm * 100) if tpm > 0 else 100.0
                overall_sla = {"day": o1, "week": o7, "month": om}[period]
                summary.update({
                    "overall_1_day": round(o1, 2),
                    "overall_7_days": round(o7, 2),
                    "overall_prev_month": round(om, 2),
                    "overall_sla": round(overall_sla, 2),
                })
            else:
                summary.update({
                    "overall_1_day": 100.0, "overall_7_days": 100.0,
                    "overall_prev_month": 100.0, "overall_sla": 100.0,
                })

            all_group_summaries.append(summary)
            all_group_data[group_name] = host_data_list

        progress.progress(1.0, text="Done!")
        time.sleep(0.3)
        progress.empty()

        # Store results in session state so they survive reruns (e.g. save button click)
        excel_files = build_excel_bytes(
            all_group_data, all_group_summaries, selected_groups,
            sla_threshold, orange_threshold, report_mode,
        )

        # Build detail for storage
        detail_for_storage = {}
        for gn, hlist in all_group_data.items():
            detail_for_storage[gn] = [{
                "name": h["name"], "host": h["host"],
                "avail_1_day": h["avail_1_day"],
                "avail_7_days": h["avail_7_days"],
                "avail_prev_month": h["avail_prev_month"],
                "device_sla": h["device_sla"],
                "sla_status": h["sla_status"],
            } for h in hlist]

        st.session_state["last_report"] = {
            "all_group_summaries": all_group_summaries,
            "all_group_data": all_group_data,
            "selected_groups": selected_groups,
            "sla_threshold": sla_threshold,
            "orange_threshold": orange_threshold,
            "period": period,
            "excel_files": excel_files,
            "detail_for_storage": detail_for_storage,
            "total_host_count": sum(s["total"] for s in all_group_summaries),
        }

    # --- Display results from session state ---
    rpt_data = st.session_state.get("last_report")
    if rpt_data:
        all_group_summaries = rpt_data["all_group_summaries"]
        all_group_data = rpt_data["all_group_data"]
        r_selected_groups = rpt_data["selected_groups"]
        r_sla = rpt_data["sla_threshold"]
        r_orange = rpt_data["orange_threshold"]
        r_period = rpt_data["period"]
        excel_files = rpt_data["excel_files"]
        detail_for_storage = rpt_data["detail_for_storage"]
        total_host_count = rpt_data["total_host_count"]

        st.success(f"Report generated for {total_host_count} hosts across {len(all_group_summaries)} groups.")

        # Summary table
        st.subheader("Summary")
        summary_rows = []
        for s in all_group_summaries:
            summary_rows.append({
                "Group": s["group_name"],
                "SLA Target (%)": s["sla_threshold"],
                "Hosts": s["total"],
                "Compliant": s["compliant"],
                "Warning": s["warning"],
                "Breach": s["breach"],
                "SLA 1 Day (%)": s.get("overall_1_day", 100.0),
                "SLA 7 Days (%)": s.get("overall_7_days", 100.0),
                "SLA Prev Month (%)": s.get("overall_prev_month", 100.0),
                "Overall SLA (%)": s.get("overall_sla", 100.0),
            })

        summary_df = pd.DataFrame(summary_rows)
        sla_cols = ["SLA 1 Day (%)", "SLA 7 Days (%)", "SLA Prev Month (%)", "Overall SLA (%)"]

        def style_summary(row):
            styles = [""] * len(row)
            target = row["SLA Target (%)"]
            for col in sla_cols:
                idx = summary_df.columns.get_loc(col)
                styles[idx] = color_sla(row[col], target, r_orange)
            return styles

        styled_summary = summary_df.style.apply(style_summary, axis=1).format(
            {c: "{:.2f}" for c in sla_cols + ["SLA Target (%)"]},
        )
        st.dataframe(styled_summary, use_container_width=True, hide_index=True)

        # Per-group tables
        for group_name, host_list in all_group_data.items():
            if not host_list:
                continue
            gc = r_selected_groups.get(group_name, {})
            grp_sla = gc.get("sla_threshold", r_sla)
            grp_orange = gc.get("orange_threshold", r_orange)

            st.subheader(group_name)
            display_data = [{
                "Host Name": h["name"],
                "Host": h["host"],
                "1 Day (%)": h["avail_1_day"],
                "7 Days (%)": h["avail_7_days"],
                "Prev Month (%)": h["avail_prev_month"],
                "Device SLA (%)": h["device_sla"],
                "SLA Target (%)": grp_sla,
                "Status": h["sla_status"],
            } for h in host_list]

            df = pd.DataFrame(display_data)
            avail_cols = ["1 Day (%)", "7 Days (%)", "Prev Month (%)", "Device SLA (%)"]

            _grp_sla = grp_sla
            _grp_orange = grp_orange

            def style_group_row(row, _sla=_grp_sla, _orange=_grp_orange):
                styles = [""] * len(row)
                for col in avail_cols:
                    idx = df.columns.get_loc(col)
                    styles[idx] = color_sla(row[col], _sla, _orange)
                status_idx = df.columns.get_loc("Status")
                styles[status_idx] = color_sla(row["Status"], _sla, _orange)
                return styles

            styled_df = df.style.apply(style_group_row, axis=1).format(
                {c: "{:.2f}" for c in avail_cols + ["SLA Target (%)"]},
            )
            st.dataframe(styled_df, use_container_width=True, hide_index=True)

        # --- Excel download + save to history ---
        st.divider()
        st.subheader("Download & Save")

        for filename, excel_bytes in excel_files:
            col_dl, col_save = st.columns([1, 1])
            with col_dl:
                st.download_button(
                    label=f"Download {filename}",
                    data=excel_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{filename}",
                )
            with col_save:
                if st.button(f"Save to history", key=f"save_{filename}"):
                    db.save_report(
                        generated_by=user["username"],
                        report_name=filename,
                        period=r_period,
                        groups_list=list(all_group_data.keys()),
                        host_count=total_host_count,
                        summary_data=all_group_summaries,
                        detail_data=detail_for_storage,
                        excel_data=excel_bytes,
                    )
                    st.success("Report saved to history.")


# ============================================================
# Page: Report History
# ============================================================

elif page == "Report History":

    st.title("Report History")

    report_count = db.get_report_count()
    if report_count == 0:
        st.info("No reports saved yet. Generate a report and click 'Save to history'.")
    else:
        st.caption(f"{report_count} report(s) in history")

        reports = db.get_reports(limit=100)

        for rpt in reports:
            with st.expander(
                f"{rpt['report_name']}  |  {rpt['generated_at']}  |  by {rpt['generated_by']}  |  {rpt['host_count']} hosts",
                expanded=False,
            ):
                col_info, col_actions = st.columns([3, 1])

                with col_info:
                    st.markdown(f"**Period:** {rpt['period']}  \n"
                                f"**Groups:** {', '.join(rpt['groups_list'])}  \n"
                                f"**Generated:** {rpt['generated_at']}  \n"
                                f"**By:** {rpt['generated_by']}")

                    # Show summary table if available
                    if rpt["summary_data"]:
                        summary_rows = []
                        for s in rpt["summary_data"]:
                            summary_rows.append({
                                "Group": s.get("group_name", ""),
                                "SLA Target (%)": s.get("sla_threshold", 99.9),
                                "Hosts": s.get("total", 0),
                                "Compliant": s.get("compliant", 0),
                                "Warning": s.get("warning", 0),
                                "Breach": s.get("breach", 0),
                                "Overall SLA (%)": s.get("overall_sla", 100.0),
                            })
                        if summary_rows:
                            sdf = pd.DataFrame(summary_rows)

                            def style_hist_row(row):
                                styles = [""] * len(row)
                                target = row["SLA Target (%)"]
                                idx = sdf.columns.get_loc("Overall SLA (%)")
                                styles[idx] = color_sla(row["Overall SLA (%)"], target, 5.0)
                                return styles

                            styled = sdf.style.apply(style_hist_row, axis=1).format(
                                {"SLA Target (%)": "{:.2f}", "Overall SLA (%)": "{:.2f}"},
                            )
                            st.dataframe(styled, use_container_width=True, hide_index=True)

                with col_actions:
                    # Download Excel
                    excel_data = db.get_report_excel(rpt["id"])
                    if excel_data:
                        st.download_button(
                            label="Download Excel",
                            data=excel_data,
                            file_name=rpt["report_name"],
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"hist_dl_{rpt['id']}",
                        )

                    # View detail
                    view_details = st.button("View Details", key=f"hist_view_{rpt['id']}")

                    # Delete (admin only)
                    if is_admin():
                        if st.button("Delete", key=f"hist_del_{rpt['id']}", type="secondary"):
                            db.delete_report(rpt["id"])
                            st.rerun()

                # Render detail tables at full width (outside columns)
                if view_details:
                    full_report = db.get_report(rpt["id"])
                    if full_report and full_report.get("detail_data"):
                        for gn, hlist in full_report["detail_data"].items():
                            st.markdown(f"**{gn}**")
                            if hlist:
                                hdf = pd.DataFrame(hlist)
                                st.dataframe(hdf, use_container_width=True, hide_index=True)


# ============================================================
# Page: My Account
# ============================================================

elif page == "My Account":

    st.title("My Account")

    st.markdown(f"**Username:** {user['username']}  \n"
                f"**Role:** {user['role']}  \n"
                f"**Display name:** {user['display_name']}")

    st.divider()
    st.subheader("Change Password")

    with st.form("change_pw_form"):
        old_pw = st.text_input("Current Password", type="password")
        new_pw = st.text_input("New Password", type="password")
        confirm_pw = st.text_input("Confirm New Password", type="password")
        submitted = st.form_submit_button("Change Password")

    if submitted:
        if not old_pw or not new_pw:
            st.error("Please fill in all fields.")
        elif new_pw != confirm_pw:
            st.error("New passwords do not match.")
        elif len(new_pw) < 4:
            st.error("Password must be at least 4 characters.")
        else:
            if db.change_password(user["id"], old_pw, new_pw):
                st.success("Password changed successfully.")
            else:
                st.error("Current password is incorrect.")


# ============================================================
# Page: User Management (admin only)
# ============================================================

elif page == "User Management":

    if not is_admin():
        st.error("Access denied.")
        st.stop()

    st.title("User Management")

    # --- Create user ---
    st.subheader("Create New User")
    with st.form("create_user_form"):
        col1, col2 = st.columns(2)
        with col1:
            new_username = st.text_input("Username")
            new_display = st.text_input("Display Name")
        with col2:
            new_password = st.text_input("Password", type="password")
            new_role = st.selectbox("Role", ["user", "admin"])
        create_submitted = st.form_submit_button("Create User", type="primary")

    if create_submitted:
        if not new_username or not new_password:
            st.error("Username and password are required.")
        elif len(new_password) < 4:
            st.error("Password must be at least 4 characters.")
        else:
            uid = db.create_user(new_username, new_password, new_role, new_display or new_username)
            if uid:
                st.success(f"User '{new_username}' created.")
                st.rerun()
            else:
                st.error(f"Username '{new_username}' already exists.")

    # --- Existing users ---
    st.divider()
    st.subheader("Existing Users")

    users_list = db.get_all_users()
    if not users_list:
        st.info("No users found.")
    else:
        user_df = pd.DataFrame(users_list)
        user_df = user_df.rename(columns={
            "id": "ID", "username": "Username", "role": "Role",
            "display_name": "Display Name", "created_at": "Created",
        })
        st.dataframe(user_df, use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("Edit User")

        edit_options = {f"{u['username']} (ID: {u['id']})": u["id"] for u in users_list}
        selected_user_label = st.selectbox("Select user", list(edit_options.keys()))
        selected_user_id = edit_options[selected_user_label]
        selected_user_data = next(u for u in users_list if u["id"] == selected_user_id)

        with st.form("edit_user_form"):
            e_col1, e_col2 = st.columns(2)
            with e_col1:
                edit_display = st.text_input("Display Name", value=selected_user_data["display_name"])
                edit_role = st.selectbox(
                    "Role", ["user", "admin"],
                    index=0 if selected_user_data["role"] == "user" else 1,
                )
            with e_col2:
                edit_password = st.text_input("New Password (leave blank to keep)", type="password")

            e_col_save, e_col_del = st.columns(2)
            with e_col_save:
                edit_submitted = st.form_submit_button("Save Changes", type="primary")
            with e_col_del:
                delete_submitted = st.form_submit_button("Delete User")

        if edit_submitted:
            # Prevent removing the last admin
            if selected_user_data["role"] == "admin" and edit_role == "user":
                admin_count = sum(1 for u in users_list if u["role"] == "admin")
                if admin_count <= 1:
                    st.error("Cannot remove the last admin user.")
                    st.stop()

            db.update_user(
                selected_user_id,
                display_name=edit_display,
                role=edit_role,
                password=edit_password if edit_password else None,
            )
            st.success("User updated.")
            st.rerun()

        if delete_submitted:
            if selected_user_id == user["id"]:
                st.error("You cannot delete your own account.")
            else:
                admin_count = sum(1 for u in users_list if u["role"] == "admin")
                if selected_user_data["role"] == "admin" and admin_count <= 1:
                    st.error("Cannot delete the last admin user.")
                else:
                    db.delete_user(selected_user_id)
                    st.success(f"User '{selected_user_data['username']}' deleted.")
                    st.rerun()
