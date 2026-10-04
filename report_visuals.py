"""SLA gauges and per-interface interactive trend charts."""
import html
import math

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


def gauge_html(label, value, target, warning):
    value = max(0, min(100, float(value)))
    color = "#16803c" if value >= target else "#d58a00" if value >= target - warning else "#c53030"
    angle = math.pi * (1 - value / 100)
    x, y = 150 + 100 * math.cos(angle), 125 - 100 * math.sin(angle)
    target_angle = math.pi * (1 - target / 100)
    tx, ty = 150 + 100 * math.cos(target_angle), 125 - 100 * math.sin(target_angle)
    return f'''<div style="font-family:system-ui;text-align:center;color:#334155">
    <div>{html.escape(str(label))}</div>
    <svg viewBox="0 0 300 170" style="width:100%;max-width:360px" role="img" aria-label="SLA {value:.3f} percent; target {target:.3f} percent">
    <path d="M50 125 A100 100 0 0 1 250 125" fill="none" stroke="#e2e8f0" stroke-width="18"/>
    <path d="M50 125 A100 100 0 0 1 {x:.2f} {y:.2f}" fill="none" stroke="{color}" stroke-width="18"/>
    <circle cx="{tx:.2f}" cy="{ty:.2f}" r="5" fill="#334155"/>
    <text x="150" y="117" text-anchor="middle" font-size="27" fill="{color}">{value:.3f}%</text>
    <text x="150" y="150" text-anchor="middle" font-size="13">Target: {target:.3f}%</text></svg></div>'''


def show_sla_gauges(summaries, group_settings, default_warning, period):
    st.subheader("SLA Gauges")
    for summary in summaries:
        group = summary["group_name"]
        st.markdown(f"**{group}**")
        warning = group_settings.get(group, {}).get("orange_threshold", default_warning)
        gauges = [{"label": row["Month"], "value": row.get("Availability (%)"),
                   "target": row.get("SLA Target (%)", summary["sla_threshold"])}
                  for row in sorted(summary.get("sla_history", []), key=lambda row: row["Month"])]
        if not gauges or period in ("day", "week"):
            gauges.append({"label": {"month": "Previous calendar month", "week": "Last 7 days", "day": "Last 24 hours"}.get(period, period),
                           "value": summary.get("overall_sla"), "target": summary["sla_threshold"]})
        for offset in range(0, len(gauges), 3):
            columns = st.columns(min(3, len(gauges) - offset))
            for column, gauge in zip(columns, gauges[offset:offset + 3]):
                with column:
                    if gauge["value"] is None:
                        st.info(f"{gauge['label']}: no availability data")
                    else:
                        components.html(gauge_html(gauge["label"], gauge["value"], gauge["target"], warning), height=215)


def line_chart(rows, date_field, metrics, y_title, key, zero=True):
    data = pd.DataFrame(rows).reindex(columns=[date_field, *metrics])
    if not any(data[metric].notna().any() for metric in metrics):
        st.info("No retained data is available for this chart.")
        return
    long = data[[date_field, *metrics]].melt(id_vars=[date_field], var_name="Metric", value_name="Value")
    # Vega-Lite leaves gaps for missing values; never substitute zeros.
    monthly = date_field == "Month"
    bar = monthly and any(metric.endswith("(GB)") for metric in metrics)
    encoding = {
        "x": {"field": date_field, "type": "ordinal", "sort": "ascending", "title": date_field,
              "axis": {"labelAngle": 0}},
        "y": {"field": "Value", "type": "quantitative", "title": y_title, "scale": {"zero": zero}},
        "color": {"field": "Metric", "type": "nominal", "title": None},
        "tooltip": [{"field": date_field, "type": "ordinal"}, {"field": "Metric"},
                    {"field": "Value", "type": "quantitative", "format": ".3f"}],
    }
    if bar:
        encoding["xOffset"] = {"field": "Metric"}
    st.vega_lite_chart(long, {
        "mark": {"type": "bar"} if bar else {"type": "line", "point": True},
        "encoding": encoding,
    }, use_container_width=True, key=key)


def interface_identity(row):
    """Legacy snapshots used the interface name before IDs were retained."""
    return row.get("Interface ID") or row.get("Interface")


def show_link_trends(all_group_data, key_prefix=""):
    def widget_key(name):
        return f"{key_prefix}{name}"
    st.subheader("Link Traffic and Capacity Trends")
    choices = {}
    for group, hosts in all_group_data.items():
        for host_index, host in enumerate(hosts):
            for row in host.get("link_history") or host.get("link_usage", []):
                identity = interface_identity(row)
                if identity is not None:
                    choices[(group, host_index, identity)] = f"{group} / {host['name']} / {row.get('Interface', identity)} — {row.get('Description', '')}"
    if not choices:
        st.info("Generate a report with matching interfaces to view trends.")
        return
    selected = st.selectbox("Interface", list(choices), format_func=choices.get, key=widget_key("trend_interface"))
    group, host_index, identity = selected
    host = all_group_data[group][host_index]
    monthly = [r for r in host.get("link_history", []) if interface_identity(r) == identity]
    daily = [r for r in host.get("link_daily", []) if interface_identity(r) == identity]
    if not monthly:
        st.info("This saved report contains monthly totals but no historical trends. Generate a new report to collect the configured history window.")
        return
    monthly = sorted(monthly, key=lambda row: row["Month"])
    st.caption(f"{len(monthly)} monthly points: {monthly[0]['Month']} to {monthly[-1]['Month']}. Volumes are estimated from hourly trends; missing data stays blank. Download/upload are relative to the monitored interface.")
    traffic_tab, peak_tab, growth_tab = st.tabs(["Traffic trends", "Peak usage trends", "Capacity and demand growth"])
    with traffic_tab:
        resolution = st.radio("Traffic interval", ["Daily", "Monthly"], horizontal=True, key=widget_key("traffic_interval"))
        rows, date_field = (daily, "Date") if resolution == "Daily" else (monthly, "Month")
        if rows:
            line_chart(rows, date_field, ["Download (GB)", "Upload (GB)"], "Traffic (GB)", widget_key("traffic_chart"))
        st.dataframe(pd.DataFrame(monthly).reindex(columns=["Month", "Download (GB)", "Upload (GB)", "Download Trend Coverage (%)", "Upload Trend Coverage (%)"]), hide_index=True, use_container_width=True)
    with peak_tab:
        resolution = st.radio("Peak interval", ["Daily", "Monthly"], horizontal=True, key=widget_key("peak_interval"))
        rows, date_field = (daily, "Date") if resolution == "Daily" else (monthly, "Month")
        if rows:
            line_chart(rows, date_field, ["Download Peak (Mbps)", "Upload Peak (Mbps)"], "Peak rate (Mbps)", widget_key("peak_chart"))
        st.caption("Each point is the maximum observed sample for that day or month, not the hourly average.")
    with growth_tab:
        st.markdown("**Traffic demand**")
        line_chart(monthly, "Month", ["Download Avg (Mbps)", "Upload Avg (Mbps)", "Download Peak (Mbps)", "Upload Peak (Mbps)"], "Demand (Mbps)", widget_key("demand_chart"))
        st.markdown("**Reported interface capacity**")
        line_chart(monthly, "Month", ["Historical Capacity (Mbps)"], "Historical speed (Mbps)", widget_key("capacity_chart"))
        st.caption("Capacity is the monthly average reported interface speed. A service limit may differ from port speed. Historical speed gaps are not filled with current capacity. Growth is compared with the immediately preceding month and requires sufficient coverage; a zero baseline has no percentage change.")
        columns = ["Month", "Historical Capacity (Mbps)", "Capacity (Mbps)", "Capacity Basis", "Capacity Trend Coverage (%)", "Capacity Growth (%)", "Download Avg Growth (%)", "Upload Avg Growth (%)", "Download Peak Growth (%)", "Upload Peak Growth (%)"]
        st.dataframe(pd.DataFrame(monthly).reindex(columns=columns), hide_index=True, use_container_width=True)


def show_sla_trends(summaries, all_group_data, key_prefix=""):
    st.subheader("Monthly SLA Trends")
    choices = {}
    for summary in summaries:
        if summary.get("sla_history"):
            choices[f"Group: {summary['group_name']}"] = summary['sla_history']
    for group, hosts in all_group_data.items():
        for host in hosts:
            if host.get('sla_history'):
                choices[f"Host: {group} / {host['name']} ({host.get('host', '')})"] = host['sla_history']
    if not choices:
        st.info("This saved report does not contain monthly SLA history. Generate a new report to collect it.")
        return
    scope = st.selectbox("SLA scope", list(choices), key=f"{key_prefix}sla_scope")
    rows = sorted(choices[scope], key=lambda row: row['Month'])
    st.caption("ICMP availability for each complete calendar month, compared with the SLA target.")
    line_chart(rows, "Month", ["Availability (%)", "SLA Target (%)"], "Availability (%)", f"{key_prefix}sla_trend_chart", zero=False)
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
