# Zabbix SLA Report Generator

Generate Excel SLA reports from Zabbix availability data based on ICMP ping monitoring.

## License

This project is licensed under the [Apache License 2.0](LICENSE).
Third-party dependencies retain their own licenses.

## Features

- Pull availability data from Zabbix API
- Support multiple host groups with individual SLA thresholds
- Generate separate or combined Excel reports
- Color-coded compliance status (Green/Orange/Red)
- Exclude specific hosts globally or per-group
- Filter only "Unavailable by ICMP ping" High severity events

## Requirements

- Python 3.8+
- Zabbix 6.0+ with API access
- API Token with read permissions

## Installation

```bash
# Create virtual environment
python3 -m venv venv
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

## Configuration

Edit `config.yaml` to configure your environment:

### Zabbix Connection

```yaml
zabbix:
  url: "http://your-zabbix-server/zabbix"
  token: "your-api-token"
```

### Host Groups with Individual SLA Thresholds

```yaml
host_groups:
  "Group A":
    sla_threshold: 99.99      # SLA target percentage
    orange_threshold: 5.0     # Warning if within 5% of target
    excluded_hosts:           # Hosts to skip (this group only)
      - "test-device-1"

  "Group B":
    sla_threshold: 99.9
    orange_threshold: 5.0
    excluded_hosts: []

  "Critical Infrastructure":
    sla_threshold: 99.999
    orange_threshold: 2.0
```

### Global Exclusions

```yaml
# Excluded from ALL groups
global_excluded_hosts:
  - "test-server-01"
  - "monitoring-probe"
```

### Report Mode

```yaml
# "separate" = One Excel file per group
# "combined" = All groups in one Excel file with separate sheets
report_mode: "separate"
```

## Usage

### Basic Usage

```bash
# Activate virtual environment
source venv/bin/activate

# Generate SLA report based on previous month (default)
python zabbix_sla_report.py --period month

# Generate SLA report based on last 7 days
python zabbix_sla_report.py --period week

# Generate SLA report based on last 1 day
python zabbix_sla_report.py --period day
```

### Command Line Options

| Option | Description |
|--------|-------------|
| `--period`, `-p` | SLA period: `day`, `week`, or `month` (default: month) |
| `--config`, `-c` | Path to config file (default: config.yaml) |
| `--output`, `-o` | Output file path (auto-generated if not specified) |
| `--groups`, `-g` | Override groups from config (space-separated) |

### Examples

```bash
# Use custom config file
python zabbix_sla_report.py --config /path/to/config.yaml --period month

# Override groups from command line
python zabbix_sla_report.py --period week --groups "Group A" "Group B"

# Specify output file
python zabbix_sla_report.py --period month --output /path/to/report.xlsx
```

## Output

### Report Files

With `report_mode: "separate"`:
```
SLA_Report_GroupA_month_20260129_103753.xlsx
SLA_Report_GroupB_month_20260129_103753.xlsx
```

With `report_mode: "combined"`:
```
SLA_Report_month_20260129_103753.xlsx
```

### Excel Structure

**Summary Sheet:**
- Host Group name
- SLA Target (%)
- Total Hosts, Compliant, Warning, Breach counts
- Overall SLA for 1 Day, 7 Days, Prev Month
- Overall Group SLA (%)
- SLA Status

**Per-Group Sheet:**
- Host Name & Technical Host
- Availability 1 Day (%)
- Availability 7 Days (%)
- Availability Prev Month (%)
- Device SLA (%) - based on selected period
- SLA Target (%)
- SLA Status (COMPLIANT/WARNING/BREACH)
- Overall Group SLA row at bottom

### Color Coding

| Color | Condition |
|-------|-----------|
| Green | >= SLA threshold (COMPLIANT) |
| Orange | >= SLA threshold - orange_threshold (WARNING) |
| Red | < SLA threshold - orange_threshold (BREACH) |

## SLA Calculation

### Events Counted

Only **High severity (level 4)** problems with name containing:
- "Unavailable by ICMP ping"
- "ICMP ping"

### Availability Formula

```
Availability % = ((Total Seconds - Downtime Seconds) / Total Seconds) * 100
```

### Periods

| Period | Description |
|--------|-------------|
| 1 Day | Yesterday (last 24 hours) |
| 7 Days | Last 7 days |
| Prev Month | Previous calendar month (actual 28-31 days) |

### Device SLA

The **Device SLA** column uses the period specified by `--period`:
- `--period day` → Device SLA = 1 Day availability
- `--period week` → Device SLA = 7 Days availability
- `--period month` → Device SLA = Prev Month availability

## Troubleshooting

### Connection Error

```
Error connecting to Zabbix: ...
```
- Verify Zabbix URL is correct
- Check API token has proper permissions
- Ensure network connectivity to Zabbix server

### No Host Groups Found

```
Warning: No host groups found matching: ...
```
- Verify host group names in config match exactly (case-sensitive)
- Check API token has permission to view the host groups

### All Hosts Show 100% Availability

- Verify ICMP monitoring is configured on hosts
- Check trigger names contain "Unavailable by ICMP ping"
- Confirm triggers are High severity (level 4)

## Files

```
.
├── config.yaml              # Configuration file
├── zabbix_sla_report.py     # Main script
├── requirements.txt         # Python dependencies
├── venv/                    # Virtual environment
└── README.md                # This file
```

## Monthly link usage

Both the CLI and Streamlit reports include usage for interfaces whose configured
Zabbix description matches a case-insensitive regular expression in `config.yaml`:

```yaml
link_usage:
  enabled: true
  description_regex: 'PRI|SEC'
```

Replace the regex with your own pattern, for example `'WAN|UPLINK'`. Only the
interface description is matched, not the host or interface name. Standard Zabbix
network discovery names (`Interface name(description): Bits received`) and
`net.if.alias` / `net.if.descr` items are supported, paired by interface index.
Descriptions reflect the current discovery/configuration, not historical labels.

Usage always covers the previous calendar month, regardless of the SLA period.
The dashboard and additional Excel sheets contain download/upload GB and TB,
average/peak Mbps, average/peak utilization percentages, current interface capacity,
and hourly trend coverage for each direction. Download means traffic received by
the monitored interface; upload means traffic sent by it.

Traffic volume is an **estimate** obtained by integrating hourly average rates over
hours with retained trends. Missing hours are not extrapolated; averages cover only
available hours. Coverage counts hours with at least one sample and does not measure
sampling gaps within those hours. Peak is the maximum retained trend sample.
Utilization uses monthly average historical interface speed where retained,
otherwise current speed; the capacity basis is shown. Interface speed may differ
from a contracted service limit. Missing data or capacity produces blank metrics, not zero.
Standard `net.if.in` / `net.if.out` rate items must use `bps`, `b/s`, `Bps`, or `B/s`;
raw cumulative counters are not supported. Retain numeric trends for the reporting
month and grant API access to `item.get` and `trend.get`.

### Local API credential

The committed configuration contains empty connection fields and example group
and interface settings. Set your URL, host groups, and description regex locally.
For credentials, set `zabbix.token_file: ".zabbix-api-token"` and place your API
token in that file. It is excluded from Git and is resolved relative to the config
file. Alternatively, enter a token in the Streamlit sidebar. You can also remove
`token_file` and use the `zabbix.token` configuration field instead; keep real
credentials out of commits.
The connection test now checks authenticated host-group access as well as the
public API version endpoint.


## Gauges and historical charts

The dashboard includes SLA gauges for each group, with the selected SLA period's
availability and target. Each gauge uses that group's warning threshold.
Select an interface to view daily or monthly traffic volumes and peak rates,
and monthly capacity and demand growth. The reporting window defaults to the
last three **complete** calendar months:

```yaml
link_usage:
  enabled: true
  description_regex: 'link description'
  history_months: 3
  growth_min_coverage_pct: 90
```

`history_months` accepts 1–12 months. No prior saved reports are required: the
report reads retained Zabbix trends. Generate a new report after changing these
settings. The latest month's table remains separate from historical comparisons.

Demand growth compares average and peak Mbps to the immediately preceding month;
average Mbps avoids differences caused only by unequal month lengths. Capacity
growth compares monthly average **historical interface speed**, where available.
Growth percentages are blank when either month's hourly coverage is below
`growth_min_coverage_pct`, or the prior value is zero/missing. A missing intervening
month is never skipped. Current speed is not substituted into historical capacity
charts or capacity growth calculations. Three months show short-term change, not
a reliable long-term forecast or upgrade recommendation.

Excel exports include latest-month usage, monthly history with traffic/peak/capacity
charts per interface, and daily measurements. Saved report detail retains the
historical measurements. SLA gauges and interactive selectors are dashboard views.

### Retaining trends in the app

Set `link_usage.history_months` to **3**, **6**, or **12** for the last quarter,
six months, or year (complete calendar months). The active window is displayed
on the Generate Report page. New reports query that window from Zabbix; existing
saved snapshots keep the original window they were generated with.

`auto_save_reports: true` automatically saves generated Excel files and their
trend measurements in the local SQLite report history. Set it to `false` to use
manual saving. Under **Report History → View Details & Trends**, saved trend
charts remain available after restarting the app without querying Zabbix again.
This setting does not change Zabbix retention or schedule automatic collection.

To create separate monthly reports plus a comparison workbook:

```bash
python3 generate_monthly_reports.py --months 3 --output-dir reports
```

The command uses configured host groups and exclusions, queries traffic history
once per host, and writes one workbook per group per complete calendar month.
A further comparison workbook includes monthly SLA and per-interface trend charts.

**View Details & Trends** includes monthly SLA charts for the overall group or an
individual host, alongside traffic and peak charts. New app reports collect SLA
history over the configured window. Previously generated monthly exports retain
both host and group SLA history when imported.

Saved-report trend views combine retained measurements for the last configured
number of complete months, while snapshot tables and gauges show the selected
report's own period. Duplicate months use the newest retained measurements.
Monthly traffic uses grouped download/upload bars with calendar-month labels;
daily traffic uses a time-series line. Missing history requires generating a
report for a longer window; changing the config alone does not collect data.

## Persistent login

Logins survive browser refreshes and app restarts using a browser cookie containing
an opaque session token. The database stores only its hash, expiry, and user ID.
The cookie never contains your password or Zabbix API token. Logout invalidates
the server session and removes the cookie; password/role changes revoke existing
sessions. User deletion removes its sessions as well.

```yaml
auth:
  remember_days: 7
```

`remember_days` accepts 1–90 days and defaults to seven. Login once after loading
this update. Cookies are local to the browser and hostname you use; private
browsing or clearing cookies requires signing in again. The local cookie component
uses SameSite=Lax and Secure when served over HTTPS. Because it is written through
JavaScript, this cookie is not HttpOnly. Existing report data is preserved.
