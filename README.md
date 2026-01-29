# Zabbix SLA Report Generator

Generate Excel SLA reports from Zabbix availability data based on ICMP ping monitoring.

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
