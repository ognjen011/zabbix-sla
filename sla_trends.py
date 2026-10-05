"""Monthly availability history for group and host SLA trend views."""
from collections import defaultdict
from datetime import datetime

from link_usage import calendar_months


def collect_sla_history(api, host_id, end, months, target, warning, latest=None):
    periods = calendar_months(end, months)
    rows = []
    for index, (start, stop) in enumerate(periods):
        stats = latest if latest is not None and index == len(periods) - 1 else api.get_host_availability(host_id, start, stop - 1)
        availability = stats['availability']
        rows.append({'Month':datetime.fromtimestamp(start).strftime('%Y-%m'),
            'Availability (%)':availability, 'SLA Target (%)':target,
            'Status':'N/A' if availability is None else 'COMPLIANT' if availability >= target else 'WARNING' if availability >= target - warning else 'BREACH',
            'Downtime (seconds)':stats['downtime_seconds'], 'Total (seconds)':stats['total_seconds']})
    return rows


def group_sla_history(hosts, target):
    months = defaultdict(list)
    for host in hosts:
        for row in host.get('sla_history', []):
            months[row['Month']].append(row)
    rows = []
    for month, data in sorted(months.items()):
        total = sum(row['Total (seconds)'] for row in data)
        downtime = sum(row['Downtime (seconds)'] for row in data)
        rows.append({'Month':month, 'Availability (%)':(1 - downtime / total) * 100 if total > 0 else None,
            'SLA Target (%)':target, 'Hosts':len(data),
            **{status.title():sum(row['Status'] == status for row in data) for status in ('COMPLIANT','WARNING','BREACH')}})
    return rows
