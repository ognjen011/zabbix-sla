"""Combine retained monthly snapshots for the configured rolling trend window."""
from copy import deepcopy
from datetime import datetime

from link_usage import calendar_months


def combine_retained_trends(selected_report, retained_reports, months, reference_date=None):
    reference_date = reference_date or datetime.now()
    end = int(reference_date.replace(day=1, hour=0, minute=0, second=0, microsecond=0).timestamp()) - 1
    wanted = {datetime.fromtimestamp(start).strftime('%Y-%m') for start, _ in calendar_months(end, months)}
    result = deepcopy(selected_report)
    host_rows, group_rows = {}, {}
    identities = {}
    for report in retained_reports:
        for group, hosts in report.get('detail_data', {}).items():
            for host in hosts:
                host_key = (group, host.get('host') or host.get('name'))
                for row in host.get('link_history', []) + host.get('link_usage', []):
                    if row.get('Interface ID') and row.get('Interface'):
                        identities.setdefault((*host_key, row['Interface']), row['Interface ID'])
    # Reports arrive newest first: retain the latest measurement for duplicate periods.
    for report in retained_reports:
        for summary in report.get('summary_data', []):
            group = summary.get('group_name')
            for row in summary.get('sla_history', []):
                if row.get('Month') in wanted:
                    group_rows.setdefault(group, {}).setdefault(row['Month'], row)
        for group, hosts in report.get('detail_data', {}).items():
            for host in hosts:
                key = (group, host.get('host') or host.get('name'))
                fields = host_rows.setdefault(key, {'sla_history':{}, 'link_history':{}, 'link_daily':{}})
                for field in fields:
                    source = host.get(field, [])
                    if field == 'link_history' and not source:
                        source = host.get('link_usage', [])
                    for row in source:
                        month = (row.get('Month') or row.get('Date') or '')[:7]
                        if month not in wanted:
                            continue
                        period = row.get('Month') or row.get('Date')
                        identity = row.get('Interface ID') or identities.get((*key, row.get('Interface'))) or row.get('Interface')
                        if field != 'sla_history' and identity is not None:
                            row = {**row, 'Interface ID':identity}
                        row_key = period if field == 'sla_history' else (identity, period)
                        fields[field].setdefault(row_key, row)
    for summary in result.get('summary_data', []):
        summary['sla_history'] = sorted(group_rows.get(summary['group_name'], {}).values(), key=lambda row: row['Month'])
    for group, hosts in result.get('detail_data', {}).items():
        for host in hosts:
            key = (group, host.get('host') or host.get('name'))
            for field, rows in host_rows.get(key, {}).items():
                host[field] = sorted(rows.values(), key=lambda row: (row.get('Interface ID') or row.get('Interface') or '', row.get('Month') or row.get('Date')))
    result['trend_months'] = months
    return result
