import unittest
from datetime import datetime

from retained_trends import combine_retained_trends
from sla_trends import collect_sla_history, group_sla_history


class AvailabilityAPI:
    def __init__(self):
        self.calls = []

    def get_host_availability(self, host_id, start, end):
        self.calls.append((start, end))
        return {'availability':99.5, 'downtime_seconds':5, 'total_seconds':1000}


class SLATrendTests(unittest.TestCase):
    def test_latest_month_is_reused_and_history_has_three_months(self):
        api = AvailabilityAPI()
        latest = {'availability':99.8, 'downtime_seconds':2, 'total_seconds':1000}
        rows = collect_sla_history(api, '1', int(datetime(2026, 9, 30, 23, 59, 59).timestamp()), 3, 99.99, 5, latest)
        self.assertEqual(len(api.calls), 2)
        self.assertEqual([row['Month'] for row in rows], ['2026-07', '2026-08', '2026-09'])
        self.assertEqual(rows[-1]['Availability (%)'], 99.8)

    def test_group_sla_is_weighted_by_seconds(self):
        row = group_sla_history([
            {'sla_history':[{'Month':'2026-09','Total (seconds)':100,'Downtime (seconds)':10,'Status':'WARNING'}]},
            {'sla_history':[{'Month':'2026-09','Total (seconds)':900,'Downtime (seconds)':0,'Status':'COMPLIANT'}]},
        ], 99.99)[0]
        self.assertEqual(row['Availability (%)'], 99)
        self.assertEqual(row['Hosts'], 2)

    def test_separate_monthly_snapshot_uses_full_retained_window_and_deduplicates(self):
        def snapshot(month, value, identity='7'):
            link = {'Month':month,'Interface':'eth0','Description':'PRI','Download (GB)':value}
            if identity:
                link['Interface ID'] = identity
            return {'summary_data':[{'group_name':'Example Group','sla_history':[{'Month':month,'Availability (%)':value}]}],
                    'detail_data':{'Example Group':[{'name':'router','host':'technical','link_usage':[link],'link_history':[link]}]}}
        selected = snapshot('2026-07', 1)
        reports = [snapshot('2026-09', 3), snapshot('2026-08', 2), selected, snapshot('2026-09', 9, None)]
        combined = combine_retained_trends(selected, reports, 3, datetime(2026,10,4))
        rows = combined['detail_data']['Example Group'][0]['link_history']
        self.assertEqual([row['Month'] for row in rows], ['2026-07','2026-08','2026-09'])
        self.assertEqual(rows[-1]['Download (GB)'], 3)
        self.assertEqual(len(combined['summary_data'][0]['sla_history']), 3)
        self.assertEqual(len(selected['detail_data']['Example Group'][0]['link_history']), 1)


if __name__ == '__main__':
    unittest.main()
