import io
import unittest
from datetime import datetime, timezone
from zoneinfo import ZoneInfo

from sla_policy import calculate_availability, resolve_policy


def stamp(value, zone='UTC'):
    return int(datetime.fromisoformat(value).replace(tzinfo=ZoneInfo(zone)).timestamp())


class PolicyTests(unittest.TestCase):
    def test_unset_policy_is_24h_and_overlaps_are_counted_once(self):
        row = calculate_availability(0, 86400, [(100, 200), (150, 250), (100, 200)])
        self.assertEqual(row['total_seconds'], 86400)
        self.assertEqual(row['downtime_seconds'], 150)

    def test_minimum_is_optional_and_exact_minute_is_counted(self):
        intervals = [(0, 59), (100, 160)]
        self.assertEqual(calculate_availability(0, 1000, intervals)['downtime_seconds'], 119)
        self.assertEqual(calculate_availability(0, 1000, intervals, {'minimum_outage_seconds':60})['downtime_seconds'], 60)
        # A long outage is not discarded just because only 30 seconds overlap the report.
        self.assertEqual(calculate_availability(130, 300, [(100,160)], {'minimum_outage_seconds':60})['downtime_seconds'], 30)

    def test_business_time_changes_denominator_and_clips_outages(self):
        policy = {'business_hours':{'enabled':True,'timezone':'UTC','weekdays':[0,1,2,3,4],'start':'08:00','end':'17:00'}}
        start, end = stamp('2026-10-05T00:00'), stamp('2026-10-06T00:00')
        row = calculate_availability(start, end, [(stamp('2026-10-05T07:00'), stamp('2026-10-05T09:00')), (stamp('2026-10-05T18:00'), stamp('2026-10-05T19:00'))], policy)
        self.assertEqual(row['total_seconds'], 9*3600)
        self.assertEqual(row['downtime_seconds'], 3600)
        row = calculate_availability(stamp('2026-10-04T00:00'), stamp('2026-10-05T00:00'), [], policy)
        self.assertEqual(row['total_seconds'], 0)
        self.assertIsNone(row['availability'])

    def test_overnight_shift_spills_into_following_day(self):
        policy = {'business_hours':{'enabled':True,'timezone':'UTC','weekdays':[4],'start':'22:00','end':'06:00'}}
        row = calculate_availability(stamp('2026-10-03T00:00'), stamp('2026-10-04T00:00'), [], policy)
        self.assertEqual(row['total_seconds'], 6*3600)

    def test_timezone_dst_and_group_overrides(self):
        policy = resolve_policy({'minimum_outage_seconds':60}, {'business_hours':{'enabled':True,'timezone':'Europe/Belgrade','weekdays':[6],'start':'01:00','end':'04:00'}})
        row = calculate_availability(stamp('2026-10-25T00:00','Europe/Belgrade'), stamp('2026-10-26T00:00','Europe/Belgrade'), [], policy)
        self.assertEqual(row['total_seconds'], 4*3600)
        self.assertEqual(policy['minimum_outage_seconds'], 60)

    def test_no_eligible_hours_export_as_na(self):
        from zabbix_sla_report import ExcelReportGenerator
        report = ExcelReportGenerator(99.9,5)
        host = {'name':'example','host':'example','avail_1_day':None,'avail_7_days':None,'avail_prev_month':None,'device_sla':None,
                'total_1_day':0,'total_7_days':0,'total_prev_month':0}
        report.create_sheet('Example', [host], 99.9)
        report.add_summary_sheet([{'group_name':'Example','total':1,'compliant':0,'warning':0,'breach':0,'overall_sla':None,
            'overall_1_day':None,'overall_7_days':None,'overall_prev_month':None}])
        report.workbook.save(io.BytesIO())
        self.assertEqual(report.workbook['Example']['H2'].value,'N/A')


if __name__ == '__main__':
    unittest.main()
