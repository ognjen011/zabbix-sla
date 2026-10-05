import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from streamlit.testing.v1 import AppTest

import database
from tests.test_trend_reports import HistoryAPI


class FrontendAPI(HistoryAPI):
    def __init__(self, url, token):
        super().__init__()

    def _call(self, method, params=None, **kwargs):
        if method == 'apiinfo.version':
            return '7.4.14'
        if method == 'hostgroup.get':
            return [{'groupid':'1', 'name':'Example Group'}]
        return super()._call(method, params)

    def get_host_groups(self, names=None):
        return [{'groupid':'1', 'name':'Example Group'}]

    def get_hosts_in_group(self, group_id):
        return [{'hostid':'1', 'name':'router', 'host':'router'}]

    def get_host_availability(self, host_id, start, end):
        from sla_policy import calculate_availability
        return calculate_availability(start, end + 1, [], getattr(self, "sla_policy", {}))


class AppTrendTests(unittest.TestCase):
    def test_legacy_snapshot_without_interface_id(self):
        app = AppTest.from_string("""
from report_visuals import show_link_trends
show_link_trends({'Example Group':[{'name':'router', 'link_usage':[{'Interface':'eth0','Description':'PRI'}]}]}, key_prefix='history_1_')
""").run(timeout=20)
        self.assertFalse(app.exception)
        self.assertTrue(any('no historical trends' in item.value for item in app.info))

    def test_generation_auto_retention_and_history_chart_interaction(self):
        config = {'zabbix':{'url':'http://localhost/zabbix', 'token':'test'},
                  'host_groups':{'Example Group':{'sla_threshold':99.99, 'orange_threshold':5}},
                  'link_usage':{'history_months':3, 'description_regex':'PRI|SEC'},
                  'report_mode':'separate', 'auto_save_reports':True,
                  'sla_calculation':{'minimum_outage_seconds':30}}
        with tempfile.TemporaryDirectory() as directory, patch.object(database, 'DB_PATH', Path(directory) / 'test.db'), patch('zabbix_sla_report.ZabbixAPI', FrontendAPI), patch('zabbix_sla_report.load_config', return_value=config):
            app = AppTest.from_file('app.py')
            app.session_state['authenticated'] = True
            app.session_state['user'] = {'username':'admin', 'display_name':'Admin', 'role':'admin', 'id':1}
            app.run(timeout=30)
            self.assertFalse(app.exception)
            next(field for field in app.number_input if field.label == 'Ignore outages shorter than (seconds)').set_value(45)
            next(button for button in app.button if button.label == 'Generate Report').click().run(timeout=30)
            self.assertFalse(app.exception)
            self.assertEqual(database.get_report_count(), 1)
            stored = database.get_report(1)
            self.assertEqual(stored['detail_data']['Example Group'][0]['sla_policy']['minimum_outage_seconds'], 45)
            self.assertEqual(len(stored['detail_data']['Example Group'][0]['link_history']), 3)
            self.assertEqual(len(stored['detail_data']['Example Group'][0]['sla_history']), 3)
            self.assertEqual(len(stored['summary_data'][0]['sla_history']), 3)
            self.assertTrue(stored['detail_data']['Example Group'][0]['link_daily'])
            next(radio for radio in app.radio if radio.label == 'Navigation').set_value('Report History').run(timeout=30)
            next(button for button in app.button if button.label == 'View Details & Trends').click().run(timeout=30)
            self.assertFalse(app.exception)
            self.assertTrue(any(select.label == 'Interface' for select in app.selectbox))
            self.assertTrue(any(select.label == 'SLA scope' for select in app.selectbox))
            next(radio for radio in app.radio if radio.label == 'Traffic interval').set_value('Monthly').run(timeout=30)
            self.assertFalse(app.exception)
            self.assertEqual(database.get_report_count(), 1)


    def test_company_hours_with_no_eligible_day_render_na(self):
        from datetime import datetime
        from zabbix_sla_report import DateRangeCalculator
        periods = DateRangeCalculator.get_availability_periods(datetime(2026, 10, 5))
        config = {'zabbix':{'url':'http://localhost/zabbix','token':'test'},
            'host_groups':{'Example Group':{'sla_threshold':99.99,'orange_threshold':5,
                'sla_calculation':{'business_hours':{'enabled':True,'timezone':'UTC','weekdays':[0,1,2,3,4],'start':'08:00','end':'17:00'}}}},
            'link_usage':{'history_months':3,'description_regex':'PRI|SEC'}, 'report_mode':'separate'}
        with tempfile.TemporaryDirectory() as directory, patch.object(database, 'DB_PATH', Path(directory) / 'test.db'), patch('zabbix_sla_report.ZabbixAPI', FrontendAPI), patch('zabbix_sla_report.load_config', return_value=config), patch.object(DateRangeCalculator, 'get_availability_periods', return_value=periods):
            app = AppTest.from_file('app.py')
            app.session_state['authenticated'] = True
            app.session_state['user'] = {'username':'admin','display_name':'Admin','role':'admin','id':1}
            app.run(timeout=30)
            next(select for select in app.selectbox if select.label == 'SLA Period').set_value('day').run(timeout=30)
            next(button for button in app.button if button.label == 'Generate Report').click().run(timeout=30)
            self.assertFalse(app.exception)
            report = database.get_report(1)
            self.assertIsNone(report['detail_data']['Example Group'][0]['device_sla'])
            self.assertIsNone(report['summary_data'][0]['overall_sla'])


if __name__ == '__main__':
    unittest.main()
