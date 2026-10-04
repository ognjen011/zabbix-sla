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
        return {'availability':100, 'downtime_seconds':0, 'total_seconds':end-start}


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
                  'report_mode':'separate', 'auto_save_reports':True}
        with tempfile.TemporaryDirectory() as directory, patch.object(database, 'DB_PATH', Path(directory) / 'test.db'), patch('zabbix_sla_report.ZabbixAPI', FrontendAPI), patch('zabbix_sla_report.load_config', return_value=config):
            app = AppTest.from_file('app.py')
            app.session_state['authenticated'] = True
            app.session_state['user'] = {'username':'admin', 'display_name':'Admin', 'role':'admin', 'id':1}
            app.run(timeout=30)
            self.assertFalse(app.exception)
            next(button for button in app.button if button.label == 'Generate Report').click().run(timeout=30)
            self.assertFalse(app.exception)
            self.assertEqual(database.get_report_count(), 1)
            stored = database.get_report(1)
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


if __name__ == '__main__':
    unittest.main()
