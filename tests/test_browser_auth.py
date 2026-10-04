import hashlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import database as db
from streamlit.testing.v1 import AppTest


class BrowserAuthTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.database_patch = patch.object(db, 'DB_PATH', Path(self.directory.name) / 'test.db')
        self.database_patch.start()
        db.init_db()
        self.user = db.authenticate('admin', 'admin')

    def tearDown(self):
        self.database_patch.stop()
        self.directory.cleanup()

    def test_session_restores_from_storage_without_password_or_raw_token(self):
        token = db.create_browser_session(self.user['id'])
        db.init_db()  # Simulate startup without losing retained sessions.
        self.assertEqual(db.get_browser_session(token)['username'], 'admin')
        with db.get_db() as connection:
            row = dict(connection.execute('SELECT * FROM browser_sessions').fetchone())
        self.assertEqual(row['token_hash'], hashlib.sha256(token.encode()).hexdigest())
        self.assertNotIn(token, row.values())
        self.assertEqual(set(db.get_browser_session(token)), {'id','username','role','display_name'})

    def test_forged_expired_and_logged_out_sessions_are_rejected(self):
        with patch('database.time.time', return_value=1000):
            token = db.create_browser_session(self.user['id'], days=1)
            self.assertIsNone(db.get_browser_session('x' * 43))
        with patch('database.time.time', return_value=1000 + 86400):
            self.assertIsNone(db.get_browser_session(token))
        token = db.create_browser_session(self.user['id'])
        db.revoke_browser_session(token)
        self.assertIsNone(db.get_browser_session(token))

    def test_password_role_and_account_deletion_revoke_sessions(self):
        token = db.create_browser_session(self.user['id'])
        self.assertTrue(db.change_password(self.user['id'], 'admin', 'new-password'))
        self.assertIsNone(db.get_browser_session(token))
        token = db.create_browser_session(self.user['id'])
        db.update_user(self.user['id'], role='user')
        self.assertIsNone(db.get_browser_session(token))
        token = db.create_browser_session(self.user['id'])
        db.delete_user(self.user['id'])
        self.assertIsNone(db.get_browser_session(token))

    def test_refresh_restores_login_and_logout_rejects_cached_cookie(self):
        token = db.create_browser_session(self.user['id'])
        script = '''
import streamlit as st
from browser_auth import restore_login, clear_login
if restore_login({}):
    st.write(st.session_state.get('user', {}).get('username', 'logged out'))
    st.button('Logout', on_click=clear_login)
'''
        with patch('browser_auth._cookie', return_value={'ready':True, 'token':token, 'operation_id':'read'}):
            # A fresh Streamlit websocket has empty session_state, like browser refresh.
            app = AppTest.from_string(script).run(timeout=20)
            self.assertFalse(app.exception)
            self.assertTrue(app.session_state['authenticated'])
            self.assertEqual(app.session_state['user']['username'], 'admin')
            app.button[0].click().run(timeout=20)
            self.assertFalse(app.exception)
            self.assertIsNone(db.get_browser_session(token))
            self.assertNotIn('authenticated', app.session_state)


if __name__ == '__main__':
    unittest.main()
