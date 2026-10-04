"""Restore browser login using an opaque, revocable SQLite-backed session cookie."""
import secrets
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

import database as db

_cookie = components.declare_component('zabbix_session_cookie', path=str(Path(__file__).parent / 'auth_cookie_component'))
COOKIE_NAME = 'zabbix_sla_session'


def remember_days(config):
    days = (config.get('auth', {}) or {}).get('remember_days', 7)
    if type(days) is not int or not 1 <= days <= 90:
        raise ValueError('auth.remember_days must be an integer from 1 to 90')
    return days


def cookie_operation(action, token=None, days=7):
    st.session_state['auth_cookie_operation'] = {'action':action, 'token':token,
        'max_age':days * 86400, 'operation_id':secrets.token_hex(8)}


def set_login(user, config):
    token = db.create_browser_session(user['id'], remember_days(config))
    st.session_state['authenticated'] = True
    st.session_state['user'] = {key:user[key] for key in ('id','username','role','display_name')}
    st.session_state['browser_session_token'] = token
    cookie_operation('set', token, remember_days(config))


def clear_login():
    db.revoke_browser_session(st.session_state.get('browser_session_token'))
    for key in list(st.session_state):
        del st.session_state[key]
    cookie_operation('clear')


def restore_login(config):
    """Return False while the initial browser cookie read is pending."""
    remember_days(config)
    operation = st.session_state.get('auth_cookie_operation')
    args = operation or {'action':'read', 'token':None, 'max_age':0, 'operation_id':'read'}
    response = _cookie(cookie_name=COOKIE_NAME, **args, key='auth_cookie_bridge', default=None)
    if operation:
        if response and response.get('operation_id') == operation['operation_id']:
            st.session_state.pop('auth_cookie_operation', None)
            if response.get('error'):
                st.warning(response['error'])
        # Do not restore a stale cookie while setting or clearing it.
        token = st.session_state.get('browser_session_token')
        if token:
            user = db.get_browser_session(token)
            if user:
                st.session_state['user'] = user
            else:
                clear_login()
        return True
    token = st.session_state.get('browser_session_token') or (response or {}).get('token')
    if token:
        user = db.get_browser_session(token)
        if user:
            st.session_state['authenticated'] = True
            st.session_state['user'] = user
            st.session_state['browser_session_token'] = token
        else:
            clear_login()
            st.rerun()
    elif st.session_state.get('authenticated'):
        # Keep a login established before this feature was loaded.
        user = db.get_user(st.session_state['user']['id'])
        if user:
            set_login(user, config)
            st.rerun()
        else:
            clear_login()
            st.rerun()
    return bool(response and response.get('ready')) or st.session_state.get('authenticated', False)
