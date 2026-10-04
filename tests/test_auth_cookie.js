const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const code = fs.readFileSync('auth_cookie_component/cookie.js', 'utf8');
const jar = new Map([['other', 'keep']]);
let cookieWrite = '';
const document = {referrer: 'https://app.example/'};
Object.defineProperty(document, 'cookie', {
  get: () => [...jar].map(([key, value]) => `${key}=${value}`).join('; '),
  set: value => {
    cookieWrite = value;
    const [pair, ...options] = value.split('; ');
    const separator = pair.indexOf('=');
    const name = pair.slice(0, separator);
    if (options.includes('Max-Age=0')) jar.delete(name);
    else jar.set(name, pair.slice(separator + 1));
  },
});
function browser() {
  const messages = [];
  let listener;
  const parent = {postMessage: (message, target) => {
    assert.equal(target, 'https://app.example');
    messages.push(message);
  }};
  const window = {parent, location: {href: 'https://app.example/component/auth/index.html', protocol: 'https:'}, addEventListener: (_, handler) => listener = handler};
  vm.runInNewContext(code, {document, window, URL});
  return {messages, render: args => listener({source: parent, origin: 'https://app.example', data: {type: 'streamlit:render', args}})};
}
let client = browser();
const read = {action: 'read', cookie_name: 'zabbix_sla_session', operation_id: 'read'};
client.render(read);
assert.equal(client.messages.find(m => m.type === 'streamlit:setComponentValue').value.token, null);
client.render({...read, action: 'set', token: 'opaque-token', max_age: 604800, operation_id: 'login'});
assert.match(cookieWrite, /SameSite=Lax; Secure/);
assert.equal(jar.get('zabbix_sla_session'), 'opaque-token');
client = browser(); // Browser reload: cookie survives, JavaScript state does not.
client.render(read);
assert.equal(client.messages.find(m => m.type === 'streamlit:setComponentValue').value.token, 'opaque-token');
const count = client.messages.filter(m => m.type === 'streamlit:setComponentValue').length;
client.render(read);
assert.equal(client.messages.filter(m => m.type === 'streamlit:setComponentValue').length, count);
client.render({...read, action: 'clear', operation_id: 'logout'});
assert.equal(jar.has('zabbix_sla_session'), false);
assert.equal(jar.get('other'), 'keep');
console.log('Cookie set, refresh, duplicate-event handling, and logout pass');
