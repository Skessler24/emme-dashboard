const test = require('node:test');
const assert = require('node:assert');
require('./fakes');
process.env.EMME_ADMIN_EMAILS = 'skessler@emmeco.org';

const schedule = require('../schedule/index.js');
const microsoft = require('../microsoft/index.js');
const me = require('../me/index.js');
const alerts = require('../admin-alerts/index.js');
const messages = require('../staff-messages/index.js');

function asUser(email) {
  const principal = Buffer.from(JSON.stringify({ userId: 'u-' + email, userDetails: email, claims: [{ typ: 'preferred_username', val: email }] })).toString('base64');
  return { 'x-ms-client-principal': principal };
}
async function call(handler, email, method, rest, body) {
  const context = {};
  await handler(context, { method, body, headers: asUser(email), params: { rest: rest || '' } });
  return context.res;
}

test('someone signed in from another organization is turned away everywhere', async () => {
  const outsider = 'teacher@cherrycreekschools.org';
  for (const [name, handler, method, rest] of [
    ['me', me, 'GET', ''],
    ['schedule', schedule, 'GET', ''],
    ['schedule save', schedule, 'PUT', ''],
    ['admin alerts', alerts, 'GET', ''],
    ['staff messages', messages, 'GET', ''],
    ['microsoft', microsoft, 'GET', 'calendar']
  ]) {
    const res = await call(handler, outsider, method, rest, {});
    assert.strictEqual(res.status, 403, name + ' answered ' + res.status);
    assert.match(JSON.stringify(res.body), /Emme staff/, name);
  }
});

test('an outsider never gets a schedule stored under their address', async () => {
  const { blobs } = require('./fakes');
  await call(schedule, 'teacher@cherrycreekschools.org', 'GET', '');
  await call(schedule, 'teacher@cherrycreekschools.org', 'PUT', '', { workspace: { boards: [] } });
  assert.ok(![...blobs.keys()].some((key) => key.includes('cherrycreek')));
});

test('Emme staff still get in', async () => {
  const ok = await call(me, 'skessler@emmeco.org', 'GET', '');
  assert.strictEqual(ok.status, 200);
  assert.strictEqual(ok.body.isAdmin, true);
  const staff = await call(schedule, 'ndelgado@emmeco.org', 'GET', '');
  assert.strictEqual(staff.status, 200);
});
