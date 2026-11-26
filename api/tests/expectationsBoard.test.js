const test = require('node:test');
const assert = require('node:assert/strict');
process.env.EMME_ADMIN_EMAILS = 'skessler@emmeco.org';
process.env.AUTOMATION_API_KEY = 'test-automation-key';
require('./fakes');
const mileage = require('../mileage/index.js');

function principal(email) {
  return Buffer.from(JSON.stringify({ userId: 'u-' + email, userDetails: email, claims: [{ typ: 'preferred_username', val: email }] })).toString('base64');
}

async function call(email, method, body, query, headers) {
  const context = {};
  await mileage(context, {
    method,
    body,
    query: query || {},
    headers: Object.assign(email ? { 'x-ms-client-principal': principal(email) } : {}, headers || {}),
  });
  return context.res;
}

test('mileage expectations save the daily email check, keep travel approval with an administrator, and the morning sync has no note text', async () => {
  const provider = 'alex@emmeco.org';
  const saved = await call(provider, 'POST', { action: 'expectations-email', checked: true });
  assert.equal(saved.status, 200, JSON.stringify(saved.body));
  const today = await call(provider, 'GET', null, { part: 'email' });
  assert.equal(today.body.checked, true);
  const zone = await call(provider, 'POST', { action: 'expectations-zone', clientAddress: '100 Main St, Denver, CO 80202' });
  assert.equal(zone.status, 400);
  assert.match(zone.body.error, /home address/i);
  const denied = await call(provider, 'POST', { action: 'expectations-decision', id: 'nope', decision: 'approved' });
  assert.equal(denied.status, 403);
  const synced = await call(null, 'POST', { action: 'expectations-sync' }, {}, { 'x-emme-automation-key': 'test-automation-key' });
  assert.equal(synced.status, 200, JSON.stringify(synced.body));
  assert.equal(typeof synced.body.lateNotes, 'number');
  assert.doesNotMatch(JSON.stringify(synced.body), /NoteText|narrative/);
});
