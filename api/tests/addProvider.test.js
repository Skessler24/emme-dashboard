const test = require('node:test');
const assert = require('node:assert');

process.env.EMME_ADMIN_EMAILS = 'boss@emmeco.org';
require('./fakes');
const handler = require('../schedule/index.js');

function who(email) {
  return Buffer.from(JSON.stringify({ userId: email, userDetails: email, claims: [] })).toString('base64');
}
async function call(email, method, rest, body, query) {
  const context = {};
  await handler(context, { method, body, query: query || {}, headers: { 'x-ms-client-principal': who(email) }, params: { rest: rest || '' } });
  return context.res;
}
function mondayTimes(workspace) {
  const board = workspace.boards.find((b) => b.name === 'Current Schedule');
  return board.days.find((d) => d.name === 'Monday').slots.map((s) => s.time);
}

test('an admin adds a provider who has no schedule yet and they get the standard one', async () => {
  const added = await call('boss@emmeco.org', 'POST', 'providers', { email: 'New.Hire@emmeco.org' });
  assert.strictEqual(added.status, 201, JSON.stringify(added.body));
  assert.strictEqual(added.body.provider.email, 'new.hire@emmeco.org');
  assert.strictEqual(added.body.created, true);

  const list = await call('boss@emmeco.org', 'GET', 'providers');
  assert.ok(list.body.providers.some((p) => p.email === 'new.hire@emmeco.org'), JSON.stringify(list.body));

  const opened = await call('boss@emmeco.org', 'GET', '', null, { provider: 'new.hire@emmeco.org' });
  assert.strictEqual(opened.status, 200);
  assert.strictEqual(mondayTimes(opened.body.workspace)[0], '7:00 AM');
  assert.strictEqual(opened.body.workspace.boards[0].name, 'Current Schedule');

  const theirs = await call('new.hire@emmeco.org', 'GET', '');
  assert.strictEqual(mondayTimes(theirs.body.workspace).length, 11);
});

test('adding a provider who already has a schedule never replaces it', async () => {
  const edited = (await call('boss@emmeco.org', 'GET', '', null, { provider: 'new.hire@emmeco.org' })).body.workspace;
  edited.boards[0].days[0].slots[0].notes = 'Real Client';
  await call('boss@emmeco.org', 'PUT', '', { workspace: edited, base: edited }, { provider: 'new.hire@emmeco.org' });
  const again = await call('boss@emmeco.org', 'POST', 'providers', { email: 'new.hire@emmeco.org' });
  assert.strictEqual(again.status, 200);
  assert.strictEqual(again.body.created, false);
  const kept = await call('new.hire@emmeco.org', 'GET', '');
  assert.strictEqual(kept.body.workspace.boards[0].days[0].slots[0].notes, 'Real Client');
});

test('only an admin can add a provider, and only a real Emme email', async () => {
  assert.strictEqual((await call('someone@emmeco.org', 'POST', 'providers', { email: 'x@emmeco.org' })).status, 403);
  assert.strictEqual((await call('boss@emmeco.org', 'POST', 'providers', { email: 'x@example.com' })).status, 400);
  assert.strictEqual((await call('boss@emmeco.org', 'POST', 'providers', {})).status, 400);
  const ghost = await call('boss@emmeco.org', 'POST', 'providers', { email: 'ghost@emmeco.org' });
  assert.strictEqual(ghost.status, 404);
  const list = await call('boss@emmeco.org', 'GET', 'providers');
  assert.ok(!list.body.providers.some((p) => p.email === 'ghost@emmeco.org'), 'nothing was created for an unknown account');
});
