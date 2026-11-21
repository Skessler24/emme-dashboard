const test = require('node:test');
const assert = require('node:assert');

require('./fakes');
const handler = require('../schedule/index.js');

function who(email) {
  return Buffer.from(JSON.stringify({ userId: email, userDetails: email, claims: [] })).toString('base64');
}
async function call(email, method, body) {
  const context = {};
  await handler(context, { method, body, query: {}, headers: { 'x-ms-client-principal': who(email) }, params: { rest: '' } });
  return context.res;
}
const board = (id, name) => ({ id, name, days: [{ id: 'd' + id, name: 'Monday', slots: [{ id: 's' + id, time: '9:00 AM', status: 'Full', notes: name, lastDay: '' }] }] });
const names = (res) => res.body.workspace.boards.map((b) => b.name);

test('Current Schedule is always the first tab and Clients not in regular spot the second', async () => {
  const saved = await call('order@emmeco.org', 'PUT', {
    workspace: { version: 1, boards: [board('x', 'Fall Plan'), board('i', 'Clients not in regular spot'), board('c', 'Current Schedule'), board('s', "Summer '26")], clients: [], dxCodes: [] }
  });
  assert.strictEqual(saved.status, 200);
  assert.deepStrictEqual(names(saved), ['Current Schedule', 'Clients not in regular spot', 'Fall Plan']);
  const loaded = await call('order@emmeco.org', 'GET');
  assert.deepStrictEqual(names(loaded), ['Current Schedule', 'Clients not in regular spot', 'Fall Plan']);
});

test('without a Current Schedule board, the regular-spot board still comes before any other board', async () => {
  const saved = await call('order2@emmeco.org', 'PUT', { workspace: { version: 1, boards: [board('x', 'Fall Plan')], clients: [], dxCodes: [] } });
  assert.deepStrictEqual(names(saved), ['Clients not in regular spot', 'Fall Plan']);
});
