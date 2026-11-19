const test = require('node:test');
const assert = require('node:assert');

process.env.EMME_ADMIN_EMAILS = 'boss@emmeco.org';
require('./fakes');
const handler = require('../schedule/index.js');
const workbook = require('../dist/scheduleWorkbook.js');

function who(email) {
  return Buffer.from(JSON.stringify({ userId: email, userDetails: email, claims: [] })).toString('base64');
}
async function call(email, method, rest, body, query) {
  const context = {};
  await handler(context, { method, body, query: query || {}, headers: { 'x-ms-client-principal': who(email) }, params: { rest: rest || '' } });
  return context.res;
}
function slot(id, time, notes, status) { return { id, time, status: status || 'Full', notes, lastDay: '' }; }
function schedule(notes) {
  return { version: 1, boards: [{ id: 'b1', name: 'Current Schedule', days: [{ id: 'dMon', name: 'Monday', color: 'blue', slots: notes.map((n, i) => slot('s' + i, (9 + i) + ':00 AM', n)) }] }], clients: [{ id: 'c1', name: 'Kept Client' }], dxCodes: [] };
}
function mondayNotes(workspace) {
  const board = workspace.boards.find((b) => b.name === 'Current Schedule');
  return board.days.find((d) => d.name === 'Monday').slots.filter((s) => s.notes || s.status).map((s) => s.notes);
}

test('an admin imports a schedule workbook onto the provider\'s schedule and saves it', async () => {
  const file = await workbook.buildScheduleWorkbook(schedule(['Uploaded One', 'Uploaded Two']));
  const opened = await call('boss@emmeco.org', 'GET', '', null, { provider: 'ndelgado@emmeco.org' });
  const shown = opened.body.workspace;

  const res = await call('boss@emmeco.org', 'POST', 'import', { data: file.toString('base64'), workspace: shown }, { provider: 'ndelgado@emmeco.org' });
  assert.strictEqual(res.status, 200, JSON.stringify(res.body));
  assert.deepStrictEqual(mondayNotes(res.body.workspace), ['Uploaded One', 'Uploaded Two']);
  assert.ok(res.body.days.includes('monday'));
  assert.strictEqual(res.body.rows, 2);

  const put = await call('boss@emmeco.org', 'PUT', '', { workspace: res.body.workspace, base: shown }, { provider: 'ndelgado@emmeco.org' });
  assert.strictEqual(put.status, 200);
  const providerView = await call('ndelgado@emmeco.org', 'GET', '');
  assert.deepStrictEqual(mondayNotes(providerView.body.workspace), ['Uploaded One', 'Uploaded Two']);
});

test('importing a schedule workbook keeps the clients and replaces only the filled rows of the days in the file', async () => {
  const file = await workbook.buildScheduleWorkbook(schedule(['From The File']));
  const current = schedule(['Old One', 'Old Two', 'Old Three']);
  const res = await call('ndelgado@emmeco.org', 'POST', 'import', { data: file.toString('base64'), workspace: current });
  assert.strictEqual(res.status, 200, JSON.stringify(res.body));
  assert.deepStrictEqual(mondayNotes(res.body.workspace), ['From The File']);
  assert.deepStrictEqual(res.body.workspace.clients.map((c) => c.name), ['Kept Client']);
});

test('a provider with no schedule yet gets one from the workbook', async () => {
  const file = await workbook.buildScheduleWorkbook(schedule(['Brand New']));
  const res = await call('someone.new@emmeco.org', 'POST', 'import', { data: file.toString('base64'), workspace: { version: 1, boards: [], clients: [], dxCodes: [] } });
  assert.strictEqual(res.status, 200, JSON.stringify(res.body));
  assert.deepStrictEqual(mondayNotes(res.body.workspace), ['Brand New']);
});

test('a file that is not a schedule workbook, or no file, is refused with a plain message', async () => {
  const missing = await call('ndelgado@emmeco.org', 'POST', 'import', { workspace: {} });
  assert.strictEqual(missing.status, 400);
  const junk = await call('ndelgado@emmeco.org', 'POST', 'import', { data: Buffer.from('not a workbook').toString('base64'), workspace: {} });
  assert.strictEqual(junk.status, 400);
  assert.ok(junk.body.error);
});

test('only an admin can import onto another provider\'s schedule', async () => {
  const file = await workbook.buildScheduleWorkbook(schedule(['Nope']));
  const res = await call('someone@emmeco.org', 'POST', 'import', { data: file.toString('base64'), workspace: {} }, { provider: 'ndelgado@emmeco.org' });
  assert.strictEqual(res.status, 403);
});
