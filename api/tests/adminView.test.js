const test = require('node:test');
const assert = require('node:assert');
const ExcelJS = require('exceljs');

process.env.EMME_ADMIN_EMAILS = 'boss@emmeco.org';
const { drive } = require('./fakes');
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
  return { version: 1, boards: [{ id: 'b1', name: 'Current Schedule', days: [{ id: 'dMon', name: 'Monday', color: 'blue', slots: notes.map((n, i) => slot('s' + i, (9 + i) + ':00 AM', n)) }] }], clients: [], dxCodes: [] };
}
function notesOf(res) {
  const board = res.body.workspace.boards.find((b) => b.name === 'Current Schedule');
  const monday = board && board.days.find((d) => d.name === 'Monday');
  return monday ? monday.slots.map((s) => s.notes) : [];
}

test('a provider saves, and an admin can list providers and open that same schedule', async () => {
  drive.bytes = await workbook.buildScheduleWorkbook(schedule(['Seed']));
  drive.version = 1;
  const saved = await call('ndelgado@emmeco.org', 'PUT', '', { workspace: schedule(['Made Up One', 'Made Up Two']) });
  assert.strictEqual(saved.status, 200);

  const list = await call('boss@emmeco.org', 'GET', 'providers');
  assert.strictEqual(list.status, 200);
  assert.ok(list.body.providers.some((p) => p.email === 'ndelgado@emmeco.org'), JSON.stringify(list.body));

  const opened = await call('boss@emmeco.org', 'GET', '', null, { provider: 'ndelgado@emmeco.org' });
  assert.strictEqual(opened.status, 200);
  assert.deepStrictEqual(notesOf(opened), ['Made Up One', 'Made Up Two']);
});

test('an admin edit is saved to the provider\'s own schedule', async () => {
  const opened = await call('boss@emmeco.org', 'GET', '', null, { provider: 'ndelgado@emmeco.org' });
  const edited = opened.body.workspace;
  edited.boards.find((b) => b.name === 'Current Schedule').days[0].slots[0].notes = 'Changed By Admin';
  const put = await call('boss@emmeco.org', 'PUT', '', { workspace: edited, base: opened.body.workspace }, { provider: 'ndelgado@emmeco.org' });
  assert.strictEqual(put.status, 200);
  const providerView = await call('ndelgado@emmeco.org', 'GET', '');
  assert.strictEqual(notesOf(providerView)[0], 'Changed By Admin');
  const adminOwn = await call('boss@emmeco.org', 'GET', '');
  assert.ok(!notesOf(adminOwn).includes('Changed By Admin'), 'the admin\'s own schedule is untouched');
});

test('a provider who is not an admin cannot list providers or open anyone else', async () => {
  const list = await call('someone@emmeco.org', 'GET', 'providers');
  assert.strictEqual(list.status, 403);
  const other = await call('someone@emmeco.org', 'GET', '', null, { provider: 'ndelgado@emmeco.org' });
  assert.strictEqual(other.status, 403);
  const bad = await call('boss@emmeco.org', 'GET', '', null, { provider: 'x@example.com' });
  assert.strictEqual(bad.status, 400);
});

test('the SharePoint workbook is a copy: dashboard saves rewrite it, and edits made in it are not read back', async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(drive.bytes);
  const sheet = wb.getWorksheet('Current Schedule');
  let edited = false;
  for (let r = 1; r <= sheet.rowCount && !edited; r++) {
    if (String(sheet.getRow(r).getCell(3).text) === 'Made Up Two') { sheet.getRow(r).getCell(3).value = 'Typed In Excel'; edited = true; }
  }
  assert.ok(edited);
  drive.bytes = Buffer.from(await wb.xlsx.writeBuffer());
  drive.version += 1;

  const loaded = await call('ndelgado@emmeco.org', 'GET', '');
  assert.ok(!notesOf(loaded).includes('Typed In Excel'), 'workbook edits are not pulled into the dashboard');

  const current = loaded.body.workspace;
  const put = await call('ndelgado@emmeco.org', 'PUT', '', { workspace: current, base: current });
  assert.strictEqual(put.body.sharePoint.ok, true, JSON.stringify(put.body.sharePoint));
  const days = await workbook.readScheduleDays(drive.bytes);
  const notes = (days.monday || []).map((row) => row.notes);
  assert.ok(!notes.includes('Typed In Excel'), 'the workbook was rewritten from the dashboard');
  assert.ok(notes.includes('Made Up Two'));
});
