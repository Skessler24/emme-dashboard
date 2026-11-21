const test = require('node:test');
const assert = require('node:assert');

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

test('an admin brings in the old schedule from the provider\'s SharePoint folder, and the file is not changed', async () => {
  drive.bytes = await workbook.buildScheduleWorkbook(schedule(['Old Folder One', 'Old Folder Two']));
  drive.version = 1;
  const before = Buffer.from(drive.bytes);
  const putsBefore = drive.putCount;
  const shown = { version: 1, boards: [], clients: [{ id: 'c1', name: 'Kept Client' }], dxCodes: [] };

  const res = await call('boss@emmeco.org', 'POST', 'import-folder', { workspace: shown }, { provider: 'ndelgado@emmeco.org' });
  assert.strictEqual(res.status, 200, JSON.stringify(res.body));
  assert.deepStrictEqual(mondayNotes(res.body.workspace), ['Old Folder One', 'Old Folder Two']);
  assert.ok(/Schedule\.xlsx$/.test(res.body.fileName), res.body.fileName);
  assert.deepStrictEqual(res.body.workspace.clients.map((c) => c.name), ['Kept Client']);
  assert.ok(Buffer.from(drive.bytes).equals(before));
  assert.strictEqual(drive.putCount, putsBefore, 'nothing was written to SharePoint');
});

test('a folder workbook with no weekday tables gets a plain message naming the file', async () => {
  const saved = drive.bytes;
  const ExcelJS = require('exceljs');
  const book = new ExcelJS.Workbook();
  book.addWorksheet('Notes').getCell('A1').value = 'nothing useful';
  drive.bytes = Buffer.from(await book.xlsx.writeBuffer());
  const res = await call('boss@emmeco.org', 'POST', 'import-folder', { workspace: {} }, { provider: 'ndelgado@emmeco.org' });
  drive.bytes = saved;
  assert.strictEqual(res.status, 400, JSON.stringify(res.body));
  assert.ok(/Natalie Delgado Schedule\.xlsx: No weekday schedule tables/.test(res.body.error), res.body.error);
});

test('only an admin can import another provider\'s folder schedule', async () => {
  const res = await call('someone@emmeco.org', 'POST', 'import-folder', { workspace: {} }, { provider: 'ndelgado@emmeco.org' });
  assert.strictEqual(res.status, 403);
});
