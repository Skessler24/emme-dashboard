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
const sample = { version: 1, boards: [{ id: 'b1', name: 'Current Schedule', days: [{ id: 'dMon', name: 'Monday', color: 'blue', slots: [{ id: 's1', time: '9:00 AM', status: 'Full', notes: 'Made Up Folder Client', lastDay: '' }] }] }], clients: [], dxCodes: [] };

test('an admin gets the old workbook from the provider\'s folder, untouched, for the page to read', async () => {
  drive.bytes = await workbook.buildScheduleWorkbook(sample);
  drive.version = 1;
  const before = Buffer.from(drive.bytes);
  const putsBefore = drive.putCount;

  const res = await call('boss@emmeco.org', 'POST', 'folder-file', {}, { provider: 'ndelgado@emmeco.org' });
  assert.strictEqual(res.status, 200, JSON.stringify(res.body));
  assert.ok(/Schedule\.xlsx$/.test(res.body.fileName), res.body.fileName);
  const bytes = Buffer.from(res.body.data, 'base64');
  assert.ok(bytes.equals(before), 'the file arrives exactly as it is in the folder');
  const days = await workbook.readScheduleDays(new Uint8Array(bytes));
  assert.strictEqual(days.monday[0].notes, 'Made Up Folder Client');
  assert.ok(Buffer.from(drive.bytes).equals(before));
  assert.strictEqual(drive.putCount, putsBefore, 'nothing was written to SharePoint');
});

test('only an admin can fetch another provider\'s folder file', async () => {
  const res = await call('someone@emmeco.org', 'POST', 'folder-file', {}, { provider: 'ndelgado@emmeco.org' });
  assert.strictEqual(res.status, 403);
});
