const test = require('node:test');
const assert = require('node:assert');

const { drive } = require('./fakes');
const handler = require('../schedule/index.js');
const workbook = require('../dist/scheduleWorkbook.js');
const principal = Buffer.from(JSON.stringify({ userId: 'n1', userDetails: 'ndelgado@emmeco.org', claims: [] })).toString('base64');

function day(id, name, slots) { return { id, name, color: 'blue', slots }; }
function slot(id, time, notes, status) { return { id, time, status, notes, lastDay: '' }; }

test('Download gives an Excel file of the schedule on screen and leaves SharePoint alone', async () => {
  const saved = { version: 1, boards: [{ id: 'b1', name: 'Current Schedule', days: [day('dMon', 'Monday', [slot('m1', '9:00 AM', 'Old Name', 'Full')])] }], clients: [], dxCodes: [] };
  drive.bytes = await workbook.buildScheduleWorkbook(saved);
  drive.version = 1;
  const before = Buffer.from(drive.bytes);
  const shown = { version: 1, boards: [{ id: 'b1', name: 'Current Schedule', days: [day('dMon', 'Monday', [slot('m1', '9:00 AM', 'Made Up Client', 'Full'), slot('m2', '10:00 AM', 'Other Made Up', 'Full')])] }], clients: [], dxCodes: [] };
  const context = {};
  await handler(context, { method: 'POST', body: { workspace: shown }, headers: { 'x-ms-client-principal': principal }, params: { rest: 'download' } });
  assert.strictEqual(context.res.status, 200);
  assert.match(context.res.headers['Content-Type'], /spreadsheetml/);
  assert.match(context.res.headers['Content-Disposition'], /Natalie Delgado Schedule\.xlsx/);
  const days = await workbook.readScheduleDays(new Uint8Array(context.res.body));
  const notes = (days.monday || []).map((row) => row.notes);
  assert.deepStrictEqual(notes, ['Made Up Client', 'Other Made Up']);
  assert.ok(Buffer.from(drive.bytes).equals(before), 'the SharePoint file was not changed');
});
