const test = require('node:test');
const assert = require('node:assert');
const ExcelJS = require('exceljs');
require('./fakes');
const workbook = require('../dist/scheduleWorkbook.js');

function day(id, name, slots) { return { id, name, color: 'blue', slots }; }
const shown = { version: 1, boards: [{ id: 'b1', name: 'Current Schedule', days: [day('dMon', 'Monday', [{ id: 'm1', time: '9:00 AM', status: 'Open 4x', notes: 'Made Up Client', lastDay: '' }])] }], clients: [], dxCodes: [] };

async function names(bytes) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(bytes));
  return wb.worksheets.map((sheet) => sheet.name);
}

test('the Current Schedule tab comes before the other tabs after a sync', async () => {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('Clients Not in Regular Space').getCell(1, 1).value = 'Made Up Name';
  const sheet = wb.addWorksheet('Current Schedule');
  sheet.getCell(1, 1).value = 'Mon Appointment Time ';
  sheet.getCell(1, 2).value = 'Status ';
  sheet.getCell(2, 1).value = 0.375;
  const before = Buffer.from(await wb.xlsx.writeBuffer());
  assert.deepStrictEqual(await names(before), ['Clients Not in Regular Space', 'Current Schedule']);
  const after = await workbook.applyScheduleWorkbook(new Uint8Array(before), shown);
  assert.deepStrictEqual(await names(after), ['Current Schedule', 'Clients Not in Regular Space']);
  const check = new ExcelJS.Workbook();
  await check.xlsx.load(after);
  assert.strictEqual(check.getWorksheet('Clients Not in Regular Space').getCell(1, 1).value, 'Made Up Name');
});

test('the workbook built from scratch has the status drop-down on every table', async () => {
  const built = await workbook.buildScheduleWorkbook(shown);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(built);
  const sheet = wb.getWorksheet('Current Schedule');
  assert.strictEqual(wb.worksheets[0].name, 'Current Schedule');
  const rule = sheet.getCell('B2').dataValidation;
  assert.strictEqual(rule && rule.type, 'list');
  const headers = [1, 2, 3, 4, 5].map((col) => String(sheet.getCell(1, col).value || ''));
  assert.ok(headers[4] === 'Schedule notes');
});

test('schedule notes are written on the far right of the read-only copy', async () => {
  const withNotes = { version: 1, boards: [{ id: 'b1', name: 'Current Schedule', days: [day('dWed', 'Wednesday', [{ id: 'w1', time: '4:30 PM', status: 'Open 2x (1&3)', notes: '', lastDay: '', scheduleNotes: 'Prefer a weekly client' }])] }], clients: [], dxCodes: [] };
  const built = await workbook.buildScheduleWorkbook(withNotes, { readOnly: true });
  const read = await workbook.readScheduleDays(new Uint8Array(built));
  assert.strictEqual(read.wednesday[0].scheduleNotes, 'Prefer a weekly client');
  assert.strictEqual(read.wednesday[0].status, 'Open 2x (1&3)');
});
