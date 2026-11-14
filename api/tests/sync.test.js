const test = require('node:test');
const assert = require('node:assert');
const ExcelJS = require('exceljs');

const { drive } = require('./fakes');

const handler = require('../schedule/index.js');
const workbook = require('../dist/scheduleWorkbook.js');
const principal = Buffer.from(JSON.stringify({ userId: 'n1', userDetails: 'ndelgado@emmeco.org', claims: [] })).toString('base64');

async function call(method, rest, body) {
  const context = {};
  await handler(context, { method, body, headers: { 'x-ms-client-principal': principal }, params: { rest: rest || '' } });
  return context.res;
}

function slot(id, time, notes, status) {
  return { id, time, status: status || '', notes, lastDay: '' };
}
function board(days) {
  return { id: 'b1', name: 'Current Schedule', days };
}
function day(id, name, slots) { return { id, name, color: 'blue', slots }; }

async function sheetDays() { return workbook.readScheduleDays(drive.bytes); }

async function editWorkbook(fn) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(drive.bytes);
  fn(wb.getWorksheet('Current Schedule'));
  drive.bytes = Buffer.from(await wb.xlsx.writeBuffer());
  drive.version += 1;
}

function findRow(sheet, text) {
  for (let r = 1; r <= sheet.rowCount; r++) if (String(sheet.getRow(r).getCell(3).text) === text) return r;
  return 0;
}

let ws;

test('setup: provider with a SharePoint schedule saves their first schedule', async () => {
  const seed = { version: 1, boards: [board([
    day('dMon', 'Monday', [slot('m1', '9:00 AM', 'Alex', 'Full'), slot('m2', '10:00 AM', 'Bo', 'Full')]),
    day('dTue', 'Tuesday', [slot('t1', '9:00 AM', 'Cy', 'Full')])
  ])], clients: [], dxCodes: [] };
  drive.bytes = await workbook.buildScheduleWorkbook(seed);
  drive.version = 1;
  const first = await call('PUT', '', { workspace: seed });
  assert.strictEqual(first.status, 200);
  assert.strictEqual(first.body.sharePoint.ok, true, JSON.stringify(first.body.sharePoint));
  ws = first.body.workspace;
  assert.strictEqual(ws.version, 1);
});

test('45 clients in one day write cleanly to SharePoint with intact tables', async () => {
  const mon = ws.boards[0].days.find((d) => d.name === 'Monday');
  for (let i = 0; i < 45; i++) mon.slots.push(slot('x' + i, (8 + (i % 10)) + ':00 AM', 'Client ' + i, 'Full'));
  const saved = await call('PUT', '', { workspace: ws, base: ws });
  assert.strictEqual(saved.body.sharePoint.ok, true, JSON.stringify(saved.body.sharePoint));
  ws = saved.body.workspace;
  const days = await sheetDays();
  assert.strictEqual(days.monday.length, 47);
  assert.strictEqual(days.tuesday.length, 1);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(drive.bytes);
  const sheet = wb.getWorksheet('Current Schedule');
  Object.values(sheet.tables).forEach((t) => {
    const ref = t.table.tableRef;
    const header = Number(ref.split(':')[0].replace(/\D/g, ''));
    assert.match(String(sheet.getRow(header).getCell(1).text), /Appointment Time/, 'table ' + t.table.name + ' sits on its header: ' + ref);
  });
});

test('blank slots stay on the dashboard and never turn into SharePoint rows', async () => {
  const mon = ws.boards[0].days.find((d) => d.name === 'Monday');
  mon.slots.push(slot('blank1', '', ''));
  mon.slots.push(slot('blank2', '', ''));
  const saved = await call('PUT', '', { workspace: ws, base: ws });
  ws = saved.body.workspace;
  assert.strictEqual((await sheetDays()).monday.length, 47);
  assert.strictEqual(ws.boards[0].days.find((d) => d.name === 'Monday').slots.length, 49);
});

test('an admin editing the SharePoint workbook shows up on the dashboard with the same slots', async () => {
  await editWorkbook((sheet) => {
    const row = findRow(sheet, 'Client 3');
    sheet.getRow(row).getCell(3).value = 'Client 3 (admin changed)';
    const bo = findRow(sheet, 'Bo');
    sheet.getRow(bo).getCell(2).value = 'Open 4x';
    sheet.getRow(bo).getCell(3).value = null;
  });
  const meta = await call('GET', 'meta');
  assert.notStrictEqual(meta.body.updatedAt, ws.updatedAt);
  const loaded = await call('GET', '');
  const mon = loaded.body.workspace.boards[0].days.find((d) => d.name === 'Monday');
  const byId = Object.fromEntries(mon.slots.map((s) => [s.id, s]));
  assert.strictEqual(byId.x3.notes, 'Client 3 (admin changed)');
  assert.strictEqual(byId.m2.status, 'Open 4x');
  assert.strictEqual(byId.m2.notes, '');
  assert.strictEqual(mon.slots.length, 49, 'nothing was lost, blank slots kept');
  assert.ok(byId.blank1 && byId.blank2);
  ws = loaded.body.workspace;
});

test('an admin adding and deleting rows is picked up', async () => {
  await editWorkbook((sheet) => {
    const last = findRow(sheet, 'Client 44');
    sheet.getRow(last).getCell(3).value = 'Client 44 (admin moved)';
    const del = findRow(sheet, 'Client 10');
    sheet.getRow(del).getCell(1).value = null;
    sheet.getRow(del).getCell(2).value = null;
    sheet.getRow(del).getCell(3).value = null;
  });
  const loaded = await call('GET', '');
  const mon = loaded.body.workspace.boards[0].days.find((d) => d.name === 'Monday');
  assert.ok(!mon.slots.some((s) => s.notes === 'Client 10'));
  assert.ok(mon.slots.some((s) => s.notes === 'Client 44 (admin moved)'));
  ws = loaded.body.workspace;
});

test('a provider edit made while an admin edit is waiting keeps both', async () => {
  await editWorkbook((sheet) => {
    sheet.getRow(findRow(sheet, 'Client 20')).getCell(3).value = 'Client 20 admin';
  });
  const mine = JSON.parse(JSON.stringify(ws));
  const mon = mine.boards[0].days.find((d) => d.name === 'Monday');
  mon.slots.find((s) => s.id === 'x30').notes = 'Client 30 provider';
  mon.slots.push(slot('new1', '3:00 PM', 'New client'));
  const saved = await call('PUT', '', { workspace: mine, base: ws });
  assert.strictEqual(saved.body.merged, true);
  const out = saved.body.workspace.boards[0].days.find((d) => d.name === 'Monday').slots;
  assert.ok(out.some((s) => s.notes === 'Client 20 admin'));
  assert.ok(out.some((s) => s.notes === 'Client 30 provider'));
  assert.ok(out.some((s) => s.notes === 'New client'));
  const days = await sheetDays();
  const notes = days.monday.map((r) => r.notes);
  assert.ok(notes.includes('Client 20 admin') && notes.includes('Client 30 provider') && notes.includes('New client'));
  ws = saved.body.workspace;
});

test('an admin save landing during the provider save is not overwritten', async () => {
  drive.beforePut = async () => {
    await editWorkbook((sheet) => {
      sheet.getRow(findRow(sheet, 'Client 5')).getCell(3).value = 'Client 5 racing admin';
    });
  };
  const mine = JSON.parse(JSON.stringify(ws));
  mine.boards[0].days.find((d) => d.name === 'Monday').slots.find((s) => s.id === 'x7').notes = 'Client 7 provider';
  const saved = await call('PUT', '', { workspace: mine, base: ws });
  assert.strictEqual(saved.body.sharePoint.ok, true, JSON.stringify(saved.body.sharePoint));
  const notes = (await sheetDays()).monday.map((r) => r.notes);
  assert.ok(notes.includes('Client 5 racing admin'), 'admin change survived');
  assert.ok(notes.includes('Client 7 provider'), 'provider change written');
  const dash = saved.body.workspace.boards[0].days.find((d) => d.name === 'Monday').slots.map((s) => s.notes);
  assert.ok(dash.includes('Client 5 racing admin'));
  ws = saved.body.workspace;
});

test('two tabs saving at once do not delete each other', async () => {
  const tabA = JSON.parse(JSON.stringify(ws));
  const tabB = JSON.parse(JSON.stringify(ws));
  tabA.boards[0].days[0].slots.push(slot('a-new', '1:00 PM', 'From tab A'));
  tabB.boards[0].days[0].slots.push(slot('b-new', '2:00 PM', 'From tab B'));
  const a = await call('PUT', '', { workspace: tabA, base: ws });
  const b = await call('PUT', '', { workspace: tabB, base: ws });
  const final = (await call('GET', '')).body.workspace.boards[0].days[0].slots.map((s) => s.notes);
  assert.ok(final.includes('From tab A') && final.includes('From tab B'));
  assert.strictEqual(a.status, 200);
  assert.strictEqual(b.body.merged, true);
});

test('saving an unchanged schedule does not rewrite the SharePoint file', async () => {
  const current = (await call('GET', '')).body.workspace;
  await call('PUT', '', { workspace: current, base: current });
  const before = drive.putCount;
  const again = (await call('GET', '')).body.workspace;
  await call('PUT', '', { workspace: again, base: again });
  assert.strictEqual(drive.putCount, before);
});

test('odd times and every status survive a round trip through SharePoint', async () => {
  const current = (await call('GET', '')).body.workspace;
  const wed = { id: 'dWed', name: 'Wednesday', color: 'blue', slots: [
    slot('w1', '4:15 PM', 'A', 'Lunch'),
    slot('w2', '12:05 PM', 'B', 'Scheduling In Progress'),
    slot('w3', '11:59 AM', 'C', 'Open 3x (1,2,3)'),
    slot('w4', '9:20 AM', 'D', 'Open 2x (1&3)'),
    slot('w5', '12:00 AM', 'E', 'Makeup')
  ] };
  current.boards[0].days.push(wed);
  const saved = await call('PUT', '', { workspace: current, base: current });
  assert.strictEqual(saved.body.sharePoint.ok, true, JSON.stringify(saved.body.sharePoint));
  await editWorkbook((sheet) => {
    sheet.getRow(findRow(sheet, 'Client 1')).getCell(3).value = 'Client 1 touched';
  });
  const loaded = (await call('GET', '')).body.workspace;
  const day = loaded.boards[0].days.find((d) => d.name === 'Wednesday');
  assert.deepStrictEqual(day.slots.map((s) => [s.id, s.time, s.status]), wed.slots.map((s) => [s.id, s.time, s.status]));
  assert.ok(loaded.boards[0].days[0].slots.some((s) => s.notes === 'Client 1 touched'));
});

test('a day an admin adds to the workbook shows up, and a day only on the dashboard is added to the workbook', async () => {
  await editWorkbook((sheet) => {
    let row = 1;
    for (let r = 1; r <= sheet.rowCount; r++) if (String(sheet.getRow(r).getCell(1).text || '').trim()) row = r;
    const start = row + 2;
    sheet.addTable({
      name: 'Thursday',
      ref: 'A' + start,
      headerRow: true,
      columns: [{ name: 'Thurs Appointment Time ' }, { name: 'Status ' }, { name: 'Notes (i.e. client names)' }, { name: 'Date of Last Visit' }],
      rows: [['3:00 PM', 'Full', 'Admin added', '']]
    });
  });
  const loaded = (await call('GET', '')).body.workspace;
  const thu = loaded.boards[0].days.find((d) => d.name === 'Thursday');
  assert.ok(thu && thu.slots.some((s) => s.notes === 'Admin added'));
  const current = JSON.parse(JSON.stringify(loaded));
  current.boards[0].days.push({ id: 'dFri', name: 'Friday', color: 'blue', slots: [slot('f1', '4:30 PM', 'Friday client', 'Full')] });
  const saved = await call('PUT', '', { workspace: current, base: loaded });
  assert.strictEqual(saved.body.sharePoint.ok, true, JSON.stringify(saved.body.sharePoint));
  const days = await sheetDays();
  assert.strictEqual(days.friday[0].notes, 'Friday client');
  assert.strictEqual(days.thursday[0].notes, 'Admin added');
});

test('a provider opening the dashboard for the first time gets the schedule already in SharePoint', async () => {
  require('./fakes').blobs.clear();
  const loaded = await call('GET', '');
  const names = loaded.body.workspace.boards.map((b) => b.name);
  assert.ok(names.includes('Current Schedule'));
  const mon = loaded.body.workspace.boards.find((b) => b.name === 'Current Schedule').days.find((d) => d.name === 'Monday');
  assert.ok(mon.slots.length >= 40, 'monday clients came over: ' + mon.slots.length);
  const putsBefore = drive.putCount;
  await call('PUT', '', { workspace: loaded.body.workspace, base: loaded.body.workspace });
  assert.strictEqual(drive.putCount, putsBefore, 'nothing needed to be written back');
});

test('a SharePoint write that fails is retried on the next sync check', async () => {
  const current = (await call('GET', '')).body.workspace;
  const edited = JSON.parse(JSON.stringify(current));
  edited.boards[0].days[0].slots.find((s) => s.notes === 'Client 12').notes = 'Client 12 saved while file was locked';
  drive.failPuts = 1;
  const saved = await call('PUT', '', { workspace: edited, base: current });
  assert.strictEqual(saved.body.sharePoint.ok, false);
  assert.ok(!(await sheetDays()).monday.some((r) => r.notes === 'Client 12 saved while file was locked'));
  const meta = await call('GET', 'meta');
  assert.strictEqual(meta.status, 200);
  assert.ok((await sheetDays()).monday.some((r) => r.notes === 'Client 12 saved while file was locked'), 'workbook caught up');
  const after = (await call('GET', '')).body.workspace.boards[0].days[0].slots.map((s) => s.notes);
  assert.ok(after.includes('Client 12 saved while file was locked'), 'dashboard kept the edit');
});
