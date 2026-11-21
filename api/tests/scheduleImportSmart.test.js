const test = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('exceljs');
const lib = require('../../schedule-import.js');
const workbook = require('../dist/scheduleWorkbook.js');

// All names below are made up.

function brief(result, day) {
  return (result.days[day] || []).map((s) => [s.time, s.status, s.notes, s.lastDay].join(' | '));
}

test('statuses are recognised however they are written', () => {
  assert.equal(lib.canonicalStatus('FULL'), 'Full');
  assert.equal(lib.canonicalStatus('open 4x'), 'Open 4x');
  assert.equal(lib.canonicalStatus('Open 3x (2,3,4)'), 'Open 3x (2,3,4)');
  assert.equal(lib.canonicalStatus('open 3x 1/3/4'), 'Open 3x (1,3,4)');
  assert.equal(lib.canonicalStatus('Open 2x (2 & 4)'), 'Open 2x (2&4)');
  assert.equal(lib.canonicalStatus('Open 1x wk 3'), 'Open 1x (3)');
  assert.equal(lib.canonicalStatus('Lunch/Paperwork'), 'Lunch');
  assert.equal(lib.canonicalStatus('make-up'), 'Makeup');
  assert.equal(lib.canonicalStatus('Scheduling in Progress'), 'Scheduling In Progress');
  assert.equal(lib.canonicalStatus('Not open'), 'Not Open');
  assert.equal(lib.canonicalStatus('Jamie P'), null);
});

test('times are read in many forms', () => {
  assert.equal(lib.findTime('9am').text, '9:00 AM');
  assert.equal(lib.findTime('9:30 a.m.').text, '9:30 AM');
  assert.equal(lib.findTime('1:30').text, '1:30 PM');
  assert.equal(lib.findTime('13:00').text, '1:00 PM');
  assert.equal(lib.findTime('9-10 AM').text, '9:00 AM - 10:00 AM');
  assert.equal(lib.findTime('11-12 PM').text, '11:00 AM - 12:00 PM');
});

test('one table per weekday with headings (the Provider Schedules layout)', () => {
  const rows = [
    ['Mon Appointment Time ', 'Status', 'Client Name', 'Date of Last Visit'],
    ['9:00 AM', 'Full', 'Jamie P', '05/01/2026'],
    ['10:00 AM', 'Open 3x (1,2,3)', '', ''],
    [],
    ['Tues Appointment Time ', 'Status', 'Client Name', 'Date of Last Visit'],
    ['9:00 AM', 'Lunch/Paperwork', '', '']
  ];
  const r = lib.scheduleFromMatrix(rows);
  assert.deepEqual(brief(r, 'monday'), ['9:00 AM | Full | Jamie P | 05/01/2026', '10:00 AM | Open 3x (1,2,3) |  | ']);
  assert.deepEqual(brief(r, 'tuesday'), ['9:00 AM | Lunch |  | ']);
});

test('a Word or PDF document with day headings and plain lines', () => {
  const text = [
    'Sam Example - Weekly Schedule',
    '',
    'MONDAY',
    '9:00 AM  Full  Jamie P',
    '10:00 AM - Open 4x',
    '11:00 AM Riley T',
    '',
    'Tuesday:',
    '8am Lunch',
    '2:30 pm Open 2x (1&3)'
  ].join('\n');
  const r = lib.scheduleFromText(text);
  assert.deepEqual(brief(r, 'monday'), ['9:00 AM | Full | Jamie P | ', '10:00 AM | Open 4x |  | ', '11:00 AM | Full | Riley T | ']);
  assert.deepEqual(brief(r, 'tuesday'), ['8:00 AM | Lunch |  | ', '2:30 PM | Open 2x (1&3) |  | ']);
  assert.ok(lib.looksLikeSchedule(r));
});

test('one line per appointment with the weekday in the line', () => {
  const r = lib.scheduleFromText('Mon 9am Full Jamie P 5/1/2026\nMon 10am open 4x\nWednesday 1:00 PM Makeup\nthe office is closed on holidays');
  assert.deepEqual(brief(r, 'monday'), ['9:00 AM | Full | Jamie P | 5/1/2026', '10:00 AM | Open 4x |  | ']);
  assert.deepEqual(brief(r, 'wednesday'), ['1:00 PM | Makeup |  | ']);
});

test('a table with a Day column, in any column order', () => {
  const rows = [
    ['Client', 'Day', 'Status', 'Time'],
    ['Jamie P', 'Monday', 'Full', '9-10 AM'],
    ['', 'Friday', 'Open 4x', '3:00 PM']
  ];
  const r = lib.scheduleFromMatrix(rows);
  assert.deepEqual(brief(r, 'monday'), ['9:00 AM - 10:00 AM | Full | Jamie P | ']);
  assert.deepEqual(brief(r, 'friday'), ['3:00 PM | Open 4x |  | ']);
});

test('a grid with weekdays across the top and times down the side', () => {
  const rows = [
    ['', 'Monday', 'Tuesday', 'Wednesday'],
    ['9:00 AM', 'Full - Sam B', 'Open 4x', 'Lunch'],
    ['10:00 AM', '', 'Riley T', 'Open 3x (1,2,3)'],
    ['', '', '', '']
  ];
  const r = lib.scheduleFromMatrix(rows);
  assert.deepEqual(brief(r, 'monday'), ['9:00 AM | Full | Sam B | ']);
  assert.deepEqual(brief(r, 'tuesday'), ['9:00 AM | Open 4x |  | ', '10:00 AM | Full | Riley T | ']);
  assert.deepEqual(brief(r, 'wednesday'), ['9:00 AM | Lunch |  | ', '10:00 AM | Open 3x (1,2,3) |  | ']);
});

test('spreadsheet times that arrive as plain numbers still work', () => {
  const rows = [['Monday'], ['0.375', 'Full', 'Jamie P'], ['0.416667', 'Open 4x'], ['0.458333', 'Full', 'Riley T']];
  const r = lib.scheduleFromMatrix(rows);
  assert.deepEqual(brief(r, 'monday').map((x) => x.split(' | ')[0]), ['9:00 AM', '10:00 AM', '11:00 AM']);
});

test('a client list is not mistaken for a schedule', () => {
  const rows = [['Name', 'Email', 'Phone'], ['Jamie Parent', 'jamie@example.com', '555-123-4567'], ['Riley Tester', 'riley@example.com', '555-222-3333']];
  assert.ok(!lib.looksLikeSchedule(lib.scheduleFromMatrix(rows)));
});

test('the sheet called Current Schedule wins over other sheets', () => {
  const pick = lib.bestSheet([
    { name: 'Summer 26', rows: [['Monday'], ['9:00 AM', 'Full', 'A One'], ['10:00 AM', 'Full', 'B Two'], ['11:00 AM', 'Full', 'C Three'], ['12:00 PM', 'Full', 'D Four']] },
    { name: 'Current Schedule', rows: [['Monday'], ['9:00 AM', 'Open 4x'], ['10:00 AM', 'Open 4x'], ['11:00 AM', 'Open 4x']] },
    { name: 'Lists', rows: [['Open 4x'], ['Full']] }
  ]);
  assert.equal(pick.name, 'Current Schedule');
});

// The browser reads cells as the text Excel shows (SheetJS raw:false); this does the same for a time cell.
function cellText(cell) {
  const v = cell.value;
  if (v instanceof Date) {
    const h = v.getUTCHours();
    return (h % 12 || 12) + ':' + String(v.getUTCMinutes()).padStart(2, '0') + ' ' + (h >= 12 ? 'PM' : 'AM');
  }
  return String(cell.text || '');
}

test('a workbook the dashboard itself wrote reads back the same as the exact reader', async () => {
  const ws = {
    version: 1,
    boards: [{ id: 'b', name: 'Current Schedule', days: [
      { id: 'd1', name: 'Monday', slots: [{ id: 's1', time: '9:00 AM', status: 'Full', notes: 'Made Up Client', lastDay: '' }, { id: 's2', time: '10:30 AM', status: 'Open 3x (2,3,4)', notes: '', lastDay: '' }] },
      { id: 'd2', name: 'Thursday', slots: [{ id: 's3', time: '1:00 PM', status: 'Makeup', notes: 'Another Made Up', lastDay: '' }] }
    ] }],
    clients: [], dxCodes: []
  };
  const bytes = await workbook.buildScheduleWorkbook(ws);
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(bytes);
  const sheets = book.worksheets.map((sheet) => {
    const rows = [];
    sheet.eachRow({ includeEmpty: true }, (row) => {
      const cells = [];
      row.eachCell({ includeEmpty: true }, (cell, col) => { cells[col - 1] = cellText(cell); });
      rows.push(cells);
    });
    return { name: sheet.name, rows };
  });
  const r = lib.bestSheet(sheets);
  assert.equal(r.name, 'Current Schedule');
  const exact = await workbook.readScheduleDays(bytes);
  ['monday', 'thursday'].forEach((day) => {
    assert.deepEqual(
      r.days[day].map((s) => [s.time, s.notes]),
      exact[day].map((s) => [s.time, s.notes]),
      day
    );
  });
  assert.deepEqual(r.days.monday.map((s) => s.status), ['Full', 'Open 3x (2,3,4)']);
  assert.deepEqual(r.days.thursday.map((s) => s.status), ['Makeup']);
});
