const test = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('exceljs');
const { blobs } = require('./fakes');

const ROOT = 'LEADERSHIP TEAM/EMME Processes/Referral Process/Provider Schedules';

test('open slots are read from the dashboard schedules, with no notes or client names', async () => {
  const workspace = {
    ownerEmail: 'ndelgado@emmeco.org',
    updatedAt: new Date().toISOString(),
    boards: [{ id: 'b1', name: 'Current Schedule', days: [
      { id: 'd2', name: 'Tuesday', slots: [{ time: '10:00 AM', status: 'Open 3x (1,2,3)', notes: '' }, { time: '11:00 AM', status: 'Full', notes: 'Made Up Client' }] },
      { id: 'd1', name: 'Monday', slots: [{ time: '9:00 AM', status: 'Open 4x', notes: '' }, { time: '12:00 PM', status: 'Lunch/Paperwork', notes: '' }] },
    ] }],
    clients: [{ name: 'Made Up Client' }],
  };
  blobs.set('schedules/by-email/ndelgado_emmeco.org.json', { data: JSON.stringify(workspace), meta: { owner: 'ndelgado@emmeco.org' } });
  const { openingsFromDashboard } = require('../dist/schedulingService');
  const found = await openingsFromDashboard();
  assert.equal(found.configured, true);
  assert.equal(found.staff.length, 1);
  const provider = found.staff[0];
  assert.equal(provider.email, 'ndelgado@emmeco.org');
  assert.deepEqual(provider.summary.openings, [
    { day: 'Monday', time: '9:00 AM', status: 'Open 4x' },
    { day: 'Tuesday', time: '10:00 AM', status: 'Open 3x (1,2,3)' },
  ]);
  assert.ok(!JSON.stringify(found).includes('Made Up Client'));
});

const written = {};
function sharePoint(existing) {
  global.fetch = async function (url, options) {
    const text = decodeURIComponent(String(url));
    const reply = (status, body) => new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
    if (text.includes('/oauth2/v2.0/token')) return reply(200, { access_token: 't', expires_in: 3600 });
    if (text.includes('/personal/jenriat_emmeco_org')) return reply(200, { id: 'site1' });
    if (text.includes('/drive/root:/' + ROOT + ':/children')) return reply(200, { value: [] });
    if (text.includes('/drive/root:/' + ROOT + '/') && options && options.method === 'PUT') {
      written.url = text;
      written.body = Buffer.from(options.body);
      written.count = (written.count || 0) + 1;
      return reply(201, { eTag: '"{1},1"', lastModifiedDateTime: new Date().toISOString() });
    }
    if (existing && text.includes('/' + existing + '?')) return reply(200, { eTag: '"{1},1"', lastModifiedDateTime: new Date().toISOString() });
    return reply(404, { error: { code: 'itemNotFound', message: 'no' } });
  };
}

test('Create from template gives a new provider the standard schedule and a read-only Dashboard Copy', async () => {
  sharePoint('');
  const { createScheduleFromTemplate, readScheduleDays } = require('../dist/scheduleWorkbook');
  const result = await createScheduleFromTemplate('kelly.example@emmeco.org', 'Kelly Example');
  assert.equal(result.ok, true, result.message);
  assert.equal(result.created, true);
  assert.equal(result.seeded, true);
  assert.equal(result.fileName, 'Kelly Example - Dashboard Copy.xlsx');
  assert.equal(written.count, 1);
  assert.match(written.url, /Kelly Example - Dashboard Copy\.xlsx/);
  assert.match(written.url, /conflictBehavior=fail/);

  // the provider's dashboard now holds the standard schedule
  const stored = JSON.parse(blobs.get('schedules/by-email/kelly.example_emmeco.org.json').data);
  assert.equal(stored.ownerEmail, 'kelly.example@emmeco.org');
  assert.equal(stored.boards[0].days.length, 5);
  assert.equal(stored.boards[0].days[0].slots.length, 11);

  // the file: Current Schedule first, five days, status drop-down, locked, with the read-only note
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(written.body);
  const sheet = wb.worksheets[0];
  assert.equal(sheet.name, 'Current Schedule');
  assert.equal(sheet.getCell('B2').dataValidation && sheet.getCell('B2').dataValidation.type, 'list');
  assert.ok(sheet.sheetProtection, 'the sheet is protected');
  assert.match(String(sheet.getCell('H1').value), /READ-ONLY COPY/);
  const days = await readScheduleDays(new Uint8Array(written.body));
  assert.deepEqual(Object.keys(days).sort(), ['friday', 'monday', 'thursday', 'tuesday', 'wednesday']);
  assert.match(days.monday.find((row) => row.time.startsWith('12')).status, /Lunch/);
});

test('Create from template never replaces a schedule or a file that is already there', async () => {
  const before = blobs.get('schedules/by-email/kelly.example_emmeco.org.json').data;
  const stored = JSON.parse(before);
  stored.boards[0].days[0].slots[0].notes = 'Made Up Note';
  blobs.set('schedules/by-email/kelly.example_emmeco.org.json', { data: JSON.stringify(stored), meta: { owner: 'kelly.example@emmeco.org' } });
  written.count = 0;
  sharePoint('Kelly Example - Dashboard Copy.xlsx');
  const { createScheduleFromTemplate } = require('../dist/scheduleWorkbook');
  const result = await createScheduleFromTemplate('kelly.example@emmeco.org', 'Kelly Example');
  assert.equal(result.ok, true);
  assert.equal(result.created, false);
  assert.equal(result.seeded, false);
  assert.equal(written.count, 0, 'nothing was written to SharePoint');
  assert.match(blobs.get('schedules/by-email/kelly.example_emmeco.org.json').data, /Made Up Note/);
});

test('Create from template needs a real Emme email', async () => {
  const { createScheduleFromTemplate } = require('../dist/scheduleWorkbook');
  const result = await createScheduleFromTemplate('someone@example.com', 'Someone');
  assert.equal(result.ok, false);
});
