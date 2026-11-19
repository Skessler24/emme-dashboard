const test = require('node:test');
const assert = require('node:assert/strict');
require('./fakes');

const ROOT = 'LEADERSHIP TEAM/EMME Processes/Referral Process/Provider Schedules';

function listing(names) {
  global.fetch = async function (url) {
    const text = decodeURIComponent(String(url));
    const reply = (status, body) => new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
    if (text.includes('/oauth2/v2.0/token')) return reply(200, { access_token: 't', expires_in: 3600 });
    if (text.includes('/personal/jenriat_emmeco_org')) return reply(200, { id: 'site1' });
    if (text.includes('/drive/root:/' + ROOT + ':/children')) return reply(200, { value: names.map(function (name) { return { name, file: {} }; }) });
    return reply(404, { error: { code: 'itemNotFound', message: 'no' } });
  };
}

test('workbooks kept straight in Provider Schedules are found by first name', async () => {
  const { findExistingProviderFile, SCHEDULE_ROOT } = require('../dist/providerSetup');
  listing(['Hannah Schedule.xlsx', 'Kelly Schedule.xlsx', 'Toby Schedule.xlsx']);
  const found = await findExistingProviderFile(SCHEDULE_ROOT, 'kelly.example@emmeco.org', 'Kelly Example', 'schedule');
  assert.equal(found.fileName, 'Kelly Schedule.xlsx');
  assert.equal(found.folder, '');
  assert.equal(found.path, '/' + ROOT + '/Kelly Schedule.xlsx');
});

test('initials in the file name also work', async () => {
  const { findExistingProviderFile, SCHEDULE_ROOT } = require('../dist/providerSetup');
  listing(['Hannah Schedule.xlsx', 'SB Schedule.xlsx']);
  const found = await findExistingProviderFile(SCHEDULE_ROOT, 'sam.b@emmeco.org', 'Sam Bishop', 'schedule');
  assert.equal(found.fileName, 'SB Schedule.xlsx');
});

test('a first-name file beats an initials file for the same person', async () => {
  const { findExistingProviderFile, SCHEDULE_ROOT } = require('../dist/providerSetup');
  listing(['SB Schedule.xlsx', 'Sam Schedule.xlsx', 'Toby Schedule.xlsx']);
  const found = await findExistingProviderFile(SCHEDULE_ROOT, 'sam.b@emmeco.org', 'Sam Bishop', 'schedule');
  assert.equal(found.fileName, 'Sam Schedule.xlsx');
});

test('two files that fit one person are not guessed between', async () => {
  const { findExistingProviderFile, SCHEDULE_ROOT } = require('../dist/providerSetup');
  listing(['KE Schedule.xlsx', 'Kelly E Schedule.xlsx']);
  const found = await findExistingProviderFile(SCHEDULE_ROOT, 'kelly.example@emmeco.org', 'Kelly Example', 'schedule');
  assert.equal(found, null);
});

test('an unrelated name finds nothing', async () => {
  const { findExistingProviderFile, SCHEDULE_ROOT } = require('../dist/providerSetup');
  listing(['Hannah Schedule.xlsx', 'Toby Schedule.xlsx']);
  const found = await findExistingProviderFile(SCHEDULE_ROOT, 'kelly.example@emmeco.org', 'Kelly Example', 'schedule');
  assert.equal(found, null);
});
