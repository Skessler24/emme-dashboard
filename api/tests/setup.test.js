const test = require('node:test');
const assert = require('node:assert');

process.env.GRAPH_CLIENT_SECRET = 'fake';
process.env.SHAREPOINT_HOSTNAME = 'emmeco031.sharepoint.com';
process.env.SHAREPOINT_SITE_PATH = '/sites/Emme';
process.env.EMME_ADMIN_EMAILS = 'skessler@emmeco.org';

const SCHEDULES = 'EMME/LEADERSHIP TEAM/EMME Processes/Referral Process/Provider Schedules';
const MILEAGE = 'EMME/PAYROLL & MILEAGE/Mileage';

const lists = new Map();
const folders = { shared1: new Set(), personal1: new Set() };
const calls = [];

function reply(status, body) {
  return new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
}

global.fetch = async function (url, options) {
  const method = (options && options.method) || 'GET';
  const text = decodeURIComponent(String(url));
  if (text.includes('/oauth2/v2.0/token')) return reply(200, { access_token: 't', expires_in: 3600 });
  if (text.includes('/sites/emmeco031.sharepoint.com:/sites/Emme')) return reply(200, { id: 'shared1' });
  if (text.includes('/personal/jenriat_emmeco_org')) return reply(200, { id: 'personal1' });
  let match = text.match(/\/sites\/(\w+)\?\$select=displayName,webUrl/);
  if (match) return reply(200, { displayName: 'Emme', webUrl: 'https://emmeco031.sharepoint.com/sites/Emme' });
  match = text.match(/\/sites\/(\w+)\/drive\/root:\/([^?]*?)(?::\/children)?(?:\?|$)/);
  if (match) {
    return folders[match[1]] && folders[match[1]].has(match[2])
      ? reply(200, { id: 'folder', folder: {}, value: [] })
      : reply(404, { error: { code: 'itemNotFound', message: 'missing' } });
  }
  if (method === 'POST' && /\/sites\/shared1\/lists$/.test(text)) {
    const body = JSON.parse(options.body);
    calls.push('create ' + body.displayName);
    lists.set(body.displayName, { id: 'id-' + body.displayName, columns: ['Title'].concat(body.columns.map((c) => c.name)) });
    return reply(201, { id: 'id-' + body.displayName });
  }
  match = text.match(/\/sites\/shared1\/lists\/([^/?]+)\/columns$/);
  if (match && method === 'POST') {
    const list = [...lists.values()].find((l) => l.id === match[1]);
    const body = JSON.parse(options.body);
    calls.push('column ' + match[1] + ' ' + body.name);
    list.columns.push(body.name);
    return reply(201, { id: 'c' });
  }
  match = text.match(/\/sites\/shared1\/lists\/([^/?]+)\?/);
  if (match && method === 'GET') {
    const list = lists.get(match[1]);
    if (!list) return reply(404, { error: { code: 'itemNotFound', message: 'no list' } });
    return reply(200, { id: list.id, columns: list.columns.map((name) => ({ name })) });
  }
  return reply(404, { error: { code: 'itemNotFound', message: 'unexpected ' + text } });
};

const setupHandler = require('../setup/index.js');
const { scheduleSiteId } = require('../dist/scheduleWorkbook');
const { LIST_SCHEMAS } = require('../dist/siteSetup');

function asAdmin(method, email) {
  const principal = Buffer.from(JSON.stringify({ userId: 'u', userDetails: email, claims: [{ typ: 'preferred_username', val: email }] })).toString('base64');
  return { method, headers: { 'x-ms-client-principal': principal }, params: {} };
}

async function run(method, email) {
  const context = {};
  await setupHandler(context, asAdmin(method, email || 'skessler@emmeco.org'));
  return context.res;
}

test('only admins can open the setup check', async () => {
  const res = await run('GET', 'ndelgado@emmeco.org');
  assert.strictEqual(res.status, 403);
  assert.strictEqual(calls.length, 0);
});

test('the check reports what is missing and where folders are', async () => {
  folders.personal1.add(SCHEDULES);
  folders.personal1.add(MILEAGE);
  const res = await run('GET');
  assert.strictEqual(res.status, 200);
  assert.strictEqual(res.body.site.name, 'Emme');
  assert.strictEqual(res.body.lists.length, 9);
  assert.ok(res.body.lists.every((l) => !l.exists));
  assert.strictEqual(res.body.ready, false);
  const schedules = res.body.folders.find((f) => f.label === 'Provider Schedules');
  assert.deepStrictEqual([schedules.onSharedSite, schedules.onPersonalOneDrive], [false, true]);
  assert.strictEqual(res.body.usingPersonalOneDrive, true);
  assert.ok(res.body.nextSteps.some((s) => /Move these folders/.test(s)));
});

test('create missing lists builds all nine and is safe to repeat', async () => {
  const first = await run('POST');
  assert.strictEqual(first.status, 200);
  assert.strictEqual(first.body.created.createdLists.length, 9);
  assert.ok(first.body.status.lists.every((l) => l.exists && l.missingColumns.length === 0));
  const before = calls.length;
  const second = await run('POST');
  assert.strictEqual(second.body.created.createdLists.length, 0);
  assert.strictEqual(calls.length, before);
});

test('a list that lost a column gets only that column back', async () => {
  const list = lists.get('Mileage');
  list.columns = list.columns.filter((name) => name !== 'Miles');
  const res = await run('POST');
  assert.deepStrictEqual(res.body.created.addedColumns, [{ list: 'Mileage', columns: ['Miles'] }]);
});

test('list definitions cover every field the dashboard writes', () => {
  const columns = (name) => LIST_SCHEMAS.find((s) => s.list === name).columns.map((c) => c.name);
  assert.deepStrictEqual(columns('Mileage'), ['StaffEmail', 'StaffName', 'TripDate', 'Miles', 'Purpose', 'Status']);
  assert.ok(columns('Cancellations').includes('ExternalId'));
  assert.ok(columns('DocumentationCompliance').includes('OpenReports'));
});

test('schedules are read from the shared Emme site once folders exist there', async () => {
  folders.shared1.add(SCHEDULES);
  assert.strictEqual(await scheduleSiteId(), 'shared1');
});

test('a folder that is not found is looked for again, and a moved folder is found within minutes', async () => {
  const { forgetFolderSites } = require('../dist/graphClient');
  forgetFolderSites();
  folders.shared1.delete(SCHEDULES);
  folders.personal1.delete(SCHEDULES);
  // Missing everywhere: the first site is used, but that answer is not remembered.
  assert.strictEqual(await scheduleSiteId(), 'shared1');
  folders.personal1.add(SCHEDULES);
  assert.strictEqual(await scheduleSiteId(), 'personal1');
  // Remembered while fresh, so a busy minute does not repeat the lookups.
  folders.shared1.add(SCHEDULES);
  assert.strictEqual(await scheduleSiteId(), 'personal1');
  // After the memory runs out, the moved folder is found on the shared site.
  const realNow = Date.now;
  Date.now = () => realNow() + 6 * 60 * 1000;
  try {
    assert.strictEqual(await scheduleSiteId(), 'shared1');
  } finally {
    Date.now = realNow;
  }
});
