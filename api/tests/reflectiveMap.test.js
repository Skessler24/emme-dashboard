const test = require('node:test');
const assert = require('node:assert');
const { blobs } = require('./fakes');

process.env.EMME_ADMIN_EMAILS = 'boss@emmeco.org';
process.env.AUTOMATION_API_KEY = 'test-automation-key';

const handler = require('../reflective-supervision/index.js');
const { blankMap, buildGrowthMapPdf, clientVisitTimes, matchingFolder, monthStanding, plusHour } = require('../dist/reflectiveMap');
const { resetReflectiveOfficeCache } = require('../dist/reflectiveOffice');

const ADMIN = 'boss@emmeco.org';
const PROVIDER = 'pat@emmeco.org';

function principal(email) {
  return Buffer.from(JSON.stringify({
    userId: 'u-' + email,
    userDetails: email,
    claims: [{ typ: 'preferred_username', val: email }, { typ: 'name', val: email === ADMIN ? 'Boss Admin' : 'Pat Provider' }],
  })).toString('base64');
}

async function call(email, method, query, body, extraHeaders) {
  const context = {};
  const headers = Object.assign({}, email ? { 'x-ms-client-principal': principal(email) } : {}, extraHeaders || {});
  await handler(context, { method, query: query || {}, body, headers });
  return context.res;
}

function reset() {
  for (const key of [...blobs.keys()]) {
    if (key.startsWith('reflective-supervision/') || key.startsWith('schedules/')) blobs.delete(key);
  }
  resetReflectiveOfficeCache();
}

test('one confirmed session a month counts, and a missed month is behind', () => {
  const done = monthStanding([
    { status: 'confirmed', start: '2026-11-10T09:00', end: '2026-11-10T10:00' },
  ], '2026-11-25');
  assert.strictEqual(done.short, 'Done');
  assert.strictEqual(done.hours, 1);
  const needed = monthStanding([
    { status: 'confirmed', start: '2026-10-10T09:00', end: '2026-10-10T10:00' },
  ], '2026-11-25');
  assert.strictEqual(needed.short, 'Still needed');
  const behind = monthStanding([
    { status: 'confirmed', start: '2026-09-10T09:00', end: '2026-09-10T10:30' },
  ], '2026-11-25');
  assert.strictEqual(behind.short, 'Behind');
  assert.strictEqual(behind.hours, 0);
  const fresh = monthStanding([], '2026-11-25');
  assert.strictEqual(fresh.short, 'Still needed');
  const scheduled = monthStanding([
    { status: 'confirmed', start: '2026-10-02T09:00', end: '2026-10-02T10:00' },
    { status: 'proposed', start: '2026-11-20T15:00', end: '2026-11-20T16:00' },
  ], '2026-11-25');
  assert.strictEqual(scheduled.short, 'Scheduled');
});

test('a new growth map is the blank form, with none of the notes from a filled copy', () => {
  const map = blankMap(PROVIDER, 'Pat Provider');
  assert.strictEqual(map.topics.length, 16);
  assert.strictEqual(map.topics[0].title, 'Observation before interpretation');
  assert.strictEqual(map.topics[0].foundation, true);
  assert.strictEqual(map.topics[0].e, '');
  const pdf = buildGrowthMapPdf(map).toString('binary');
  assert.ok(pdf.startsWith('%PDF-1.4'));
  assert.ok(pdf.includes('Observation before interpretation'));
  assert.ok(!pdf.includes('being okay'));
  assert.strictEqual(plusHour('2026-11-25T09:30'), '2026-11-25T10:30');
});

test('client visit times come from booked slots and leave out the notes', () => {
  const visits = clientVisitTimes({
    boards: [{
      name: 'Current Schedule',
      days: [{
        name: 'Monday',
        slots: [
          { time: '9:00 AM', status: 'Full', notes: 'private client note' },
          { time: '10:00 AM', status: 'Open 4x' },
          { time: '11:00 AM', status: 'Makeup', notes: 'another private note' },
        ],
      }],
    }],
  }, 7, '2026-11-23');
  assert.deepStrictEqual(visits.map((item) => item.label + ' ' + item.time), ['With a client 9:00 AM', 'Makeup visit 11:00 AM']);
  assert.ok(!JSON.stringify(visits).includes('private'));
});

test('a provider folder is chosen by name inside RS Providers', () => {
  const folders = [{ id: '1', name: 'Pat Provider' }, { id: '2', name: 'Someone Else' }];
  assert.strictEqual(matchingFolder(folders, 'Pat Provider', PROVIDER).id, '1');
  assert.strictEqual(matchingFolder(folders, 'No Match', 'nomatch@emmeco.org'), null);
});

test('a provider can save their own map, and cannot open someone else', async () => {
  reset();
  const first = blankMap(PROVIDER, 'Pat Provider');
  first.topics[0].e = 'I noticed the child reach first.';
  const saved = await call(PROVIDER, 'PUT', { part: 'map' }, { baseUpdatedAt: '', map: first });
  assert.strictEqual(saved.status, 200);
  assert.strictEqual(saved.body.map.topics[0].e, 'I noticed the child reach first.');
  assert.strictEqual(saved.body.map.topics[0].title, 'Observation before interpretation');
  const again = await call(PROVIDER, 'GET', { part: 'map' });
  assert.strictEqual(again.body.map.topics[0].e, 'I noticed the child reach first.');
  const denied = await call(PROVIDER, 'GET', { part: 'map', provider: ADMIN });
  assert.strictEqual(denied.status, 403);
  const pdf = await call(PROVIDER, 'GET', { part: 'pdf' });
  assert.strictEqual(pdf.status, 200);
  assert.ok(Buffer.isBuffer(pdf.body) || typeof pdf.body === 'string' || pdf.body instanceof Uint8Array);
  const text = Buffer.from(pdf.body).toString('binary');
  assert.ok(text.includes('I noticed the child reach first.'));
});

test('an administrator and the provider share one live map, and a stale save is refused', async () => {
  reset();
  const providerSave = await call(PROVIDER, 'PUT', { part: 'map' }, { baseUpdatedAt: '', map: blankMap(PROVIDER, 'Pat') });
  assert.strictEqual(providerSave.status, 200);
  const opened = await call(ADMIN, 'GET', { part: 'map', provider: PROVIDER });
  assert.strictEqual(opened.status, 200);
  const stale = await call(ADMIN, 'PUT', { part: 'map', provider: PROVIDER }, { baseUpdatedAt: '', map: opened.body.map });
  assert.strictEqual(stale.status, 409);
  opened.body.map.topics[1].p = 'Tried joining the play.';
  const fresh = await call(ADMIN, 'PUT', { part: 'map', provider: PROVIDER }, { baseUpdatedAt: opened.body.map.updatedAt, map: opened.body.map });
  assert.strictEqual(fresh.status, 200);
  const seen = await call(PROVIDER, 'GET', { part: 'map' });
  assert.strictEqual(seen.body.map.topics[1].p, 'Tried joining the play.');
});

test('the provider accepts a proposed time, and a failed calendar leaves it ready to retry', async () => {
  reset();
  const proposed = await call(ADMIN, 'POST', { part: 'appointments', provider: PROVIDER }, {
    action: 'propose', provider: PROVIDER, start: '2026-12-02T15:00', end: '2026-12-02T16:00',
  });
  assert.strictEqual(proposed.status, 200, JSON.stringify(proposed.body));
  assert.strictEqual(proposed.body.appointment.status, 'proposed');
  const accepted = await call(PROVIDER, 'POST', { part: 'appointments' }, { action: 'accept', id: proposed.body.appointment.id });
  assert.strictEqual(accepted.status, 200, JSON.stringify(accepted.body));
  assert.strictEqual(accepted.body.appointment.status, 'accepted');
  assert.ok(accepted.body.calendarWarning);
  const other = await call('sam@emmeco.org', 'POST', { part: 'appointments', provider: PROVIDER }, { action: 'decline', id: proposed.body.appointment.id });
  assert.strictEqual(other.status, 403);
});

test('a provider can suggest another time and an administrator can confirm it onto the calendars', async () => {
  reset();
  const originalFetch = global.fetch;
  global.fetch = async function (url, options) {
    const text = String(url);
    if (text.includes('/events')) {
      const who = text.includes('kelly.fitzpatrick') ? 'admin' : 'provider';
      return new Response(JSON.stringify({ id: 'evt-' + who }), { status: 201, headers: { 'content-type': 'application/json' } });
    }
    if (text.includes('Kelly Fitzpatrick') || text.includes('Kelly%20Fitzpatrick')) {
      return new Response(JSON.stringify({ value: [{ displayName: 'Kelly Fitzpatrick', mail: 'kelly.fitzpatrick@emmeco.org', userPrincipalName: 'kelly.fitzpatrick@emmeco.org' }] }), { status: 200, headers: { 'content-type': 'application/json' } });
    }
    if (text.includes('/sendMail')) return new Response('', { status: 202 });
    return originalFetch(url, options);
  };
  try {
    const proposed = await call(ADMIN, 'POST', { part: 'appointments', provider: PROVIDER }, {
      action: 'propose', provider: PROVIDER, start: '2026-12-03T09:00', end: '2026-12-03T10:00',
    });
    const suggested = await call(PROVIDER, 'POST', { part: 'appointments' }, {
      action: 'suggest', id: proposed.body.appointment.id, start: '2026-12-04T11:00', end: '2026-12-04T12:00',
    });
    assert.strictEqual(suggested.body.appointment.status, 'suggested');
    const confirmed = await call(ADMIN, 'POST', { part: 'appointments', provider: PROVIDER }, { action: 'confirm', id: proposed.body.appointment.id });
    assert.strictEqual(confirmed.status, 200, JSON.stringify(confirmed.body));
    assert.strictEqual(confirmed.body.appointment.status, 'confirmed');
    assert.strictEqual(confirmed.body.appointment.start, '2026-12-04T11:00');
    assert.ok(confirmed.body.appointment.adminEventId);
    assert.ok(confirmed.body.appointment.providerEventId);
    assert.strictEqual(confirmed.body.appointment.adminEmail, 'kelly.fitzpatrick@emmeco.org');
    assert.ok(!JSON.stringify(confirmed.body).includes('I noticed'));
  } finally {
    global.fetch = originalFetch;
  }
});

test('the read-only copy is uploaded into the provider folder under RS Providers', async () => {
  reset();
  await call(PROVIDER, 'PUT', { part: 'map' }, { baseUpdatedAt: '', map: blankMap(PROVIDER, 'Pat Provider') });
  const originalFetch = global.fetch;
  let uploaded = '';
  global.fetch = async function (url, options) {
    const text = decodeURIComponent(String(url));
    const method = (options && options.method) || 'GET';
    if (text.includes('/shares/')) {
      return new Response(JSON.stringify({ id: 'folder-root', name: 'RS Providers', parentReference: { driveId: 'drive1' } }), { status: 200, headers: { 'content-type': 'application/json' } });
    }
    if (text.includes('/children')) {
      return new Response(JSON.stringify({ value: [{ id: 'folder-pat', name: 'Natalie Delgado', folder: {} }] }), { status: 200, headers: { 'content-type': 'application/json' } });
    }
    if (text.includes('/content') && method === 'PUT') {
      uploaded = text;
      return new Response(JSON.stringify({ webUrl: 'https://example.test/map.pdf' }), { status: 200, headers: { 'content-type': 'application/json' } });
    }
    return originalFetch(url, options);
  };
  try {
    const copy = await call(PROVIDER, 'POST', { part: 'copy' });
    assert.strictEqual(copy.status, 200, JSON.stringify(copy.body));
    assert.strictEqual(copy.body.copy.ok, true);
    assert.ok(copy.body.copy.path.includes('Natalie Delgado'));
    assert.ok(uploaded.includes('folder-pat'));
    assert.ok(uploaded.includes('Reflective Supervision.pdf'));
  } finally {
    global.fetch = originalFetch;
    resetReflectiveOfficeCache();
  }
});

test('the observation folder stays admin-only and still refuses a change without a part', async () => {
  reset();
  assert.strictEqual((await call(null, 'GET', {})).status, 401);
  assert.strictEqual((await call(PROVIDER, 'GET', {})).status, 403);
  assert.strictEqual((await call(ADMIN, 'POST', {})).status, 405);
});

test('email reminders go to the other providers, not to Kelly and not outside Emme', async () => {
  reset();
  const originalFetch = global.fetch;
  const sent = [];
  global.fetch = async function (url, options) {
    const text = decodeURIComponent(String(url));
    if (text.includes('Kelly Fitzpatrick')) {
      return new Response(JSON.stringify({ value: [{ displayName: 'Kelly Fitzpatrick', mail: 'kelly.fitzpatrick@emmeco.org', userPrincipalName: 'kelly.fitzpatrick@emmeco.org' }] }), { status: 200, headers: { 'content-type': 'application/json' } });
    }
    if (text.includes('/sendMail')) {
      sent.push(JSON.parse(options.body));
      return new Response('', { status: 202 });
    }
    return originalFetch(url, options);
  };
  try {
    const secret = 'private map sentence that must stay off the email';
    const map = blankMap(PROVIDER, 'Pat Provider');
    map.topics[0].e = secret;
    await call(PROVIDER, 'PUT', { part: 'map' }, { baseUpdatedAt: '', map });
    await call(ADMIN, 'PUT', { part: 'map', provider: 'kelly.fitzpatrick@emmeco.org' }, { baseUpdatedAt: '', map: blankMap('kelly.fitzpatrick@emmeco.org', 'Kelly Fitzpatrick') });
    const first = await call(null, 'POST', { part: 'remind' }, {}, { 'x-emme-automation-key': 'test-automation-key' });
    assert.strictEqual(first.status, 200, JSON.stringify(first.body));
    const addresses = sent.map((item) => item.message.toRecipients[0].emailAddress.address);
    assert.ok(addresses.includes(PROVIDER));
    assert.ok(addresses.includes('kelly.fitzpatrick@emmeco.org'));
    assert.ok(addresses.every((address) => address.endsWith('@emmeco.org')));
    assert.ok(!JSON.stringify(sent).includes(secret));
    assert.ok(!first.body.reminded.some((item) => item.email === 'kelly.fitzpatrick@emmeco.org'));
    const again = await call(null, 'POST', { part: 'remind' }, {}, { 'x-emme-automation-key': 'test-automation-key' });
    assert.strictEqual(again.body.reminded.length, 0);
    assert.strictEqual((await call(null, 'POST', { part: 'remind' }, {})).status, 401);
  } finally {
    global.fetch = originalFetch;
    resetReflectiveOfficeCache();
  }
});
