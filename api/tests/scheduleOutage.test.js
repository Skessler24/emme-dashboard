const test = require('node:test');
const assert = require('node:assert');

// A provider's schedule must save and load even when Microsoft 365 is down or very slow.
require('./fakes');
process.env.GRAPH_TIMEOUT_MS = '150';
delete process.env.AZURE_CLIENT_SECRET;

const handler = require('../schedule/index.js');
const { resetGraphBreaker } = require('../dist/graphClient');

const realFetch = global.fetch;

function principal(email) {
  return Buffer.from(JSON.stringify({ userId: 'u-' + email, userDetails: email, claims: [] })).toString('base64');
}
async function call(email, method, rest, body) {
  const context = {};
  await handler(context, { method, body, headers: { 'x-ms-client-principal': principal(email) }, params: { rest: rest || '' } });
  return context.res;
}
function workspace(note) {
  return { version: 1, boards: [{ id: 'b1', name: 'Current Schedule', days: [{ id: 'd1', name: 'Monday', color: 'blue', slots: [{ id: 's1', time: '9:00 AM', status: 'Full', notes: note, lastDay: '' }] }] }], clients: [], dxCodes: [] };
}
function slotNote(res) {
  const board = res.body.workspace.boards.find((b) => b.name === 'Current Schedule');
  return board.days[0].slots[0].notes;
}

test('saving and loading work when Microsoft cannot be reached at all', async () => {
  resetGraphBreaker();
  let calls = 0;
  global.fetch = async () => { calls += 1; throw new TypeError('fetch failed'); };
  try {
    const email = 'down1@emmeco.org';
    const first = await call(email, 'PUT', '', { workspace: workspace('first save') });
    assert.strictEqual(first.status, 200, JSON.stringify(first.body));
    assert.strictEqual(first.body.sharePoint.ok, false);
    assert.strictEqual(slotNote(first), 'first save');

    const second = await call(email, 'PUT', '', { workspace: workspace('second save'), base: first.body.workspace });
    assert.strictEqual(second.status, 200, JSON.stringify(second.body));
    assert.strictEqual(slotNote(second), 'second save');

    const loaded = await call(email, 'GET', '');
    assert.strictEqual(loaded.status, 200, JSON.stringify(loaded.body));
    assert.strictEqual(slotNote(loaded), 'second save');

    const meta = await call(email, 'GET', 'meta');
    assert.strictEqual(meta.status, 200);
    assert.strictEqual(meta.body.exists, true);
  } finally {
    global.fetch = realFetch;
    resetGraphBreaker();
  }
});

test('after Microsoft fails to connect, later saves do not keep waiting on it', async () => {
  resetGraphBreaker();
  let calls = 0;
  global.fetch = async () => { calls += 1; throw new TypeError('fetch failed'); };
  try {
    const email = 'down2@emmeco.org';
    const first = await call(email, 'PUT', '', { workspace: workspace('one') });
    assert.strictEqual(first.status, 200);
    const callsAfterFirst = calls;
    assert.ok(callsAfterFirst >= 1);
    const second = await call(email, 'PUT', '', { workspace: workspace('two'), base: first.body.workspace });
    assert.strictEqual(second.status, 200);
    assert.strictEqual(calls, callsAfterFirst, 'no further calls while Microsoft is paused');
    assert.strictEqual(slotNote(second), 'two');
  } finally {
    global.fetch = realFetch;
    resetGraphBreaker();
  }
});

test('a Microsoft call that never answers is cut off and the schedule still saves', async () => {
  resetGraphBreaker();
  global.fetch = (url, options) => new Promise((resolve, reject) => {
    const signal = options && options.signal;
    if (signal) signal.addEventListener('abort', () => reject(new DOMException('timed out', 'TimeoutError')));
  });
  try {
    const email = 'slow1@emmeco.org';
    const started = Date.now();
    const saved = await call(email, 'PUT', '', { workspace: workspace('saved while slow') });
    assert.strictEqual(saved.status, 200, JSON.stringify(saved.body));
    assert.strictEqual(slotNote(saved), 'saved while slow');
    assert.ok(Date.now() - started < 5000, 'finished in ' + (Date.now() - started) + ' ms');
    resetGraphBreaker();
    const loaded = await call(email, 'GET', '');
    assert.strictEqual(loaded.status, 200);
    assert.strictEqual(slotNote(loaded), 'saved while slow');
  } finally {
    global.fetch = realFetch;
    resetGraphBreaker();
  }
});
