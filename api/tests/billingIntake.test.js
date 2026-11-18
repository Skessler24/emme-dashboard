const test = require('node:test');
const assert = require('node:assert');
const { blobs } = require('./fakes');
process.env.EMME_ADMIN_EMAILS = 'skessler@emmeco.org';
process.env.AUTOMATION_API_KEY = 'test-automation-key';

const intake = require('../billing-intake/index.js');
const { GENERAL_QUESTIONS, SOURCE_QUESTIONS, SOURCE_KINDS, CONNECTION_GUIDE, FILE_KINDS, GENERAL_SECTIONS, SOURCE_SECTIONS } = require('../dist/billingQuestions');
const { assessSource, buildPrompt, missingItems } = require('../dist/billingPrompt');

const ADMIN = 'skessler@emmeco.org';

function principal(email) {
  return Buffer.from(JSON.stringify({ userId: 'u-' + email, userDetails: email, claims: [{ typ: 'preferred_username', val: email }] })).toString('base64');
}
async function call(email, method, rest, body, extraHeaders) {
  const context = {};
  const headers = Object.assign(email ? { 'x-ms-client-principal': principal(email) } : {}, extraHeaders || {});
  await intake(context, { method, body, headers, params: { rest: rest || '' } });
  return context.res;
}
function answer(text, choice) { return { text: text || '', choice: choice || [], unsure: false }; }
function reset() { for (const key of [...blobs.keys()]) if (key.startsWith('billing-intake/')) blobs.delete(key); }
function source(id, kind, name, answers) { return { id, kind, name, answers: answers || {} }; }
function b64(text) { return Buffer.from(text).toString('base64'); }

test('only a signed-in Emme Admin can open Billing setup, and the automation key is not enough', async () => {
  reset();
  assert.strictEqual((await call(null, 'GET')).status, 401);
  assert.strictEqual((await call('teacher@cherrycreekschools.org', 'GET')).status, 403);
  assert.strictEqual((await call('ndelgado@emmeco.org', 'GET')).status, 403);
  assert.strictEqual((await call(null, 'GET', '', null, { 'x-emme-automation-key': 'test-automation-key' })).status, 401);
  const ok = await call(ADMIN, 'GET');
  assert.strictEqual(ok.status, 200);
  assert.ok(ok.body.generalQuestions.length > 10);
  assert.deepStrictEqual(ok.body.answers.sources, []);
});

test('answers save and come back, and unknown or oversized input is cleaned', async () => {
  reset();
  const saved = await call(ADMIN, 'PUT', '', {
    baseUpdatedAt: '',
    general: {
      g_hours: answer('x'.repeat(9000)),
      g_frequency: answer('', ['daily', 'not-a-real-option']),
      g_alert_how: answer('', ['email', 'email', 'teams']),
      not_a_question: answer('sneaky'),
    },
    sources: [
      source('ccb-one', 'ei', 'Example Board', { ei_board: answer('Example region'), mc_program: answer('wrong kind, dropped') }),
      source('Bad Id!', 'unknown-kind', 'Dropped'),
    ],
  });
  assert.strictEqual(saved.status, 200);
  const state = (await call(ADMIN, 'GET')).body;
  assert.strictEqual(state.answers.general.g_hours.text.length, 6000);
  assert.deepStrictEqual(state.answers.general.g_frequency.choice, ['daily']);
  assert.deepStrictEqual(state.answers.general.g_alert_how.choice, ['email', 'teams']);
  assert.strictEqual(state.answers.general.not_a_question, undefined);
  assert.strictEqual(state.answers.sources.length, 1);
  assert.strictEqual(state.answers.sources[0].answers.ei_board.text, 'Example region');
  assert.strictEqual(state.answers.sources[0].answers.mc_program, undefined);
  assert.strictEqual(state.answers.updatedBy, ADMIN);
});

test('two people cannot overwrite each other without noticing', async () => {
  reset();
  const first = await call(ADMIN, 'PUT', '', { baseUpdatedAt: '', general: { g_hours: answer('ten') }, sources: [] });
  assert.strictEqual(first.status, 200);
  const stale = await call(ADMIN, 'PUT', '', { baseUpdatedAt: '', general: { g_hours: answer('twenty') }, sources: [] });
  assert.strictEqual(stale.status, 409);
  assert.match(stale.body.error, /Reload/);
  const fresh = await call(ADMIN, 'PUT', '', { baseUpdatedAt: first.body.updatedAt, general: { g_hours: answer('twenty') }, sources: [] });
  assert.strictEqual(fresh.status, 200);
  assert.strictEqual((await call(ADMIN, 'GET')).body.answers.general.g_hours.text, 'twenty');
});

test('uploads need a confirmation, an allowed type, a real size, and a saved funding source', async () => {
  reset();
  await call(ADMIN, 'PUT', '', { baseUpdatedAt: '', general: {}, sources: [source('ccb-one', 'ei', 'Example Board')] });
  const good = { scope: 'general', kind: 'spreadsheet', name: 'Invoice check.xlsx', contentBase64: b64('fake spreadsheet bytes'), confirmedNoClientInfo: true };
  assert.strictEqual((await call(ADMIN, 'POST', 'files', Object.assign({}, good, { confirmedNoClientInfo: false }))).status, 400);
  assert.strictEqual((await call(ADMIN, 'POST', 'files', Object.assign({}, good, { name: 'run.exe' }))).status, 400);
  assert.strictEqual((await call(ADMIN, 'POST', 'files', Object.assign({}, good, { kind: 'nonsense' }))).status, 400);
  assert.strictEqual((await call(ADMIN, 'POST', 'files', Object.assign({}, good, { contentBase64: '' }))).status, 400);
  assert.strictEqual((await call(ADMIN, 'POST', 'files', Object.assign({}, good, { contentBase64: 'not base64!!' }))).status, 400);
  const big = Buffer.alloc(6 * 1024 * 1024 + 10, 1).toString('base64');
  assert.strictEqual((await call(ADMIN, 'POST', 'files', Object.assign({}, good, { contentBase64: big }))).status, 400);
  const missingSource = await call(ADMIN, 'POST', 'files', Object.assign({}, good, { scope: 'never-saved' }));
  assert.strictEqual(missingSource.status, 400);
  assert.match(missingSource.body.error, /Save the funding source first/);
  assert.strictEqual((await call(ADMIN, 'POST', 'files', Object.assign({}, good, { scope: 'ccb-one' }))).status, 200);
  assert.strictEqual((await call('ndelgado@emmeco.org', 'POST', 'files', good)).status, 403);
});

test('an uploaded file downloads unchanged, always as a download, and can be removed', async () => {
  reset();
  const bytes = Buffer.from([0, 1, 2, 250, 251, 252, 80, 75]);
  const up = await call(ADMIN, 'POST', 'files', { scope: 'general', kind: 'blank-bill', note: 'Blank form', name: '../..\\Bill <form>.xlsx', contentBase64: bytes.toString('base64'), confirmedNoClientInfo: true });
  assert.strictEqual(up.status, 200);
  assert.strictEqual(up.body.file.name, 'Bill _form_.xlsx');
  assert.strictEqual(up.body.file.uploadedBy, ADMIN);
  const down = await call(ADMIN, 'GET', 'files/' + up.body.file.id);
  assert.strictEqual(down.status, 200);
  assert.ok(Buffer.compare(down.body, bytes) === 0);
  assert.strictEqual(down.headers['Content-Type'], 'application/octet-stream');
  assert.match(down.headers['Content-Disposition'], /^attachment;/);
  assert.strictEqual(down.headers['X-Content-Type-Options'], 'nosniff');
  assert.strictEqual((await call('ndelgado@emmeco.org', 'GET', 'files/' + up.body.file.id)).status, 403);
  assert.strictEqual((await call(ADMIN, 'DELETE', 'files/' + up.body.file.id)).status, 200);
  assert.strictEqual((await call(ADMIN, 'GET', 'files/' + up.body.file.id)).status, 404);
  assert.strictEqual((await call(ADMIN, 'GET')).body.files.length, 0);
});

test('removing a funding source also removes its files', async () => {
  reset();
  const saved = await call(ADMIN, 'PUT', '', { baseUpdatedAt: '', general: {}, sources: [source('ccb-one', 'ei', 'Board'), source('plan-two', 'private', 'Insurer')] });
  await call(ADMIN, 'POST', 'files', { scope: 'ccb-one', kind: 'blank-bill', name: 'a.csv', contentBase64: b64('a,b'), confirmedNoClientInfo: true });
  await call(ADMIN, 'POST', 'files', { scope: 'plan-two', kind: 'blank-bill', name: 'b.csv', contentBase64: b64('c,d'), confirmedNoClientInfo: true });
  await call(ADMIN, 'POST', 'files', { scope: 'general', kind: 'other', name: 'c.txt', contentBase64: b64('general'), confirmedNoClientInfo: true });
  await call(ADMIN, 'PUT', '', { baseUpdatedAt: saved.body.updatedAt, general: {}, sources: [source('plan-two', 'private', 'Insurer')] });
  const files = (await call(ADMIN, 'GET')).body.files;
  assert.deepStrictEqual(files.map((f) => f.name).sort(), ['b.csv', 'c.txt']);
  assert.strictEqual([...blobs.keys()].filter((key) => key.startsWith('billing-intake/files/')).length, 2);
});

test('the build prompt carries the answers, the files, and what is still unknown', async () => {
  reset();
  await call(ADMIN, 'PUT', '', {
    baseUpdatedAt: '',
    general: { g_hours: answer('About twelve hours a week'), g_first_source: answer('Medicaid, because it is the largest'), g_frequency: { text: '', choice: ['not-sure'], unsure: true } },
    sources: [source('ccb-one', 'ei', 'Example Board', { s_send_how: answer('Typed in by Pat', ['portal-typed']) })],
  });
  await call(ADMIN, 'POST', 'files', { scope: 'ccb-one', kind: 'blank-bill', note: 'Blank invoice', name: 'Board invoice.xlsx', contentBase64: b64('x'), confirmedNoClientInfo: true });
  const res = await call(ADMIN, 'GET', 'prompt');
  assert.strictEqual(res.status, 200);
  const text = res.body.prompt;
  assert.match(text, /at most five steps/);
  assert.match(text, /Do not merge it/);
  assert.match(text, /About twelve hours a week/);
  assert.match(text, /Medicaid, because it is the largest/);
  assert.match(text, /Example Board \(Early Intervention Community Centered Board\)/);
  assert.match(text, /`Board invoice.xlsx`/);
  assert.match(text, /Typed in by Pat/);
  assert.match(text, /Marked "not sure yet"/);
  assert.match(text, /Open questions to ask me before building/);
  assert.ok(res.body.missing.length > 10);
  assert.strictEqual(res.body.assessments[0].sending.level, 'by-hand');
  assert.ok(!/password\s*[:=]/i.test(text));
});

test('the assessment says what can be automatic and what is blocked', () => {
  const s = (answers) => ({ id: 'a-b-c', kind: 'medicaid', name: 'Test', answers });
  const a = (choice) => ({ choice, text: '', unsure: false });

  const clearing = assessSource(s({ s_send_how: a(['clearinghouse']), s_pay_how: a(['remittance-file']) }));
  assert.strictEqual(clearing.sending.level, 'automatic');
  assert.strictEqual(clearing.checking.level, 'automatic');
  assert.deepStrictEqual(clearing.blockers, []);

  const website = assessSource(s({ s_send_how: a(['portal-typed']), s_pay_how: a(['portal-report']), s_second_step: a(['text-code']), s_official_connection: a(['no']) }));
  assert.strictEqual(website.sending.level, 'by-hand');
  assert.strictEqual(website.checking.level, 'partly');
  assert.ok(website.blockers.some((text) => /second step/.test(text)));
  assert.ok(website.cautions.some((text) => /only way to see payments/.test(text)));

  const forbidden = assessSource(s({ s_send_how: a(['portal-upload']), s_pay_how: a(['portal-report']), s_automation_allowed: a(['no']) }));
  assert.ok(forbidden.blockers.some((text) => /do not allow/.test(text)));

  const official = assessSource(s({ s_send_how: a(['fax']), s_pay_how: a(['phone']), s_official_connection: a(['connection']), s_second_step: a(['text-code']) }));
  assert.strictEqual(official.sending.level, 'automatic');
  assert.deepStrictEqual(official.blockers, []);

  const paper = assessSource(s({ s_send_how: a(['mail', 'fax']), s_pay_how: a(['paper-remittance', 'phone']) }));
  assert.strictEqual(paper.sending.level, 'by-hand');
  assert.strictEqual(paper.checking.level, 'by-hand');

  const blank = assessSource(s({}));
  assert.strictEqual(blank.sending.level, 'unknown');
  assert.ok(blank.toConfirm.length >= 3);
});

test('an empty setup still produces a prompt that asks everything', () => {
  const empty = { updatedAt: '', updatedBy: '', general: {}, sources: [] };
  const text = buildPrompt(empty, [], new Date('2026-11-14T12:00:00Z'));
  assert.match(text, /No funding sources have been added yet/);
  assert.strictEqual(missingItems(empty).length, GENERAL_QUESTIONS.length);
});

test('the questions are written in plain words with no abbreviations', () => {
  const everything = [];
  const grab = (value) => { if (value) everything.push(String(value)); };
  [...GENERAL_QUESTIONS, ...SOURCE_QUESTIONS].forEach((q) => { grab(q.text); grab(q.help); (q.options || []).forEach((o) => grab(o.label)); });
  [...GENERAL_SECTIONS, ...SOURCE_SECTIONS].forEach((s) => { grab(s.title); grab(s.intro); });
  SOURCE_KINDS.forEach((k) => { grab(k.label); grab(k.description); });
  FILE_KINDS.forEach((k) => grab(k.label));
  CONNECTION_GUIDE.forEach((c) => { grab(c.title); grab(c.what); grab(c.works); grab(c.blocked); grab(c.needs); });
  const offenders = everything.filter((text) => /\b[A-Z]{2,}\b/.test(text.replace(/\bIntakeQ\b/g, '')));
  assert.deepStrictEqual(offenders, []);
});

test('every question can be saved and shown', () => {
  const ids = new Set();
  [...GENERAL_QUESTIONS, ...SOURCE_QUESTIONS].forEach((q) => {
    assert.ok(!ids.has(q.id), 'duplicate ' + q.id);
    ids.add(q.id);
    if (q.kind === 'text') assert.ok(!q.options, q.id);
    else assert.ok(q.options && q.options.length >= 2, q.id);
    (q.kinds || []).forEach((kind) => assert.ok(SOURCE_KINDS.some((k) => k.id === kind), q.id));
  });
  GENERAL_QUESTIONS.forEach((q) => assert.ok(GENERAL_SECTIONS.some((s) => s.id === q.section), q.id));
  SOURCE_QUESTIONS.forEach((q) => assert.ok(SOURCE_SECTIONS.some((s) => s.id === q.section), q.id));
});
