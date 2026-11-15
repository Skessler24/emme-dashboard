const test = require('node:test');
const assert = require('node:assert');
const { blobs } = require('./fakes');
process.env.EMME_ADMIN_EMAILS = 'skessler@emmeco.org';
process.env.AUTOMATION_API_KEY = 'test-automation-key';

const intake = require('../intake/index.js');
const { TOPICS, getTopic } = require('../dist/intakeTopics');
const { buildPrompt, missingItems, assessAll } = require('../dist/intakePrompt');

const BILLING = getTopic('billing');
const GENERAL_QUESTIONS = BILLING.questions;
const assessSource = (item) => {
  const a = BILLING.assess(item);
  const row = (key) => a.rows.find((r) => r.key === key);
  return Object.assign({}, a, { sending: row('sending'), checking: row('checking') });
};

const ADMIN = 'skessler@emmeco.org';

function principal(email) {
  return Buffer.from(JSON.stringify({ userId: 'u-' + email, userDetails: email, claims: [{ typ: 'preferred_username', val: email }] })).toString('base64');
}
async function callTopic(topic, email, method, rest, body, extraHeaders) {
  const context = {};
  const headers = Object.assign(email ? { 'x-ms-client-principal': principal(email) } : {}, extraHeaders || {});
  await intake(context, { method, body, headers, params: { topic, rest: rest === undefined ? 'state' : rest } });
  return context.res;
}
// Billing is the default topic, and the state address is spelled out as /state.
function call(email, method, rest, body, extraHeaders) {
  return callTopic('billing', email, method, rest === '' ? 'state' : rest, body, extraHeaders);
}
function answer(text, choice) { return { text: text || '', choice: choice || [], unsure: false }; }
function reset() { for (const key of [...blobs.keys()]) if (key.startsWith('billing-intake/') || key.startsWith('intake/')) blobs.delete(key); }
function source(id, kind, name, answers) { return { id, kind, name, answers: answers || {} }; }
function b64(text) { return Buffer.from(text).toString('base64'); }

test('only a signed-in Emme Admin can open Billing setup, and the automation key is not enough', async () => {
  reset();
  assert.strictEqual((await call(null, 'GET', 'state')).status, 401);
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
  assert.strictEqual(res.body.assessments[0].rows.find((r) => r.key === 'sending').level, 'by-hand');
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
  const text = buildPrompt(BILLING, empty, [], new Date('2026-11-14T12:00:00Z'));
  assert.match(text, /No funding sources have been added yet/);
  assert.strictEqual(missingItems(BILLING, empty).length, GENERAL_QUESTIONS.length);
});

test('the questions are written in plain words with no abbreviations, in every topic', () => {
  TOPICS.forEach((topic) => {
    const everything = [];
    const grab = (value) => { if (value) everything.push(String(value)); };
    const addQuestions = (list) => list.forEach((q) => { grab(q.text); grab(q.help); (q.options || []).forEach((o) => grab(o.label)); });
    grab(topic.title);
    topic.goal.forEach(grab);
    addQuestions(topic.questions);
    topic.sections.forEach((s) => { grab(s.title); grab(s.intro); });
    topic.fileKinds.forEach((k) => grab(k.label));
    topic.guide.forEach((c) => { grab(c.title); grab(c.what); grab(c.works); grab(c.blocked); grab(c.needs); });
    if (topic.group) {
      addQuestions(topic.group.questions);
      grab(topic.group.intro);
      topic.group.sections.forEach((s) => { grab(s.title); grab(s.intro); });
      topic.group.kinds.forEach((k) => { grab(k.label); grab(k.description); });
    }
    const offenders = everything.filter((text) => /\b[A-Z]{2,}\b/.test(text.replace(/\bIntakeQ\b/g, '')));
    assert.deepStrictEqual(offenders, [], topic.id);
  });
});

test('every topic is complete: unique ids, valid sections, and the questions the build prompt needs', () => {
  const prefixes = new Set();
  const topicIds = new Set();
  TOPICS.forEach((topic) => {
    assert.ok(!topicIds.has(topic.id), 'duplicate topic ' + topic.id);
    topicIds.add(topic.id);
    assert.ok(/^[a-z-]+$/.test(topic.id), topic.id);
    assert.ok(topic.storagePrefix.endsWith('/'), topic.id);
    assert.ok(!prefixes.has(topic.storagePrefix), 'shared storage prefix ' + topic.storagePrefix);
    prefixes.add(topic.storagePrefix);
    const ids = new Set();
    const check = (questions, sections, kinds, label) => questions.forEach((q) => {
      assert.ok(!ids.has(q.id), topic.id + ' duplicate ' + q.id);
      ids.add(q.id);
      if (q.kind === 'text') assert.ok(!q.options, q.id);
      else assert.ok(q.options && q.options.length >= 2, q.id);
      assert.ok(sections.some((s) => s.id === q.section), label + ' section for ' + q.id);
      (q.kinds || []).forEach((kind) => assert.ok(kinds.some((k) => k.id === kind), q.id));
    });
    check(topic.questions, topic.sections, [], topic.id);
    if (topic.group) check(topic.group.questions, topic.group.sections, topic.group.kinds, topic.id + ' group');
    ['g_success', 'g_deadline'].forEach((id) => assert.ok(topic.questions.some((q) => q.id === id), topic.id + ' needs ' + id));
    if (topic.firstId) assert.ok(topic.questions.some((q) => q.id === topic.firstId), topic.id + ' firstId');
    if (topic.group) assert.ok(topic.assess, topic.id + ' has a group but no assessment');
    assert.ok(topic.fileKinds.length >= 3 && topic.fileKinds.some((k) => k.id === 'other'), topic.id);
    assert.ok(topic.guide.length >= 3, topic.id);
    topic.sections.forEach((s) => assert.ok(topic.questions.some((q) => q.section === s.id), topic.id + ' empty section ' + s.id));
  });
});

test('every topic works end to end and keeps its answers and files apart from every other topic', async () => {
  reset();
  for (const topic of TOPICS) {
    const open = await callTopic(topic.id, ADMIN, 'GET');
    assert.strictEqual(open.status, 200, topic.id);
    assert.strictEqual(open.body.topic.id, topic.id);
    assert.strictEqual(open.body.generalQuestions.length, topic.questions.length);
    assert.strictEqual(!!open.body.topic.group, !!topic.group);
    const saved = await callTopic(topic.id, ADMIN, 'PUT', 'state', {
      baseUpdatedAt: '',
      general: { g_success: answer('Marker for ' + topic.id) },
      sources: topic.group ? [source('one-' + topic.id.slice(0, 5), topic.group.kinds[0].id, 'Example')] : [source('sneaky-one', 'ei', 'Should be dropped')],
    });
    assert.strictEqual(saved.status, 200, topic.id);
    const up = await callTopic(topic.id, ADMIN, 'POST', 'files', { scope: 'general', kind: 'other', name: topic.id + '.txt', contentBase64: b64('file for ' + topic.id), confirmedNoClientInfo: true });
    assert.strictEqual(up.status, 200, topic.id);
  }
  for (const topic of TOPICS) {
    const state = (await callTopic(topic.id, ADMIN, 'GET')).body;
    assert.strictEqual(state.answers.general.g_success.text, 'Marker for ' + topic.id);
    assert.deepStrictEqual(state.files.map((f) => f.name), [topic.id + '.txt']);
    assert.strictEqual(state.answers.sources.length, topic.group ? 1 : 0, topic.id);
    const prompt = (await callTopic(topic.id, ADMIN, 'GET', 'prompt')).body;
    assert.match(prompt.prompt, new RegExp('Build request: ' + topic.title));
    assert.match(prompt.prompt, new RegExp('Marker for ' + topic.id));
    assert.match(prompt.prompt, /at most five steps/);
    assert.match(prompt.prompt, /Do not merge it/);
    assert.strictEqual(prompt.assessments.length, topic.assess ? 1 : 0, topic.id);
    // A file from one topic cannot be downloaded through another.
    const other = TOPICS.find((t) => t.id !== topic.id);
    assert.strictEqual((await callTopic(other.id, ADMIN, 'GET', 'files/' + state.files[0].id)).status, 404);
  }
});

test('a prompt for a topic with no repeating group still lists every open question and the connection options', async () => {
  reset();
  const referrals = getTopic('referrals');
  const staff = getTopic('staff');
  const empty = { updatedAt: '', updatedBy: '', general: {}, sources: [] };
  const text = buildPrompt(staff, empty, [], new Date('2026-11-14T12:00:00Z'));
  assert.match(text, /Build request: Staff Information/);
  assert.strictEqual(missingItems(staff, empty).length, staff.questions.length);
  assert.ok(!/invoice|re-bill/i.test(text), 'no billing wording in other topics');
  assert.match(text, /Feasibility: the connection options/);
  assert.match(buildPrompt(referrals, empty, [], new Date()), /No referral sources have been added yet/);
});

test('referral sources are assessed from how referrals arrive and how Emme answers', () => {
  const referrals = getTopic('referrals');
  const a = (choice) => ({ choice, text: '', unsure: false });
  const run = (answers) => referrals.assess({ id: 'x-y-z', kind: 'clinic', name: 'Clinic', answers });
  const row = (result, key) => result.rows.find((r) => r.key === key).level;

  const form = run({ r_arrive: a(['form']), r_reply: a(['email']) });
  assert.strictEqual(row(form, 'arrive'), 'automatic');
  assert.strictEqual(row(form, 'reply'), 'partly');

  const fax = run({ r_arrive: a(['fax', 'phone']), r_reply: a(['phone']) });
  assert.strictEqual(row(fax, 'arrive'), 'by-hand');
  assert.ok(fax.cautions.some((text) => /stay by hand/.test(text)));

  const blocked = run({ r_arrive: a(['portal']), r_automation_allowed: a(['no']), r_second_step: a(['text-code']) });
  assert.strictEqual(blocked.blockers.length, 2);

  assert.strictEqual(row(run({}), 'arrive'), 'unknown');
});

test('unknown topics, plain access rules, and tampering with the topic name are refused', async () => {
  reset();
  assert.strictEqual((await callTopic('nonsense', ADMIN, 'GET')).status, 404);
  assert.strictEqual((await callTopic('constructor', ADMIN, 'GET')).status, 404);
  assert.strictEqual((await callTopic('__proto__', ADMIN, 'GET')).status, 404);
  assert.strictEqual((await callTopic('', ADMIN, 'GET')).status, 404);
  assert.strictEqual((await callTopic('onboarding', null, 'GET')).status, 401);
  assert.strictEqual((await callTopic('onboarding', 'ndelgado@emmeco.org', 'GET')).status, 403);
  assert.strictEqual((await callTopic('onboarding', null, 'GET', 'state', null, { 'x-emme-automation-key': 'test-automation-key' })).status, 401);
  assert.strictEqual((await callTopic('onboarding', ADMIN, 'GET', 'nope')).status, 404);
  // Leaving the path empty still means the saved state.
  assert.strictEqual((await callTopic('onboarding', ADMIN, 'GET', '')).status, 200);
});

test('billing keeps reading answers saved before the other topics existed', async () => {
  reset();
  const old = { updatedAt: '2026-11-01T00:00:00.000Z', updatedBy: ADMIN, general: { g_hours: { text: 'saved earlier', choice: [], unsure: false } }, sources: [] };
  blobs.set('billing-intake/answers.json', { data: Buffer.from(JSON.stringify(old)) });
  const state = (await call(ADMIN, 'GET')).body;
  assert.strictEqual(state.answers.general.g_hours.text, 'saved earlier');
  assert.strictEqual((await callTopic('onboarding', ADMIN, 'GET')).body.answers.updatedAt, '');
});