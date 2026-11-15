const test = require('node:test');
const assert = require('node:assert');
const { blobs } = require('./fakes');
process.env.EMME_ADMIN_EMAILS = 'skessler@emmeco.org';
process.env.AUTOMATION_API_KEY = 'test-automation-key';

const setup = require('../tile-setup/index.js');
const { TILE_SETUPS } = require('../dist/tileSetupQuestions');
const { buildTilePrompt, missingItems, seededAnswers } = require('../dist/tilePrompt');

const ADMIN = 'skessler@emmeco.org';
const IDS = ['onboarding', 'referrals', 'staff-information', 'staff-expectations', 'reflective-supervision'];

function principal(email) {
  return Buffer.from(JSON.stringify({ userId: 'u-' + email, userDetails: email, claims: [{ typ: 'preferred_username', val: email }] })).toString('base64');
}
async function call(email, method, rest, body, extraHeaders) {
  const context = {};
  const headers = Object.assign(email ? { 'x-ms-client-principal': principal(email) } : {}, extraHeaders || {});
  await setup(context, { method, body, headers, params: { rest: rest || '' } });
  return context.res;
}
function answer(text, choice, unsure) { return { text: text || '', choice: choice || [], unsure: unsure === true }; }
function reset() { for (const key of [...blobs.keys()]) if (key.startsWith('tile-setup/')) blobs.delete(key); }
function b64(text) { return Buffer.from(text).toString('base64'); }
function tile(id) { return TILE_SETUPS.find((item) => item.id === id); }

test('all five tiles are all there', () => {
  assert.deepStrictEqual(TILE_SETUPS.map((item) => item.id), IDS);
});

test('only a signed-in Emme Admin can open a tile setup, and the automation key is not enough', async () => {
  reset();
  assert.strictEqual((await call(null, 'GET', 'onboarding')).status, 401);
  assert.strictEqual((await call('teacher@cherrycreekschools.org', 'GET', 'onboarding')).status, 403);
  assert.strictEqual((await call('ndelgado@emmeco.org', 'GET', 'onboarding')).status, 403);
  assert.strictEqual((await call(null, 'GET', 'onboarding', null, { 'x-emme-automation-key': 'test-automation-key' })).status, 401);
  const ok = await call(ADMIN, 'GET', 'onboarding');
  assert.strictEqual(ok.status, 200);
  assert.strictEqual(ok.body.tile.id, 'onboarding');
  assert.ok(ok.body.questions.length > 10);
  assert.strictEqual((await call(ADMIN, 'GET', 'billing')).status, 404);
  assert.strictEqual((await call(ADMIN, 'GET', '')).body.tiles.length, 5);
});

test('before anyone saves, the answers start as a first draft from the documents, and the draft is not sent as part of the questions', async () => {
  reset();
  const state = (await call(ADMIN, 'GET', 'referrals')).body;
  assert.strictEqual(state.answers.firstDraft, true);
  assert.strictEqual(state.answers.updatedAt, '');
  assert.match(state.answers.answers.r_ei_days.text, /28 working day/);
  assert.match(state.answers.answers.r_ei_steps.text, /Provider notification/);
  assert.deepStrictEqual(state.answers.answers.r_assign.choice, ['yes']);
  assert.ok(state.questions.every((q) => q.seed === undefined));
  const onboarding = (await call(ADMIN, 'GET', 'onboarding')).body.answers.answers;
  assert.deepStrictEqual(onboarding.o_notice.choice, ['one-month']);
  const expectations = (await call(ADMIN, 'GET', 'staff-expectations')).body.answers.answers;
  assert.match(expectations.e_mile_rule.text, /5 miles/);
  assert.strictEqual(expectations.e_mile_rule.unsure, false);
});

test('a first draft answer the documents only hinted at is marked not sure', () => {
  const referrals = seededAnswers(tile('referrals'));
  assert.strictEqual(referrals.r_family_schedule.unsure, true);
  assert.strictEqual(referrals.r_private_steps.unsure, true);
  assert.strictEqual(seededAnswers(tile('onboarding')).o_upload.unsure, true);
});

test('saving replaces the first draft, and unknown or oversized input is cleaned', async () => {
  reset();
  const saved = await call(ADMIN, 'PUT', 'onboarding', {
    baseUpdatedAt: '',
    answers: {
      o_active: answer('x'.repeat(9000)),
      o_notice: answer('', ['two-months', 'not-a-real-option']),
      o_notify_how: answer('', ['email', 'email', 'teams']),
      r_assign: answer('belongs to another tile'),
      not_a_question: answer('sneaky'),
    },
  });
  assert.strictEqual(saved.status, 200);
  const state = (await call(ADMIN, 'GET', 'onboarding')).body;
  assert.strictEqual(state.answers.firstDraft, false);
  assert.strictEqual(state.answers.updatedBy, ADMIN);
  assert.strictEqual(state.answers.answers.o_active.text.length, 6000);
  assert.deepStrictEqual(state.answers.answers.o_notice.choice, ['two-months']);
  assert.deepStrictEqual(state.answers.answers.o_notify_how.choice, ['email', 'teams']);
  assert.strictEqual(state.answers.answers.r_assign, undefined);
  assert.strictEqual(state.answers.answers.not_a_question, undefined);
  assert.strictEqual(state.answers.answers.o_jobs, undefined, 'the draft is gone once saved');
});

test('two people cannot overwrite each other without noticing', async () => {
  reset();
  const first = await call(ADMIN, 'PUT', 'referrals', { baseUpdatedAt: '', answers: { r_dead: answer('ten') } });
  assert.strictEqual(first.status, 200);
  const stale = await call(ADMIN, 'PUT', 'referrals', { baseUpdatedAt: '', answers: { r_dead: answer('twenty') } });
  assert.strictEqual(stale.status, 409);
  assert.match(stale.body.error, /Reload/);
  const fresh = await call(ADMIN, 'PUT', 'referrals', { baseUpdatedAt: first.body.updatedAt, answers: { r_dead: answer('twenty') } });
  assert.strictEqual(fresh.status, 200);
  assert.strictEqual((await call(ADMIN, 'GET', 'referrals')).body.answers.answers.r_dead.text, 'twenty');
});

test('each tile keeps its own answers and files', async () => {
  reset();
  await call(ADMIN, 'PUT', 'onboarding', { baseUpdatedAt: '', answers: { o_active: answer('onboarding only') } });
  const other = (await call(ADMIN, 'GET', 'staff-information')).body;
  assert.strictEqual(other.answers.firstDraft, true);
  const up = await call(ADMIN, 'POST', 'onboarding/files', { kind: 'checklist', name: 'Checklist.docx', contentBase64: b64('fake'), confirmedNoClientInfo: true });
  assert.strictEqual(up.status, 200);
  assert.deepStrictEqual((await call(ADMIN, 'GET', 'staff-information')).body.files, []);
  assert.strictEqual((await call(ADMIN, 'GET', 'staff-information/files/' + up.body.file.id)).status, 404);
  assert.strictEqual((await call(ADMIN, 'GET', 'onboarding/files/' + up.body.file.id)).status, 200);
});

test('uploads need a confirmation, an allowed type, a kind that fits the tile, and a real size', async () => {
  reset();
  const good = { kind: 'tracking', name: 'Referral list.xlsx', contentBase64: b64('fake spreadsheet bytes'), confirmedNoClientInfo: true, note: 'fake names' };
  assert.strictEqual((await call(ADMIN, 'POST', 'referrals/files', Object.assign({}, good, { confirmedNoClientInfo: false }))).status, 400);
  assert.strictEqual((await call(ADMIN, 'POST', 'referrals/files', Object.assign({}, good, { name: 'run.exe' }))).status, 400);
  assert.strictEqual((await call(ADMIN, 'POST', 'referrals/files', Object.assign({}, good, { kind: 'nonsense' }))).status, 400);
  assert.strictEqual((await call(ADMIN, 'POST', 'onboarding/files', good)).status, 400, 'a referral file kind is not an onboarding file kind');
  assert.strictEqual((await call(ADMIN, 'POST', 'referrals/files', Object.assign({}, good, { contentBase64: '' }))).status, 400);
  assert.strictEqual((await call(ADMIN, 'POST', 'referrals/files', Object.assign({}, good, { contentBase64: 'not base64!!' }))).status, 400);
  const big = Buffer.alloc(6 * 1024 * 1024 + 10, 1).toString('base64');
  assert.strictEqual((await call(ADMIN, 'POST', 'referrals/files', Object.assign({}, good, { contentBase64: big }))).status, 400);
  const ok = await call(ADMIN, 'POST', 'referrals/files', good);
  assert.strictEqual(ok.status, 200);
  assert.strictEqual(ok.body.file.uploadedBy, ADMIN);
});

test('an uploaded file downloads as an attachment, never in the page, and can be removed', async () => {
  reset();
  const up = await call(ADMIN, 'POST', 'staff-expectations/files', { kind: 'mileage-form', name: '../../evil name.xlsx', contentBase64: b64('hello bytes'), confirmedNoClientInfo: true });
  assert.strictEqual(up.status, 200);
  assert.ok(!up.body.file.name.includes('/'));
  const got = await call(ADMIN, 'GET', 'staff-expectations/files/' + up.body.file.id);
  assert.strictEqual(got.status, 200);
  assert.strictEqual(got.headers['Content-Type'], 'application/octet-stream');
  assert.match(got.headers['Content-Disposition'], /^attachment;/);
  assert.strictEqual(got.headers['X-Content-Type-Options'], 'nosniff');
  assert.strictEqual(Buffer.from(got.body).toString(), 'hello bytes');
  assert.strictEqual((await call(ADMIN, 'DELETE', 'staff-expectations/files/' + up.body.file.id)).status, 200);
  assert.strictEqual((await call(ADMIN, 'GET', 'staff-expectations/files/' + up.body.file.id)).status, 404);
  assert.strictEqual((await call(ADMIN, 'DELETE', 'staff-expectations/files/' + up.body.file.id)).status, 404);
});

test('the build prompt carries the rules, the answers, the files, and what is still open', async () => {
  reset();
  const first = await call(ADMIN, 'GET', 'referrals/prompt');
  assert.strictEqual(first.status, 200);
  const draft = first.body.prompt;
  assert.match(draft, /# Build request: Referrals/);
  assert.match(draft, /Nobody has reviewed or saved these answers yet/);
  assert.match(draft, /at most five steps/);
  assert.match(draft, /Do not merge it/);
  assert.match(draft, /28 working day/);
  assert.match(draft, /Never put secrets/);
  assert.match(draft, /Do not change the fixed SharePoint folder paths/);
  assert.match(draft, /What already exists for this tile/);
  assert.match(draft, /Not answered/);
  assert.match(draft, /\(not sure\)/);
  assert.ok(first.body.missing.some((item) => item.reason === 'not-sure'));

  await call(ADMIN, 'PUT', 'onboarding', { baseUpdatedAt: '', answers: { o_active: answer('Background check and signed policies'), o_expire: answer('', [], true) } });
  await call(ADMIN, 'POST', 'onboarding/files', { kind: 'checklist', name: 'Checklist.docx', note: 'Fake data', contentBase64: b64('x'), confirmedNoClientInfo: true });
  const saved = (await call(ADMIN, 'GET', 'onboarding/prompt')).body;
  assert.match(saved.prompt, /Background check and signed policies/);
  assert.match(saved.prompt, /Last saved by skessler@emmeco.org/);
  assert.doesNotMatch(saved.prompt, /Nobody has reviewed/);
  assert.match(saved.prompt, /`Checklist.docx`: A checklist you use now/);
  assert.match(saved.prompt, /Marked "not sure yet"/);
  assert.match(saved.prompt, /Employees\/<employee name>\/Onboarding\/CurrentDocuments/);
  assert.doesNotMatch(saved.prompt, /Referrals/, 'no other tile leaks in');
});

test('items left blank or marked not sure are listed so the builder asks about them', () => {
  const def = tile('staff-information');
  const missing = missingItems(def, { s_see: answer('', ['own']), s_admin_who: answer('', [], true) });
  assert.ok(!missing.some((item) => item.question === def.questions.find((q) => q.id === 's_see').text));
  assert.strictEqual(missing.find((item) => item.question === def.questions.find((q) => q.id === 's_admin_who').text).reason, 'not-sure');
  assert.strictEqual(missing.length, def.questions.length - 1);
  assert.match(buildTilePrompt(def, null, [], new Date('2026-11-14T12:00:00Z')), /Generated 2026-11-14/);
});

test('the questions are written in plain words with no abbreviations', () => {
  const everything = [];
  const grab = (value) => { if (value) everything.push(String(value)); };
  TILE_SETUPS.forEach((def) => {
    grab(def.title);
    grab(def.goal);
    def.questions.forEach((q) => { grab(q.text); grab(q.help); (q.options || []).forEach((o) => grab(o.label)); });
    def.sections.forEach((s) => { grab(s.title); grab(s.intro); });
    def.fileKinds.forEach((k) => grab(k.label));
  });
  const offenders = everything.filter((text) => /\b[A-Z]{2,}\b/.test(text.replace(/\bIntakeQ\b/g, '')));
  assert.deepStrictEqual(offenders, []);
});

test('every question can be saved and shown', () => {
  const allIds = new Set();
  TILE_SETUPS.forEach((def) => {
    assert.ok(def.fileKinds.some((k) => k.id === 'other'), def.id);
    assert.ok(def.existing.length > 0, def.id);
    def.sections.forEach((s) => assert.ok(def.questions.some((q) => q.section === s.id), def.id + ' ' + s.id));
    def.questions.forEach((q) => {
      assert.ok(!allIds.has(q.id), 'duplicate ' + q.id);
      allIds.add(q.id);
      assert.ok(def.sections.some((s) => s.id === q.section), q.id);
      if (q.kind === 'text') assert.ok(!q.options, q.id);
      else assert.ok(q.options && q.options.length >= 2, q.id);
      if (q.seed) {
        const valid = (q.options || []).map((o) => o.id);
        (q.seed.choice || []).forEach((id) => assert.ok(valid.includes(id), q.id + ' seed ' + id));
        if (q.kind === 'one') assert.ok((q.seed.choice || []).length <= 1, q.id);
      }
    });
  });
});
