const test = require('node:test');
const assert = require('node:assert');
const { blobs } = require('./fakes');

process.env.EMME_ADMIN_EMAILS = 'skessler@emmeco.org';
process.env.AUTOMATION_API_KEY = 'test-automation-key';

const { dispatch } = require('../dist/dispatch');
const {
  EI_FIRST_EMAIL,
  handleReferrals,
  memoryList,
  summarizeOpenings,
  workingDaysBetween,
} = require('../dist/referralsService');

const ACTOR = { email: 'skessler@emmeco.org', name: 'Admin' };
const TODAY = '2026-11-25';

function principal(email) {
  return Buffer.from(JSON.stringify({
    userId: 'u-' + email,
    userDetails: email,
    claims: [{ typ: 'preferred_username', val: email }],
  })).toString('base64');
}

function resetStore() {
  for (const key of [...blobs.keys()]) {
    if (key.startsWith('referrals/')) blobs.delete(key);
  }
}

async function call(method, options) {
  const context = {};
  const req = {
    method: method,
    body: options.body || {},
    query: options.query || {},
    params: { rest: options.rest || '' },
    headers: {},
  };
  await handleReferrals(context, req, ACTOR, Object.assign({ today: TODAY, list: options.list }, options.deps || {}));
  return context.res;
}

test('referrals requires an Emme admin', async () => {
  async function hit(email) {
    const context = {};
    const headers = email ? { 'x-ms-client-principal': principal(email) } : {};
    await dispatch(context, { method: 'GET', headers: headers, query: {}, params: {} }, 'referrals');
    return context.res.status;
  }
  assert.strictEqual(await hit(null), 401);
  assert.strictEqual(await hit('teacher@cherrycreekschools.org'), 403);
  assert.strictEqual(await hit('person@emmeco.org'), 403);
});

test('referral tracking, drafts, and openings', async (t) => {
  t.diagnostic('sequential');

  await t.test('an Early Intervention referral keeps the seven steps and does not time any messages', async () => {
    resetStore();
    const list = memoryList();
    const created = await call('POST', {
      list: list,
      body: {
        familyName: 'Test Family',
        kind: 'ei',
        serviceCoordinator: 'Coordinator Example',
        therapies: ['slp', 'nope'],
        source: '',
        notes: 'Call after lunch',
        receivedOn: '2026-11-20',
        serviceGridOn: '2026-11-02',
      },
    });
    assert.strictEqual(created.status, 201);
    const referral = created.body.referral;
    assert.deepStrictEqual(referral.steps.map((step) => step.label), [
      'Call',
      'MID (Most Important Document), and put Medicaid and insurance on the PAR/HIA form',
      'Provider notification',
      'Welcome email',
      'Service coordinator notified',
      'Record the days it takes to place',
      'Remove from the referral list',
    ]);
    assert.strictEqual(referral.stepsUnconfirmed, false);
    assert.deepStrictEqual(referral.therapies, ['slp']);
    assert.strictEqual(referral.todos.length, 0);
    assert.ok(referral.drafts.every((draft) => draft.status === 'draft'));
    assert.strictEqual(referral.drafts.find((draft) => draft.id === 'ei-first').incomplete, true);
    assert.strictEqual(referral.drafts.find((draft) => draft.id === 'ei-first').body, EI_FIRST_EMAIL);
    assert.strictEqual(list.rows[0].fields.ReferrerName, 'Test Family');
    assert.strictEqual(list.rows[0].fields.Status, 'Active');
    assert.strictEqual(list.rows[0].fields.Notes, 'Call after lunch');
    assert.ok(!String(list.rows[0].fields.Notes).includes('EMME'));
    assert.strictEqual(JSON.stringify(referral).includes('"sent"'), false);
  });

  await t.test('a private referral uses the unconfirmed draft steps', async () => {
    resetStore();
    const list = memoryList();
    const created = await call('POST', {
      list: list,
      body: { familyName: 'Private Family', kind: 'private', receivedOn: TODAY },
    });
    assert.strictEqual(created.status, 201);
    assert.strictEqual(created.body.referral.stepsUnconfirmed, true);
    assert.deepStrictEqual(created.body.referral.steps.map((step) => step.id), ['first', 'needs', 'placement', 'follow', 'removal']);
    assert.strictEqual(created.body.referral.drafts[0].body.includes('website'), true);
    assert.strictEqual(created.body.referral.todos.length, 0);
  });

  await t.test('the old add form still stores a row, and a name is required', async () => {
    resetStore();
    const list = memoryList();
    const missing = await call('POST', { list: list, body: { notes: 'no name' } });
    assert.strictEqual(missing.status, 400);
    assert.strictEqual(missing.body.error, 'Family name is required.');
    const old = await call('POST', {
      list: list,
      body: { referrerName: 'Old Form Family', assignedTo: 'Someone', notes: 'kept', status: 'New' },
    });
    assert.strictEqual(old.status, 201);
    assert.strictEqual(old.body.referral.kind, '');
    assert.strictEqual(old.body.referral.steps.length, 0);
    assert.strictEqual(old.body.referral.assignedTo, 'Someone');
    assert.strictEqual(old.body.referral.todos.length, 0);
    assert.strictEqual(list.rows[0].fields.Notes, 'kept');
  });

  await t.test('days to place skip weekends and flag an Early Intervention referral past 28 working days', async () => {
    assert.strictEqual(workingDaysBetween('2026-11-02', '2026-11-06'), 4);
    assert.strictEqual(workingDaysBetween('2026-11-06', '2026-11-09'), 1);
    assert.strictEqual(workingDaysBetween('2026-11-25', '2026-11-25'), 0);
    assert.strictEqual(workingDaysBetween('2026-11-10', '2026-11-01'), null);
    resetStore();
    const list = memoryList();
    const created = await call('POST', {
      list: list,
      body: { familyName: 'Days Family', kind: 'ei', serviceGridOn: '2026-10-01', receivedOn: '2026-10-01' },
    });
    assert.strictEqual(created.body.referral.elapsedWorkingDays, 39);
    assert.strictEqual(created.body.referral.daysOverGoal, true);
    assert.strictEqual(created.body.referral.remainingWorkingDays, 28 - 39);
    const placed = await call('POST', {
      list: list,
      rest: created.body.referral.id,
      body: {
        action: 'save',
        baseUpdatedAt: created.body.referral.updatedAt,
        familyName: 'Days Family',
        kind: 'ei',
        serviceGridOn: '2026-11-02',
        placedOn: '2026-11-06',
        receivedOn: '2026-10-01',
      },
    });
    assert.strictEqual(placed.status, 200);
    assert.strictEqual(placed.body.referral.daysToPlace, 4);
    assert.strictEqual(placed.body.referral.daysOverGoal, false);
    assert.strictEqual(placed.body.referral.status, 'Placed');
    assert.ok(placed.body.referral.todos.some((todo) => todo.id === 'tracking-doc' && todo.owner === 'Administrator' && todo.dueOn === TODAY));
    const fresh = await call('POST', { list: list, body: { familyName: 'No Dates', kind: 'ei' } });
    const denied = await call('POST', {
      list: list,
      rest: fresh.body.referral.id,
      body: { action: 'step', stepId: 'days', doneOn: TODAY, baseUpdatedAt: fresh.body.referral.updatedAt },
    });
    assert.strictEqual(denied.status, 400);
  });

  await t.test('recording days waits for both dates, and the Most Important Document step becomes a to-do', async () => {
    resetStore();
    const list = memoryList();
    const created = await call('POST', { list: list, body: { familyName: 'Step Family', kind: 'ei' } });
    const id = created.body.referral.id;
    const denied = await call('POST', {
      list: list,
      rest: id,
      body: { action: 'step', stepId: 'days', doneOn: TODAY, baseUpdatedAt: created.body.referral.updatedAt },
    });
    assert.strictEqual(denied.status, 400);
    const mid = await call('POST', {
      list: list,
      rest: id,
      body: { action: 'step', stepId: 'mid', doneOn: TODAY, owner: 'Jordan', baseUpdatedAt: created.body.referral.updatedAt },
    });
    assert.strictEqual(mid.status, 200);
    const todo = mid.body.referral.todos.find((item) => item.id === 'par-hia');
    assert.ok(todo);
    assert.strictEqual(todo.owner, 'Jordan');
    assert.strictEqual(todo.dueOn, TODAY);
    assert.strictEqual(todo.reminderOn, TODAY);
    const removed = await call('POST', {
      list: list,
      rest: id,
      body: { action: 'step', stepId: 'remove', doneOn: TODAY, baseUpdatedAt: mid.body.referral.updatedAt },
    });
    assert.strictEqual(removed.body.referral.status, 'Removed');
    assert.strictEqual(list.rows[0].fields.Status, 'Removed');
  });

  await t.test('assigning staff does not send anything, and a late answer comes back', async () => {
    resetStore();
    const list = memoryList();
    const created = await call('POST', { list: list, body: { familyName: 'Assign Family', kind: 'ei' } });
    const id = created.body.referral.id;
    const assigned = await call('POST', {
      list: list,
      rest: id,
      body: {
        action: 'assign',
        assignedTo: 'Pat Provider',
        assignedEmail: 'Pat@EmmeCo.org',
        baseUpdatedAt: created.body.referral.updatedAt,
      },
    });
    assert.strictEqual(assigned.status, 200);
    const referral = assigned.body.referral;
    assert.strictEqual(referral.assignedTo, 'Pat Provider');
    assert.strictEqual(referral.assignedEmail, 'pat@emmeco.org');
    assert.strictEqual(list.rows[0].fields.AssignedTo, 'Pat Provider');
    assert.ok(referral.drafts.every((draft) => draft.status === 'draft'));
    assert.strictEqual(referral.drafts.filter((draft) => draft.status === 'approved').length, 0);
    assert.strictEqual(JSON.stringify(referral).includes('"sent"'), false);
    ['assign-provider', 'assign-welcome', 'assign-sc'].forEach((todoId) => {
      assert.ok(referral.todos.some((todo) => todo.id === todoId && !todo.doneOn));
    });
    const again = await call('POST', {
      list: list,
      rest: id,
      body: { action: 'assign', assignedTo: 'Pat Provider', assignedEmail: 'pat@emmeco.org', baseUpdatedAt: referral.updatedAt },
    });
    assert.strictEqual(again.body.referral.todos.filter((todo) => todo.id === 'assign-provider').length, 1);
    const removed = await call('POST', {
      list: list,
      rest: id,
      body: { action: 'step', stepId: 'remove', doneOn: TODAY, baseUpdatedAt: again.body.referral.updatedAt },
    });
    const back = await call('POST', {
      list: list,
      rest: id,
      body: { action: 'answered', baseUpdatedAt: removed.body.referral.updatedAt },
    });
    assert.strictEqual(back.body.referral.status, 'Answered');
    assert.strictEqual(back.body.referral.removedOn, '');
    assert.strictEqual(back.body.referral.steps.find((step) => step.id === 'remove').doneOn, '');
    const sara = back.body.referral.todos.find((todo) => todo.id === 'to-sara-beth');
    assert.ok(sara);
    assert.strictEqual(sara.owner, 'Sara Beth');
    assert.match(sara.title, /after being removed/);
    assert.match(sara.title, /Sara Beth/);
  });

  await t.test('a message stays a draft until a person finishes and approves it', async () => {
    resetStore();
    const list = memoryList();
    const created = await call('POST', { list: list, body: { familyName: 'Message Family', kind: 'ei' } });
    const id = created.body.referral.id;
    const blocked = await call('POST', {
      list: list,
      rest: id,
      body: { action: 'approve', draftId: 'ei-first', baseUpdatedAt: created.body.referral.updatedAt },
    });
    assert.strictEqual(blocked.status, 400);
    const finished = 'Hello, this is the finished first message.';
    const approved = await call('POST', {
      list: list,
      rest: id,
      body: { action: 'approve', draftId: 'ei-first', body: finished, baseUpdatedAt: created.body.referral.updatedAt },
    });
    assert.strictEqual(approved.status, 200);
    const draft = approved.body.referral.drafts.find((item) => item.id === 'ei-first');
    assert.strictEqual(draft.status, 'approved');
    assert.strictEqual(draft.approvedBy, 'Admin');
    assert.strictEqual(draft.body, finished);
    assert.ok(approved.body.referral.todos.some((todo) => todo.id === 'send-ei-first' && !todo.doneOn));
    assert.strictEqual(JSON.stringify(approved.body.referral).includes('"sent"'), false);
    const edited = await call('POST', {
      list: list,
      rest: id,
      body: { action: 'save-draft', draftId: 'ei-first', body: finished + ' Extra line.', baseUpdatedAt: approved.body.referral.updatedAt },
    });
    const again = edited.body.referral.drafts.find((item) => item.id === 'ei-first');
    assert.strictEqual(again.status, 'draft');
    assert.ok(!edited.body.referral.todos.some((todo) => todo.id === 'send-ei-first' && !todo.doneOn));
  });

  await t.test('a stale save is refused, and openings are counted once', async () => {
    resetStore();
    const list = memoryList();
    const created = await call('POST', { list: list, body: { familyName: 'Race Family', kind: 'private' } });
    const stale = await call('POST', {
      list: list,
      rest: created.body.referral.id,
      body: { action: 'save', familyName: 'Race Family', kind: 'private', baseUpdatedAt: '' },
    });
    assert.strictEqual(stale.status, 409);
    const slots = [];
    for (let i = 0; i < 21; i++) slots.push({ date: '2026-11-25', day: 'Wednesday', time: '9:' + i, status: 'open 4x' });
    slots.push({ date: '2026-12-20', day: 'Sunday', time: '1:00 PM', status: 'open 1x (3)' });
    slots.push({ date: '2026-11-20', day: 'Friday', time: '8:00 AM', status: 'open 4x' });
    const summary = summarizeOpenings({
      today: TODAY,
      kinds: [{ email: 'pat@emmeco.org', name: 'Pat Provider', therapies: ['slp', 'ot'] }],
      staff: [
        { staffEmail: 'pat@emmeco.org', staffName: 'Pat Provider', openings: slots },
        { staffEmail: 'sam@emmeco.org', staffName: 'Sam Provider', openings: [{ date: '2026-11-26', day: 'Thursday', time: '10:00 AM', status: 'open 4x' }] },
      ],
    });
    assert.strictEqual(summary.total, 23);
    assert.strictEqual(summary.weekTotal, 22);
    assert.strictEqual(summary.overGoal, true);
    assert.strictEqual(summary.goalMax, 20);
    assert.strictEqual(summary.disciplineCount, 2);
    assert.strictEqual(summary.byTherapy.find((row) => row.label === 'Speech').count, 22);
    assert.strictEqual(summary.byTherapy.find((row) => row.label === 'Occupational therapy').count, 22);
    assert.strictEqual(summary.unlabeled, 1);
    assert.ok(summary.byMonth.some((row) => row.month === '2026-11' && row.label === 'November 2026'));
    assert.ok(summary.byMonth.some((row) => row.month === '2026-12' && row.count === 1));
    assert.strictEqual(summary.providers.find((row) => row.staffEmail === 'pat@emmeco.org').next30, 22);
    assert.strictEqual(summary.slots.length, 23);
  });

  await t.test('opening counts come from the scheduling window and the kind of therapy set here', async () => {
    resetStore();
    const saved = await call('POST', {
      list: memoryList(),
      body: { action: 'provider-kinds', providers: [{ email: 'Pat@EmmeCo.org', name: 'Pat Provider', therapies: ['pt'] }] },
    });
    assert.strictEqual(saved.status, 200);
    const missing = await call('GET', {
      list: memoryList(),
      query: { view: 'openings' },
      deps: { loadAvailability: async () => ({ configured: false, staff: [] }) },
    });
    assert.strictEqual(missing.status, 503);
    const openings = await call('GET', {
      list: memoryList(),
      query: { view: 'openings' },
      deps: {
        loadAvailability: async () => ({
          configured: true,
          asOf: TODAY,
          staff: [{
            staffEmail: 'pat@emmeco.org',
            staffName: 'Pat Provider',
            openings: [
              { date: '2026-11-25', day: 'Wednesday', time: '9:00 AM', status: 'open 4x' },
              { date: '2026-12-02', day: 'Wednesday', time: '9:00 AM', status: 'open 4x' },
            ],
          }],
        }),
      },
    });
    assert.strictEqual(openings.status, 200);
    assert.strictEqual(openings.body.source, 'emme-scheduling-db');
    assert.strictEqual(openings.body.total, 2);
    assert.strictEqual(openings.body.weekTotal, 1);
    assert.strictEqual(openings.body.overGoal, false);
    assert.strictEqual(openings.body.disciplineCount, 1);
    assert.strictEqual(openings.body.providers[0].therapies[0], 'Physical therapy');
    assert.strictEqual(openings.body.byMonth.length, 2);
  });

  await t.test('two outreaches and 15 days move an Early Intervention family to the bottom until they confirm', async () => {
    resetStore();
    const list = memoryList();
    const waiting = await call('POST', {
      list: list,
      body: { familyName: 'Still Waiting', kind: 'ei', receivedOn: '2026-11-20' },
    });
    const waitingId = waiting.body.referral.id;
    const first = await call('POST', {
      list: list,
      rest: waitingId,
      body: { action: 'outreach', on: '2026-11-20', baseUpdatedAt: waiting.body.referral.updatedAt },
    });
    assert.strictEqual(first.status, 200);
    assert.strictEqual(first.body.referral.status, 'Active');
    assert.strictEqual(first.body.referral.outreachCount, 1);
    assert.strictEqual(first.body.referral.respondBy, '2026-12-05');
    const second = await call('POST', {
      list: list,
      rest: waitingId,
      body: { action: 'outreach', on: '2026-11-24', baseUpdatedAt: first.body.referral.updatedAt },
    });
    assert.strictEqual(second.body.referral.outreachCount, 2);
    assert.strictEqual(second.body.referral.status, 'Active');
    assert.strictEqual(second.body.referral.atBottom, false);
    assert.ok(second.body.referral.drafts.every((draft) => draft.status === 'draft'));
    assert.strictEqual(JSON.stringify(second.body.referral).includes('"sent"'), false);

    const quiet = await call('POST', {
      list: list,
      body: { familyName: 'Quiet Family', kind: 'ei', receivedOn: '2026-11-01', serviceGridOn: '2026-10-01' },
    });
    const quietId = quiet.body.referral.id;
    assert.ok(quiet.body.referral.todos.some((todo) => todo.id === 'day-goal-up'));
    const one = await call('POST', {
      list: list,
      rest: quietId,
      body: { action: 'outreach', on: '2026-11-01', baseUpdatedAt: quiet.body.referral.updatedAt },
    });
    assert.strictEqual(one.body.referral.status, 'Active');
    const two = await call('POST', {
      list: list,
      rest: quietId,
      body: { action: 'outreach', on: '2026-11-08', baseUpdatedAt: one.body.referral.updatedAt },
    });
    assert.strictEqual(two.body.referral.status, 'Bottom');
    assert.strictEqual(two.body.referral.atBottom, true);
    assert.strictEqual(two.body.referral.respondBy, '2026-11-16');
    assert.strictEqual(list.rows.find((row) => row.id === quietId).fields.Status, 'Bottom');

    const listed = await call('GET', { list: list });
    const names = listed.body.referrals.map((row) => row.familyName);
    assert.deepStrictEqual(names, ['Still Waiting', 'Quiet Family']);
    assert.ok(listed.body.sources.some((item) => item.id === 'rmhs' && item.label === 'RMHS'));

    const confirmed = await call('POST', {
      list: list,
      rest: quietId,
      body: { action: 'answered', baseUpdatedAt: two.body.referral.updatedAt },
    });
    assert.strictEqual(confirmed.body.referral.status, 'Answered');
    assert.strictEqual(confirmed.body.referral.atBottom, false);
    assert.ok(confirmed.body.referral.todos.some((todo) => todo.id === 'to-sara-beth' && todo.owner === 'Sara Beth'));

    const privateFamily = await call('POST', { list: list, body: { familyName: 'Private Family', kind: 'private' } });
    const privateOne = await call('POST', {
      list: list,
      rest: privateFamily.body.referral.id,
      body: { action: 'outreach', on: '2026-11-01', baseUpdatedAt: privateFamily.body.referral.updatedAt },
    });
    const privateTwo = await call('POST', {
      list: list,
      rest: privateFamily.body.referral.id,
      body: { action: 'outreach', on: '2026-11-02', baseUpdatedAt: privateOne.body.referral.updatedAt },
    });
    assert.strictEqual(privateTwo.body.referral.status, 'Active');
  });

  await t.test('RMHS follow-up and Developmental Pathways expiry become to-dos, and website referrals skip the broker line', async () => {
    resetStore();
    const list = memoryList();
    const rmhs = await call('POST', {
      list: list,
      body: { familyName: 'RMHS Family', kind: 'ei', source: 'RMHS', receivedOn: '2026-11-20' },
    });
    assert.strictEqual(rmhs.body.referral.source, 'rmhs');
    const follow = rmhs.body.referral.todos.find((todo) => todo.id === 'rmhs-followup');
    assert.ok(follow);
    assert.strictEqual(follow.dueOn, '2026-11-22');
    assert.strictEqual(follow.reminderOn, '2026-11-22');
    assert.ok(rmhs.body.referral.paperworkLine.includes('referrals@emmeco.org'));
    const done = await call('POST', {
      list: list,
      rest: rmhs.body.referral.id,
      body: { action: 'todo-done', todoId: 'rmhs-followup', baseUpdatedAt: rmhs.body.referral.updatedAt },
    });
    const next = done.body.referral.todos.find((todo) => todo.id !== 'rmhs-followup' && todo.id.indexOf('rmhs-followup') === 0 && !todo.doneOn);
    assert.ok(next);
    assert.strictEqual(next.dueOn, '2026-11-27');

    const dp = await call('POST', {
      list: list,
      body: { familyName: 'DP Family', kind: 'ei', source: 'dp', receivedOn: '2026-11-22' },
    });
    assert.strictEqual(dp.body.referral.dpExpiresOn, '2026-11-24');
    assert.ok(dp.body.referral.todos.some((todo) => todo.id === 'dp-expire' && todo.dueOn === '2026-11-24'));

    const site = await call('POST', {
      list: list,
      body: { familyName: 'Site Family', kind: 'private', source: 'website' },
    });
    assert.strictEqual(site.body.referral.paperworkLine, '');
    assert.ok(site.body.referral.sourceNote.includes('pediatrician'));
    assert.strictEqual(site.body.referral.todos.length, 0);
  });
});
