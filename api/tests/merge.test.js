const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const { merge } = require('../schedule/merge');

function slot(id, notes, extra) {
  return Object.assign({ id, time: '9:00 AM', status: '', notes: notes || '', lastDay: '' }, extra || {});
}
function ws(slots, clients) {
  return {
    version: 1,
    boards: [{ id: 'b1', name: 'Current Schedule', days: [{ id: 'd1', name: 'Monday', color: 'blue', slots }] }],
    clients: clients || [],
    dxCodes: []
  };
}
const slotsOf = (w) => w.boards[0].days[0].slots;

test('the browser copy matches the server copy', () => {
  const a = fs.readFileSync(path.join(__dirname, '../schedule/merge.js'), 'utf8');
  const b = fs.readFileSync(path.join(__dirname, '../../schedule-merge.js'), 'utf8');
  assert.strictEqual(a, b);
});

test('keeps edits from both sides in different rows', () => {
  const base = ws([slot('a', 'A'), slot('b', 'B')]);
  const local = ws([slot('a', 'A mine'), slot('b', 'B')]);
  const remote = ws([slot('a', 'A'), slot('b', 'B theirs')]);
  const out = slotsOf(merge(base, local, remote));
  assert.deepStrictEqual(out.map((s) => s.notes), ['A mine', 'B theirs']);
});

test('keeps rows added on both sides and blank notes', () => {
  const base = ws([slot('a', 'A')]);
  const local = ws([slot('a', 'A'), slot('n1', '')]);
  const remote = ws([slot('a', 'A'), slot('r1', 'Remote client')]);
  const out = slotsOf(merge(base, local, remote)).map((s) => s.id);
  assert.deepStrictEqual(out.sort(), ['a', 'n1', 'r1']);
});

test('a delete on one side wins unless the other side edited the row', () => {
  const base = ws([slot('a', 'A'), slot('b', 'B'), slot('c', 'C')]);
  const local = ws([slot('b', 'B'), slot('c', 'C mine')]);
  const remote = ws([slot('a', 'A edited'), slot('b', 'B')]);
  const out = slotsOf(merge(base, local, remote));
  assert.deepStrictEqual(out.map((s) => s.id).sort(), ['a', 'b', 'c']);
  assert.strictEqual(out.find((s) => s.id === 'a').notes, 'A edited');
  const gone = slotsOf(merge(base, ws([slot('b', 'B'), slot('c', 'C')]), ws([slot('a', 'A'), slot('b', 'B'), slot('c', 'C')])));
  assert.deepStrictEqual(gone.map((s) => s.id), ['b', 'c']);
});

test('same field changed on both sides: local wins', () => {
  const base = ws([slot('a', 'A')]);
  const out = slotsOf(merge(base, ws([slot('a', 'mine')]), ws([slot('a', 'theirs')])));
  assert.strictEqual(out[0].notes, 'mine');
});

test('a reorder on one side is kept', () => {
  const base = ws([slot('a'), slot('b'), slot('c')]);
  const local = ws([slot('c'), slot('a'), slot('b')]);
  const remote = ws([slot('a'), slot('b'), slot('c'), slot('d', 'new')]);
  const out = slotsOf(merge(base, local, remote)).map((s) => s.id);
  assert.deepStrictEqual(out, ['c', 'a', 'b', 'd']);
});

test('merging with identical copies changes nothing and handles 60 rows', () => {
  const rows = Array.from({ length: 60 }, (_, i) => slot('s' + i, 'Client ' + i));
  const base = ws(rows);
  const out = merge(base, ws(rows.slice()), ws(rows.slice()));
  assert.deepStrictEqual(slotsOf(out), rows);
});

test('clients merge by id', () => {
  const base = ws([], [{ id: 'c1', name: 'Sam' }]);
  const local = ws([], [{ id: 'c1', name: 'Sam', phone: '1' }, { id: 'c2', name: 'New' }]);
  const remote = ws([], [{ id: 'c1', name: 'Samuel' }]);
  const out = merge(base, local, remote).clients;
  assert.strictEqual(out.length, 2);
  assert.deepStrictEqual(out[0], { id: 'c1', name: 'Samuel', phone: '1' });
});
