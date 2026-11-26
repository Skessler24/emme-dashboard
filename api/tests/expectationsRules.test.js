const test = require('node:test');
const assert = require('node:assert/strict');
const rules = require('../dist/expectationsRules');

test('five straight-line miles is inside the zone and anything over is outside', () => {
  const home = { lat: 39.7392, lon: -104.9903 };
  const inside = rules.haversineMiles(home.lat, home.lon, home.lat + (5 / 69), home.lon);
  assert.ok(inside <= 5.2 && inside >= 4.8);
  assert.equal(rules.outsideZone(5), false);
  assert.equal(rules.outsideZone(5.1), true);
  assert.equal(rules.outsideZone(rules.haversineMiles(home.lat, home.lon, home.lat, home.lon)), false);
});

test('a note is late after 48 hours and an evaluation is past due after Friday', () => {
  const now = new Date('2026-11-27T18:00:00Z');
  assert.equal(rules.noteIsLate('2026-11-25T12:00:00Z', now), true);
  assert.equal(rules.noteIsLate('2026-11-26T12:00:00Z', now), false);
  assert.equal(rules.fridayOfWeek('2026-11-23'), '2026-11-27');
  assert.equal(rules.pastEndOfWeek('2026-11-23T15:00:00Z', new Date('2026-11-28T18:00:00Z')), true);
  assert.equal(rules.pastEndOfWeek('2026-11-23T15:00:00Z', new Date('2026-11-27T18:00:00Z')), false);
});

test('cancellations without a reschedule count, and only no-shows and late cancellations are incidents', () => {
  assert.deepEqual(rules.cancellationClass('Rescheduled'), { kind: 'rescheduled', productivity: false, incident: false });
  assert.equal(rules.cancellationClass('Late Cancel').incident, true);
  assert.equal(rules.cancellationClass('No Show').incident, true);
  assert.equal(rules.cancellationClass('Client Canceled').productivity, true);
  assert.equal(rules.cancellationClass('Client Canceled').incident, false);
  assert.equal(rules.incidentLimit('weekly'), 4);
  assert.equal(rules.incidentLimit('twice-monthly'), 2);
  assert.equal(rules.incidentLimit('monthly'), 1);
  assert.match(rules.OUTREACH_MESSAGE, /another family/);
});

test('a client inside the drawn map area is in range, and one outside it is not', () => {
  const ring = [
    { lat: 39.6, lon: -105.1 },
    { lat: 39.6, lon: -104.8 },
    { lat: 39.9, lon: -104.8 },
    { lat: 39.9, lon: -105.1 },
  ];
  assert.equal(rules.pointInRing(39.75, -104.99, ring), true);
  assert.equal(rules.pointInRing(40.2, -104.99, ring), false);
  const parsed = rules.parseMapKml('<Placemark><Polygon><outerBoundaryIs><LinearRing><coordinates>-105.1,39.6,0 -104.8,39.6,0 -104.8,39.9,0 -105.1,39.9,0</coordinates></LinearRing></outerBoundaryIs></Polygon></Placemark><Placemark><Point><coordinates>-104.99,39.74,0</coordinates></Point></Placemark>');
  assert.equal(parsed.polygons.length, 1);
  assert.equal(parsed.points.length, 1);
  assert.equal(rules.mapIdFromLink('https://www.google.com/maps/d/viewer?mid=abc123'), 'abc123');
});

test('the monthly mileage ratio is high only above 1.8', () => {
  assert.equal(rules.mileageRatio(180, 100), 1.8);
  assert.equal(rules.ratioIsHigh(1.8), false);
  assert.equal(rules.ratioIsHigh(1.81), true);
  assert.equal(rules.ratioIsHigh(rules.mileageRatio(10, 0)), false);
  assert.equal(rules.isDashboardCopyName('Dashboard Mileage Copy.xlsx'), true);
  assert.equal(rules.isDashboardCopyName('Kyla Mileage 2026.xlsx'), false);
});
