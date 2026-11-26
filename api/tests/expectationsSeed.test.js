const test = require('node:test');
const assert = require('node:assert/strict');
const { TILE_SETUPS } = require('../dist/tileSetupQuestions');

function expectations() {
  return TILE_SETUPS.find((item) => item.id === 'staff-expectations');
}

function question(id) {
  return expectations().questions.find((item) => item.id === id);
}

test('openings stay on the schedules saved in the app', () => {
  const openings = question('e_openings');
  assert.equal(openings.text, 'What should staff schedules show about therapy openings, and who looks at them?');
  assert.match(openings.seed.text, /schedules each provider saves/);
  assert.match(openings.seed.text, /as soon as a provider saves/);
  assert.doesNotMatch(openings.seed.text, /Power Automate writes/);
  assert.doesNotMatch(openings.seed.text, /SharePoint files/);
  assert.match(expectations().existing[0], /schedules saved in the app/);
  assert.match(expectations().existing[1], /EMME\/PAYROLL & MILEAGE\/Mileage/);
});

test('answered staff expectation rules are saved and question wording stays put', () => {
  assert.equal(question('e_mile_zone').text, 'How is the five mile zone measured? From where to where? Is it a straight line or driving distance?');
  assert.match(question('e_mile_zone').seed.text, /straight line/);
  assert.deepEqual(question('e_notes_when').seed.choice, ['other']);
  assert.match(question('e_notes_when').seed.text, /48 hours/);
  assert.match(question('e_cancel_types').seed.text, /not rescheduled/);
  assert.deepEqual(question('e_email_how').seed.choice, ['self']);
  assert.deepEqual(question('e_alert_how').seed.choice, ['email', 'banner']);
  assert.equal(question('e_mile_approve').seed.unsure, undefined);
  assert.match(question('e_mile_rule').seed.text, /5 miles/);
  assert.match(question('e_email').seed.text, /checked and responded to/);
});
