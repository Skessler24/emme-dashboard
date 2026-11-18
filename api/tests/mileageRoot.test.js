const test = require('node:test');
const assert = require('node:assert/strict');
process.env.MILEAGE_FOLDER_URL = 'https://example.sharepoint.com/:f:/g/personal/someone/FakeLinkValue';
require('./fakes');

test('a sharing link to the Mileage folder sets where the mileage files are looked for', async () => {
  let shareCalls = 0;
  global.fetch = async function (url) {
    const text = decodeURIComponent(String(url));
    const reply = (status, body) => new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
    if (text.includes('/oauth2/v2.0/token')) return reply(200, { access_token: 't', expires_in: 3600 });
    if (text.includes('/shares/')) {
      shareCalls += 1;
      return reply(200, { name: 'Mileage', parentReference: { path: '/drive/root:/Payroll%20Files/Staff' } });
    }
    return reply(404, { error: { message: 'no' } });
  };
  const { MILEAGE_ROOT, resolveMileageRoot } = require('../dist/providerSetup');
  assert.deepEqual(MILEAGE_ROOT, ['EMME', 'PAYROLL & MILEAGE', 'Mileage']);
  await resolveMileageRoot();
  assert.deepEqual(MILEAGE_ROOT, ['Payroll Files', 'Staff', 'Mileage']);
  await resolveMileageRoot();
  assert.equal(shareCalls, 1, 'the answer is kept for a few minutes');
});

test('a link that cannot be read leaves the fixed folder path alone', async () => {
  const providerSetup = require('../dist/providerSetup');
  const before = providerSetup.MILEAGE_ROOT.slice();
  global.fetch = async function (url) {
    const text = String(url);
    if (text.includes('/oauth2/v2.0/token')) return new Response(JSON.stringify({ access_token: 't', expires_in: 3600 }), { status: 200 });
    return new Response(JSON.stringify({ error: { message: 'no' } }), { status: 404 });
  };
  await providerSetup.resolveMileageRoot();
  assert.deepEqual(providerSetup.MILEAGE_ROOT, before);
});
