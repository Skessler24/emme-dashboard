const test = require('node:test');
const assert = require('node:assert/strict');
process.env.SHAREPOINT_HOSTNAME = 'example.sharepoint.com';
process.env.SHAREPOINT_SITE_PATH = '/sites/Teamsite';
require('./fakes');

const { forgetFolderSites, resetGraphBreaker } = require('../dist/graphClient');

test('the guidebook is found in the shared Documents library by its name', async () => {
  forgetFolderSites();
  resetGraphBreaker();
  let searches = 0;
  global.fetch = async function (url) {
    const text = decodeURIComponent(String(url));
    const reply = (status, body) => new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
    if (text.includes('/oauth2/v2.0/token')) return reply(200, { access_token: 't', expires_in: 3600 });
    if (text.includes('/personal/jenriat_emmeco_org')) return reply(404, { error: { message: 'gone' } });
    if (text.includes('example.sharepoint.com:/sites/Teamsite')) return reply(200, { id: 'shared1' });
    if (text.includes('/sites/shared1/drive/root/search')) {
      searches += 1;
      return reply(200, {
        value: [
          { id: 'a1', name: '~$Guidebook.docx', file: {}, lastModifiedDateTime: '2026-11-17T10:00:00Z' },
          { id: 'a2', name: 'Guidebook old copy.docx', file: {}, lastModifiedDateTime: '2026-11-17T10:00:00Z' },
          { id: 'a3', name: 'EMME Provider Guidebook.docx', file: {}, lastModifiedDateTime: '2026-10-01T10:00:00Z' },
          { id: 'a4', name: 'Guidebook.pdf', file: {}, lastModifiedDateTime: '2026-11-17T10:00:00Z' },
          { id: 'a5', name: 'Meeting notes.docx', file: {}, lastModifiedDateTime: '2026-11-17T10:00:00Z' },
          { id: 'a6', name: 'Guidebook folder', folder: {}, lastModifiedDateTime: '2026-11-17T10:00:00Z' },
        ],
      });
    }
    return reply(404, { error: { message: 'no' } });
  };
  const { findGuidebook } = require('../guidebook/index.js');
  assert.equal(await findGuidebook(), '/sites/shared1/drive/items/a3');
  // The answer is remembered for a few minutes.
  assert.equal(await findGuidebook(), '/sites/shared1/drive/items/a3');
  assert.equal(searches, 1);
});
