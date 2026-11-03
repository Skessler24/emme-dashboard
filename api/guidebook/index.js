/**
 * Reads the current SharePoint guidebook and returns it as themed sections.
 * The Word document stays the source. This does not change that file.
 */
const { requireUser } = require('../dist/auth');
const { graphGet, graphGetBytes, GraphError } = require('../dist/graphClient');
const { docxToHtml } = require('./docx');
const { splitSections } = require('./render');

const GUIDE_URL = 'https://emmeco031-my.sharepoint.com/:w:/g/personal/jenriat_emmeco_org/IQBtpoWOjqKuQKKQa3HGgAXqARRet5OWW-QTF1gsabG4sIA?e=CYQ0Zw';
const MAX_BYTES = 15 * 1024 * 1024;

let cache = { etag: '', payload: null };

function json(context, status, body) {
  context.res = {
    status: status,
    headers: { 'Content-Type': 'application/json', 'Cache-Control': 'private, max-age=60' },
    body: body,
  };
}

function shareId(url) {
  return 'u!' + Buffer.from(url, 'utf8').toString('base64').replace(/=+$/g, '').replace(/\//g, '_').replace(/\+/g, '-');
}

function updatedLabel(iso) {
  const when = new Date(iso);
  if (Number.isNaN(when.getTime())) return '';
  return new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/Denver',
    month: 'long',
    day: 'numeric',
    year: 'numeric',
  }).format(when);
}

module.exports = async function (context, req) {
  if ((req.method || 'GET').toUpperCase() === 'OPTIONS') {
    context.res = { status: 204 };
    return;
  }
  const principal = await requireUser(context, req);
  if (!principal) return;
  if (!/@emmeco\.org$/i.test(principal.email || '')) {
    json(context, 403, { error: 'Sign in with your emmeco.org account to read the guidebook.' });
    return;
  }
  try {
    const id = encodeURIComponent(shareId(GUIDE_URL));
    const item = await graphGet('/shares/' + id + '/driveItem?$select=name,lastModifiedDateTime,eTag,size');
    const etag = String((item && item.eTag) || '');
    if (etag && cache.etag === etag && cache.payload) {
      json(context, 200, cache.payload);
      return;
    }
    const size = Number(item && item.size) || 0;
    if (size > MAX_BYTES) {
      json(context, 413, { error: 'The guidebook file is too large to show here.' });
      return;
    }
    const bytes = await graphGetBytes('/shares/' + id + '/driveItem/content');
    const shaped = splitSections(docxToHtml(Buffer.from(bytes)));
    const payload = {
      title: shaped.coverTitle || 'Provider Guidebook',
      updated: (item && item.lastModifiedDateTime) || '',
      updatedLabel: updatedLabel(item && item.lastModifiedDateTime),
      sourceName: String((item && item.name) || 'Guidebook'),
      sections: shaped.sections,
    };
    cache = { etag: etag, payload: payload };
    json(context, 200, payload);
  } catch (err) {
    const status = err instanceof GraphError ? err.status : 502;
    const message = err instanceof Error ? err.message : 'Could not read the guidebook';
    json(context, status === 401 || status === 403 ? 502 : (status || 502), {
      error: message.slice(0, 240) + ' The guidebook needs Files.ReadWrite.All or Sites.ReadWrite.All admin consent.',
    });
  }
};
