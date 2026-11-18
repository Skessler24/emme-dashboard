/**
 * Reads the current SharePoint guidebook and returns it as themed sections.
 * The Word document stays the source. This does not change that file.
 *
 * Where the document is found, in order:
 *   1. The GUIDEBOOK_URL application setting (a sharing link), if one is set.
 *   2. A Word file with "guidebook" in its name in the Emme team site Documents library.
 *   3. The old sharing link on Jenny's OneDrive, as a last resort.
 */
const { requireUser } = require('../dist/auth');
const { graphGet, graphGetBytes, GraphError, candidateSiteIds } = require('../dist/graphClient');
const { docxToHtml } = require('./docx');
const { splitSections } = require('./render');

const OLD_URL = 'https://emmeco031-my.sharepoint.com/:w:/g/personal/jenriat_emmeco_org/IQBtpoWOjqKuQKKQa3HGgAXqARRet5OWW-QTF1gsabG4sIA?e=CYQ0Zw';
const MAX_BYTES = 15 * 1024 * 1024;
const FIND_MEMORY_MS = 5 * 60 * 1000;
const ITEM_FIELDS = '?$select=name,lastModifiedDateTime,eTag,size';

let cache = { etag: '', payload: null };
let located = { base: '', at: 0 };

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

function shareBase(url) {
  return '/shares/' + encodeURIComponent(shareId(url)) + '/driveItem';
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

function isGuidebook(item) {
  const name = String((item && item.name) || '');
  return !!(item && item.file) && /\.docx$/i.test(name) && !/^~\$/.test(name) && /guide\s*-?\s*book/i.test(name);
}

function rank(item) {
  const name = String(item.name || '').toLowerCase();
  let score = 0;
  if (name.indexOf('emme') !== -1) score += 2;
  if (name.indexOf('provider') !== -1) score += 1;
  if (/\b(old|copy|draft|archive|backup)\b/.test(name)) score -= 3;
  return score;
}

/** Looks through the Emme team site Documents library for the guidebook Word file. Returns a Graph path, or ''. */
async function findGuidebook() {
  if (located.base && Date.now() - located.at < FIND_MEMORY_MS) return located.base;
  const sites = await candidateSiteIds();
  for (let i = 0; i < sites.length; i++) {
    let data;
    try {
      data = await graphGet('/sites/' + sites[i] + "/drive/root/search(q='guidebook')?$select=id,name,file,lastModifiedDateTime&$top=50");
    } catch (err) {
      if (err instanceof GraphError && (err.status === 401 || err.status === 403 || err.status === 404)) continue;
      throw err;
    }
    const matches = ((data && data.value) || []).filter(isGuidebook).sort(function (a, b) {
      return (rank(b) - rank(a)) || (Date.parse(b.lastModifiedDateTime || '') || 0) - (Date.parse(a.lastModifiedDateTime || '') || 0);
    });
    if (matches.length) {
      located = { base: '/sites/' + sites[i] + '/drive/items/' + matches[0].id, at: Date.now() };
      return located.base;
    }
  }
  return '';
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
    let base = '';
    let item;
    if (process.env.GUIDEBOOK_URL) {
      base = shareBase(process.env.GUIDEBOOK_URL);
      item = await graphGet(base + ITEM_FIELDS);
    } else {
      base = await findGuidebook();
      if (base) {
        item = await graphGet(base + ITEM_FIELDS);
      } else {
        try {
          base = shareBase(OLD_URL);
          item = await graphGet(base + ITEM_FIELDS);
        } catch (_) {
          json(context, 404, { error: 'The guidebook was not found in the Emme Documents library. Save the Word file there with "Guidebook" in its name.' });
          return;
        }
      }
    }
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
    const bytes = await graphGetBytes(base + '/content');
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
    located = { base: '', at: 0 };
    const status = err instanceof GraphError ? err.status : 502;
    const message = err instanceof Error ? err.message : 'Could not read the guidebook';
    const needsConsent = status === 401 || status === 403;
    json(context, needsConsent ? 502 : (status || 502), {
      error: message.slice(0, 240) + (needsConsent ? ' The guidebook needs Files.ReadWrite.All or Sites.ReadWrite.All admin consent.' : ''),
    });
  }
};

module.exports.findGuidebook = findGuidebook;
