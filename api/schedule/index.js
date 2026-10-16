/**
 * Emme Schedule & Caseload API
 * Stores one workspace per signed-in provider in Azure Blob Storage.
 * The original shared file stays with skessler@emmeco.org.
 *
 * Routes (via function.json route "schedule/{*rest}"):
 *   GET  /api/schedule           load this provider's workspace
 *   PUT  /api/schedule           save this provider's workspace
 *   GET  /api/schedule/meta      lightweight updatedAt check for polling
 */
const crypto = require('crypto');
const { BlobServiceClient } = require('@azure/storage-blob');

const CONTAINER = process.env.AZURE_STORAGE_CONTAINER || 'emme-reports';
const LEGACY_BLOB = process.env.AZURE_SCHEDULE_BLOB || 'emme-schedule-workspace.json';
const LEGACY_OWNER = 'skessler@emmeco.org';

function json(context, status, body) {
  context.res = {
    status,
    headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
    body
  };
}

function getBlobService() {
  const conn = process.env.AZURE_STORAGE_CONNECTION_STRING || '';
  if (!conn) return null;
  return BlobServiceClient.fromConnectionString(conn);
}

function emptyWorkspace() {
  return {
    version: 1,
    updatedAt: new Date().toISOString(),
    boards: [],
    clients: [],
    dxCodes: []
  };
}

function header(req, name) {
  const headers = (req && req.headers) || {};
  const want = name.toLowerCase();
  if (headers[want]) return String(headers[want]);
  const key = Object.keys(headers).find(function (k) { return k.toLowerCase() === want; });
  return key ? String(headers[key]) : '';
}

function readPrincipal(req) {
  const directId = header(req, 'x-ms-client-principal-id');
  const directName = header(req, 'x-ms-client-principal-name');
  const encoded = header(req, 'x-ms-client-principal');
  let parsed = null;
  if (encoded) {
    try {
      parsed = JSON.parse(Buffer.from(encoded, 'base64').toString('utf8'));
    } catch (_) {
      parsed = null;
    }
  }
  const userId = directId || (parsed && parsed.userId) || '';
  const userDetails = directName || (parsed && parsed.userDetails) || '';
  if (!userId && !userDetails) return null;
  const emails = [];
  function addEmail(value) {
    const email = String(value || '').trim().toLowerCase();
    if (email.indexOf('@') !== -1 && emails.indexOf(email) === -1) emails.push(email);
  }
  ((parsed && parsed.claims) || []).forEach(function (claim) {
    const typ = String((claim && claim.typ) || '').toLowerCase();
    const val = String((claim && claim.val) || '').trim().toLowerCase();
    if (val.indexOf('@') === -1) return;
    if (typ.indexOf('email') !== -1 || typ.indexOf('preferred_username') !== -1 || typ.endsWith('/upn')) {
      emails.unshift(val);
    } else {
      addEmail(val);
    }
  });
  addEmail(userDetails);
  const preferred = String(userDetails || '').trim().toLowerCase();
  const unique = [];
  if (preferred.indexOf('@') !== -1) unique.push(preferred);
  emails.forEach(function (email) {
    if (unique.indexOf(email) === -1) unique.push(email);
  });
  return {
    userId: String(userId || userDetails),
    userDetails: String(userDetails || ''),
    emails: unique
  };
}

function safeBlobKey(value) {
  return String(value || 'user')
    .toLowerCase()
    .replace(/[^a-z0-9._-]/g, '_')
    .slice(0, 120) || 'user';
}

function emailBlobName(principal) {
  return 'schedules/by-email/' + safeBlobKey((principal.emails && principal.emails[0]) || principal.userId) + '.json';
}

function idBlobName(principal) {
  return 'schedules/' + safeBlobKey(principal.userId) + '.json';
}

function isLegacyOwner(principal) {
  return principal.emails.indexOf(LEGACY_OWNER) !== -1;
}

function hasProviderData(workspace) {
  if (!workspace) return false;
  if ((workspace.clients || []).length) return true;
  return (workspace.boards || []).some(function (board) {
    return (board.days || []).some(function (day) {
      return (day.slots || []).some(function (slot) {
        return slot.notes || slot.clientName || slot.status;
      });
    });
  });
}

function normalizeWorkspace(raw) {
  const src = raw && typeof raw === 'object' ? raw : {};
  return {
    version: Number(src.version) || 1,
    updatedAt: src.updatedAt || new Date().toISOString(),
    boards: Array.isArray(src.boards) ? src.boards : [],
    clients: Array.isArray(src.clients) ? src.clients : [],
    dxCodes: Array.isArray(src.dxCodes) ? src.dxCodes : [],
    ownerEmail: String(src.ownerEmail || '').trim().toLowerCase()
  };
}

function belongsTo(workspace, principal) {
  if (!workspace) return false;
  const tagged = String(workspace.ownerEmail || '').toLowerCase();
  if (!tagged) return true;
  return principal.emails.indexOf(tagged) !== -1;
}

async function metadataOwner(blob) {
  if (!(await blob.exists())) return '';
  const props = await blob.getProperties();
  const meta = props.metadata || {};
  return String(meta.owner || '').toLowerCase();
}

function ownerMatches(owner, principal) {
  if (!owner) return true;
  return principal.emails.indexOf(owner) !== -1;
}

module.exports = async function (context, req) {
  if (req.method === 'OPTIONS') {
    context.res = { status: 204, headers: { 'Content-Type': 'application/json' } };
    return;
  }

  const service = getBlobService();
  if (!service) {
    json(context, 503, {
      configured: false,
      error: 'Cloud schedule sync is not configured. Add AZURE_STORAGE_CONNECTION_STRING on the Static Web App, or use device save.'
    });
    return;
  }

  try {
    const principal = readPrincipal(req);
    if (!principal) {
      json(context, 401, { configured: true, error: 'Sign in required' });
      return;
    }

    const container = service.getContainerClient(CONTAINER);
    await container.createIfNotExists();
    const personal = container.getBlockBlobClient(emailBlobName(principal));
    const previous = container.getBlockBlobClient(idBlobName(principal));
    const legacy = container.getBlockBlobClient(LEGACY_BLOB);
    const ownerEmail = principal.emails[0] || '';

    const rest = String((req.params && (req.params.rest || req.params.id)) || '')
      .replace(/^\/+|\/+$/g, '');
    const method = (req.method || 'GET').toUpperCase();

    async function adopt(blob) {
      if (!(await ownerMatches(await metadataOwner(blob), principal))) return null;
      const workspace = await readWorkspace(blob);
      if (!belongsTo(workspace, principal) || !hasProviderData(workspace)) return null;
      workspace.ownerEmail = ownerEmail;
      await writeWorkspace(personal, workspace, ownerEmail);
      return workspace;
    }

    async function loadForProvider() {
      const own = await readWorkspace(personal);
      if (belongsTo(own, principal) && hasProviderData(own)) return own;
      const fromPrevious = await adopt(previous);
      if (fromPrevious) return fromPrevious;
      if (isLegacyOwner(principal)) {
        const fromLegacy = await adopt(legacy);
        if (fromLegacy) return fromLegacy;
      }
      return emptyWorkspace();
    }

    if (rest === 'share' && method === 'POST') {
      const workspace = await loadForProvider();
      const token = crypto.randomBytes(18).toString('hex');
      const record = {
        token: token,
        owner: ownerEmail,
        createdAt: new Date().toISOString(),
        workspace: workspace
      };
      const shareBlob = container.getBlockBlobClient('shares/' + token + '.json');
      const payload = JSON.stringify(record);
      await shareBlob.upload(payload, Buffer.byteLength(payload), {
        blobHTTPHeaders: { blobContentType: 'application/json' }
      });
      json(context, 201, { configured: true, token: token });
      return;
    }

    if (method === 'GET' && rest.indexOf('share/') === 0) {
      const token = rest.slice('share/'.length);
      if (!/^[a-f0-9]{20,64}$/.test(token)) {
        json(context, 404, { error: 'Share link not found' });
        return;
      }
      const shareBlob = container.getBlockBlobClient('shares/' + token + '.json');
      if (!(await shareBlob.exists())) {
        json(context, 404, { error: 'Share link not found' });
        return;
      }
      const download = await shareBlob.download();
      const text = await streamToString(download.readableStreamBody);
      let record;
      try { record = JSON.parse(text); } catch (_) { record = null; }
      if (!record || !record.workspace) {
        json(context, 404, { error: 'Share link not found' });
        return;
      }
      json(context, 200, {
        configured: true,
        owner: record.owner || '',
        createdAt: record.createdAt || '',
        workspace: normalizeWorkspace(record.workspace)
      });
      return;
    }

    if (method === 'GET' && rest === 'meta') {
      const workspace = await loadForProvider();
      json(context, 200, {
        configured: true,
        exists: hasProviderData(workspace) || (workspace.boards || []).length > 0,
        updatedAt: workspace.updatedAt || null
      });
      return;
    }

    if (method === 'GET') {
      const workspace = await loadForProvider();
      json(context, 200, { configured: true, workspace: workspace, owner: ownerEmail });
      return;
    }

    if (method === 'PUT') {
      let body = req.body;
      if (typeof body === 'string') {
        try { body = JSON.parse(body); } catch (_) { body = {}; }
      }
      body = body || {};
      const incoming = body.workspace || body;
      const workspace = normalizeWorkspace(incoming);
      workspace.ownerEmail = ownerEmail;
      workspace.updatedAt = new Date().toISOString();
      workspace.version = (Number(workspace.version) || 0) + 1;
      await writeWorkspace(personal, workspace, ownerEmail);
      let sharePoint = { ok: false, skipped: true };
      try {
        const workbook = require('../dist/scheduleWorkbook');
        sharePoint = await workbook.syncScheduleWorkbook(workspace, ownerEmail);
      } catch (err) {
        sharePoint = { ok: false, error: String((err && err.message) || 'SharePoint update failed').slice(0, 200) };
      }
      json(context, 200, { configured: true, workspace: workspace, owner: ownerEmail, sharePoint: sharePoint });
      return;
    }

    json(context, 405, { error: 'Method not allowed' });
  } catch (err) {
    json(context, 500, { error: (err && err.message) || 'Schedule sync error' });
  }
};

async function readWorkspace(blob) {
  if (!(await blob.exists())) return null;
  const download = await blob.download();
  const text = await streamToString(download.readableStreamBody);
  try {
    return normalizeWorkspace(JSON.parse(text));
  } catch (_) {
    return emptyWorkspace();
  }
}

async function writeWorkspace(blob, workspace, ownerEmail) {
  const payload = JSON.stringify(workspace);
  const owner = String(ownerEmail || '').toLowerCase().replace(/[^a-z0-9@._+-]/g, '').slice(0, 180);
  await blob.upload(payload, Buffer.byteLength(payload), {
    blobHTTPHeaders: { blobContentType: 'application/json' },
    metadata: { updatedat: workspace.updatedAt || '', owner: owner }
  });
}

function streamToString(readableStream) {
  return new Promise(function (resolve, reject) {
    const chunks = [];
    readableStream.on('data', function (data) {
      chunks.push(data.toString());
    });
    readableStream.on('end', function () {
      resolve(chunks.join(''));
    });
    readableStream.on('error', reject);
  });
}
