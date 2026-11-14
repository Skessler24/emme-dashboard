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
const scheduleMerge = require('./merge');

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
        return !!slot;
      });
    });
  });
}

function isSummerBoard(name) {
  const text = String(name || '').toLowerCase().replace(/['’]/g, '');
  return text.indexOf('summer') !== -1 && text.indexOf('26') !== -1;
}

function isIrregularBoard(name) {
  return String(name || '').toLowerCase().indexOf('not in regular spot') !== -1;
}

function shapedBoards(boards) {
  const next = (Array.isArray(boards) ? boards : []).filter(function (board) {
    return board && !isSummerBoard(board.name);
  });
  let irregular = null;
  for (let i = 0; i < next.length; i++) {
    if (isIrregularBoard(next[i].name)) {
      irregular = next[i];
      break;
    }
  }
  if (!irregular) {
    next.push({
      id: 'board_' + crypto.randomBytes(4).toString('hex'),
      name: 'Clients not in regular spot',
      days: [{ id: 'day_' + crypto.randomBytes(4).toString('hex'), name: 'Clients', color: 'blue', slots: [] }]
    });
  } else if (!Array.isArray(irregular.days) || !irregular.days.length) {
    irregular.days = [{ id: 'day_' + crypto.randomBytes(4).toString('hex'), name: 'Clients', color: 'blue', slots: [] }];
  }
  return next;
}

function normalizeWorkspace(raw) {
  const src = raw && typeof raw === 'object' ? raw : {};
  return {
    version: Number(src.version) || 1,
    updatedAt: src.updatedAt || new Date().toISOString(),
    boards: shapedBoards(src.boards),
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


const PULL_THROTTLE_MS = process.env.SCHEDULE_PULL_THROTTLE_MS === undefined ? 5000 : Number(process.env.SCHEDULE_PULL_THROTTLE_MS) || 0;
const lastPull = new Map();
const locks = new Map();

function withLock(key, fn) {
  const previous = locks.get(key) || Promise.resolve();
  const run = previous.catch(function () {}).then(fn);
  const tail = run.catch(function () {});
  locks.set(key, tail);
  tail.then(function () { if (locks.get(key) === tail) locks.delete(key); });
  return run;
}

function loadSyncModule() {
  try {
    return require('../dist/scheduleSync');
  } catch (_) {
    return null;
  }
}

function syncBlobName(principal) {
  return 'schedules/sync/' + safeBlobKey((principal.emails && principal.emails[0]) || principal.userId) + '.json';
}

async function readJsonBlob(blob) {
  if (!(await blob.exists())) return {};
  const download = await blob.download();
  try {
    const parsed = JSON.parse(await streamToString(download.readableStreamBody));
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch (_) {
    return {};
  }
}

async function writeJsonBlob(blob, value) {
  const payload = JSON.stringify(value);
  await blob.upload(payload, Buffer.byteLength(payload), {
    blobHTTPHeaders: { blobContentType: 'application/json' }
  });
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
    const syncBlob = container.getBlockBlobClient(syncBlobName(principal));
    const lockKey = emailBlobName(principal);
    const ownerEmail = (principal.emails.find(function (email) { return /@emmeco\.org$/i.test(email); }) || principal.emails[0] || '');

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

    async function pullFromSharePoint(workspace, force) {
      const sync = loadSyncModule();
      if (!sync || !ownerEmail || !workspace || !(workspace.boards || []).length) {
        return { workspace: workspace, status: 'skipped' };
      }
      const now = Date.now();
      if (!force && now - (lastPull.get(lockKey) || 0) < PULL_THROTTLE_MS) {
        return { workspace: workspace, status: 'throttled' };
      }
      lastPull.set(lockKey, now);
      const meta = await readJsonBlob(syncBlob);
      const result = await sync.pullSchedule(workspace, ownerEmail, meta);
      if (result.status === 'imported' && result.workspace) {
        const next = normalizeWorkspace(result.workspace);
        next.ownerEmail = ownerEmail;
        next.updatedAt = new Date().toISOString();
        next.version = (Number(workspace.version) || 0) + 1;
        await writeWorkspace(personal, next, ownerEmail);
        await writeJsonBlob(syncBlob, result.meta);
        return { workspace: next, status: 'imported' };
      }
      if (JSON.stringify(result.meta || {}) !== JSON.stringify(meta)) {
        await writeJsonBlob(syncBlob, result.meta);
      }
      return { workspace: workspace, status: result.status, error: result.error };
    }

    async function pullAndReconcile(loaded, force) {
      const pulled = await pullFromSharePoint(loaded, force);
      let workspace = pulled.workspace;
      if (pulled.status === 'unchanged' || pulled.status === 'imported') {
        const sync = loadSyncModule();
        const meta = sync ? await readJsonBlob(syncBlob) : null;
        if (sync && sync.needsPush(workspace, meta)) {
          workspace = (await pushToSharePoint(workspace)).workspace;
        }
      }
      return workspace;
    }

    async function bootstrapFromSharePoint() {
      const sync = loadSyncModule();
      if (!sync || !ownerEmail) return null;
      const meta = await readJsonBlob(syncBlob);
      const result = await sync.importInitialSchedule(ownerEmail, meta);
      if (result.status !== 'imported' || !result.workspace) return null;
      const next = normalizeWorkspace(result.workspace);
      next.ownerEmail = ownerEmail;
      next.updatedAt = new Date().toISOString();
      next.version = 1;
      await writeWorkspace(personal, next, ownerEmail);
      await writeJsonBlob(syncBlob, result.meta);
      return next;
    }

    async function pushToSharePoint(workspace) {
      const sync = loadSyncModule();
      if (!sync) {
        try {
          const workbook = require('../dist/scheduleWorkbook');
          return { workspace: workspace, result: await workbook.syncScheduleWorkbook(workspace, ownerEmail) };
        } catch (err) {
          return { workspace: workspace, result: { ok: false, error: String((err && err.message) || 'SharePoint update failed').slice(0, 200) } };
        }
      }
      let current = workspace;
      let result = { ok: false, skipped: true };
      for (let attempt = 0; attempt < 3; attempt++) {
        const meta = await readJsonBlob(syncBlob);
        result = await sync.pushSchedule(current, ownerEmail, meta);
        if (result.conflict) {
          const pulled = await pullFromSharePoint(current, true);
          current = pulled.workspace;
          continue;
        }
        if (result.meta && JSON.stringify(result.meta) !== JSON.stringify(meta)) {
          await writeJsonBlob(syncBlob, result.meta);
        }
        break;
      }
      if (result.conflict) result = { ok: false, error: 'The SharePoint schedule keeps changing. It will sync on the next save.' };
      return { workspace: current, result: { ok: !!result.ok, skipped: result.skipped, path: result.path, error: result.error } };
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
      const workspace = await withLock(lockKey, async function () {
        const loaded = await loadForProvider();
        return pullAndReconcile(loaded, false);
      });
      json(context, 200, {
        configured: true,
        exists: hasProviderData(workspace) || (workspace.boards || []).length > 0,
        updatedAt: workspace.updatedAt || null
      });
      return;
    }

    if (method === 'GET') {
      const workspace = await withLock(lockKey, async function () {
        const loaded = await loadForProvider();
        if (!(loaded.boards || []).length) {
          const fromSheet = await bootstrapFromSharePoint();
          if (fromSheet) return fromSheet;
        }
        return pullAndReconcile(loaded, true);
      });
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
      const baseCopy = body.base && typeof body.base === 'object' ? body.base : null;
      const outcome = await withLock(lockKey, async function () {
        const hasStored = await personal.exists();
        let current = hasStored ? await loadForProvider() : null;
        if (current) current = (await pullFromSharePoint(current, true)).workspace;

        let workspace = normalizeWorkspace(incoming);
        let merged = false;
        if (current && baseCopy && Number(baseCopy.version) !== Number(current.version)) {
          workspace = normalizeWorkspace(scheduleMerge.merge(baseCopy, workspace, current));
          merged = true;
        }
        workspace.ownerEmail = ownerEmail;
        workspace.updatedAt = new Date().toISOString();
        workspace.version = (current ? Number(current.version) || 0 : 0) + 1;
        await writeWorkspace(personal, workspace, ownerEmail);

        const pushed = await pushToSharePoint(workspace);
        workspace = pushed.workspace;
        return { workspace: workspace, merged: merged, sharePoint: pushed.result };
      });
      const workspace = outcome.workspace;
      let mileage = { ok: false, skipped: true };
      let oneDrive = { ok: false, skipped: true };
      await Promise.all([
        (async function () {
          try {
            const mileageBook = require('../dist/mileageWorkbook');
            mileage = await mileageBook.syncMileageWorkbook(workspace, ownerEmail);
          } catch (err) {
            mileage = { ok: false, error: String((err && err.message) || 'Mileage update failed').slice(0, 200) };
          }
        })(),
        (async function () {
          try {
            const drive = require('../microsoft/oneDrive');
            oneDrive = await drive.saveProviderFile(
              ownerEmail,
              'Schedule.json',
              Buffer.from(JSON.stringify(workspace)),
              'application/json'
            );
          } catch (err) {
            oneDrive = { ok: false, error: String((err && err.message) || 'OneDrive save failed').slice(0, 220) };
          }
        })()
      ]);
      json(context, 200, {
        configured: true,
        workspace: workspace,
        merged: outcome.merged,
        owner: ownerEmail,
        sharePoint: outcome.sharePoint,
        mileage: mileage,
        oneDrive: oneDrive
      });
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
