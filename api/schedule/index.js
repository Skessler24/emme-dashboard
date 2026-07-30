/**
 * Emme Schedule & Caseload API
 * Stores one shared workspace document in Azure Blob Storage when
 * AZURE_STORAGE_CONNECTION_STRING is set. Without storage, returns 503
 * so the client can use local browser storage.
 *
 * Routes (via function.json route "schedule/{*rest}"):
 *   GET  /api/schedule           load workspace
 *   PUT  /api/schedule           save workspace { boards, clients, ... }
 *   GET  /api/schedule/meta      lightweight updatedAt check for polling
 */
const { BlobServiceClient } = require('@azure/storage-blob');

const CONTAINER = process.env.AZURE_STORAGE_CONTAINER || 'emme-reports';
const BLOB_NAME = process.env.AZURE_SCHEDULE_BLOB || 'emme-schedule-workspace.json';

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

function normalizeWorkspace(raw) {
  const src = raw && typeof raw === 'object' ? raw : {};
  return {
    version: Number(src.version) || 1,
    updatedAt: src.updatedAt || new Date().toISOString(),
    boards: Array.isArray(src.boards) ? src.boards : [],
    clients: Array.isArray(src.clients) ? src.clients : [],
    dxCodes: Array.isArray(src.dxCodes) ? src.dxCodes : []
  };
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
    const container = service.getContainerClient(CONTAINER);
    await container.createIfNotExists();
    const blob = container.getBlockBlobClient(BLOB_NAME);

    const rest = String((req.params && (req.params.rest || req.params.id)) || '')
      .replace(/^\/+|\/+$/g, '');
    const method = (req.method || 'GET').toUpperCase();

    if (method === 'GET' && rest === 'meta') {
      if (!(await blob.exists())) {
        json(context, 200, { configured: true, exists: false, updatedAt: null });
        return;
      }
      const props = await blob.getProperties();
      json(context, 200, {
        configured: true,
        exists: true,
        updatedAt: props.metadata && props.metadata.updatedat
          ? props.metadata.updatedat
          : (props.lastModified && props.lastModified.toISOString()) || null,
        etag: props.etag || null
      });
      return;
    }

    if (method === 'GET') {
      if (!(await blob.exists())) {
        const empty = emptyWorkspace();
        json(context, 200, { configured: true, workspace: empty });
        return;
      }
      const download = await blob.download();
      const text = await streamToString(download.readableStreamBody);
      let parsed;
      try {
        parsed = JSON.parse(text);
      } catch (_) {
        parsed = emptyWorkspace();
      }
      json(context, 200, { configured: true, workspace: normalizeWorkspace(parsed) });
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
      workspace.updatedAt = new Date().toISOString();
      workspace.version = (Number(workspace.version) || 0) + 1;

      const payload = JSON.stringify(workspace);
      await blob.upload(payload, Buffer.byteLength(payload), {
        blobHTTPHeaders: { blobContentType: 'application/json' },
        metadata: { updatedat: workspace.updatedAt }
      });
      json(context, 200, { configured: true, workspace });
      return;
    }

    json(context, 405, { error: 'Method not allowed' });
  } catch (err) {
    json(context, 500, { error: (err && err.message) || 'Schedule sync error' });
  }
};

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
