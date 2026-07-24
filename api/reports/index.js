/**
 * Emme Report Library API
 * Stores generated reports in Azure Blob Storage (cloud) when
 * AZURE_STORAGE_CONNECTION_STRING is set.
 * Falls back to in-memory is not used — without storage, returns 503
 * so the client can use local browser storage.
 *
 * Routes (via function.json route "reports/{*rest}"):
 *   GET    /api/reports           list
 *   POST   /api/reports           create { title, client, provider, reportType, content }
 *   GET    /api/reports/{id}      get one
 *   DELETE /api/reports/{id}      delete
 */
const { BlobServiceClient } = require('@azure/storage-blob');

const CONTAINER = process.env.AZURE_STORAGE_CONTAINER || 'emme-reports';

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

function safeId() {
  return 'rpt_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 8);
}

function parsePath(req) {
  // rest may be "id" or empty
  const raw = (req.params && (req.params.rest || req.params.id)) || '';
  return String(raw).replace(/^\/+|\/+$/g, '');
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
      error: 'Cloud report library is not configured. Add AZURE_STORAGE_CONNECTION_STRING on the Static Web App, or use device save.'
    });
    return;
  }

  try {
    const container = service.getContainerClient(CONTAINER);
    await container.createIfNotExists();

    const id = parsePath(req);
    const method = (req.method || 'GET').toUpperCase();

    if (method === 'GET' && !id) {
      const items = [];
      for await (const blob of container.listBlobsFlat()) {
        if (!blob.name.endsWith('.json')) continue;
        try {
          const client = container.getBlockBlobClient(blob.name);
          const download = await client.download();
          const text = await streamToString(download.readableStreamBody);
          const meta = JSON.parse(text);
          items.push({
            id: meta.id || blob.name.replace(/\.json$/, ''),
            title: meta.title || 'Untitled report',
            client: meta.client || '',
            provider: meta.provider || '',
            reportType: meta.reportType || '',
            createdAt: meta.createdAt || blob.properties.createdOn || null,
            updatedAt: meta.updatedAt || blob.properties.lastModified || null
          });
        } catch (_) { /* skip bad blob */ }
      }
      items.sort(function (a, b) {
        return String(b.createdAt || '').localeCompare(String(a.createdAt || ''));
      });
      json(context, 200, { configured: true, items });
      return;
    }

    if (method === 'GET' && id) {
      const client = container.getBlockBlobClient(id + '.json');
      if (!(await client.exists())) {
        json(context, 404, { error: 'Report not found' });
        return;
      }
      const download = await client.download();
      const text = await streamToString(download.readableStreamBody);
      json(context, 200, JSON.parse(text));
      return;
    }

    if (method === 'POST') {
      let body = req.body;
      if (typeof body === 'string') {
        try { body = JSON.parse(body); } catch (_) { body = {}; }
      }
      body = body || {};
      if (!body.content || !String(body.content).trim()) {
        json(context, 400, { error: 'Report content is required' });
        return;
      }
      const report = {
        id: safeId(),
        title: String(body.title || 'Emme Report').slice(0, 200),
        client: String(body.client || '').slice(0, 200),
        provider: String(body.provider || '').slice(0, 200),
        reportType: String(body.reportType || '').slice(0, 100),
        content: String(body.content),
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      };
      const client = container.getBlockBlobClient(report.id + '.json');
      const payload = JSON.stringify(report);
      await client.upload(payload, Buffer.byteLength(payload), {
        blobHTTPHeaders: { blobContentType: 'application/json' }
      });
      json(context, 201, { configured: true, id: report.id, report });
      return;
    }

    if (method === 'DELETE' && id) {
      const client = container.getBlockBlobClient(id + '.json');
      await client.deleteIfExists();
      json(context, 200, { ok: true, id });
      return;
    }

    json(context, 405, { error: 'Method not allowed' });
  } catch (err) {
    json(context, 500, { error: (err && err.message) || 'Report library error' });
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
