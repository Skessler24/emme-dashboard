const Module = require('module');
const { Readable } = require('stream');

process.env.AZURE_STORAGE_CONNECTION_STRING = 'fake';
process.env.GRAPH_CLIENT_SECRET = 'fake';
process.env.SCHEDULE_PULL_THROTTLE_MS = '0';

const blobs = new Map();
class FakeBlob {
  constructor(name) { this.name = name; }
  async exists() { return blobs.has(this.name); }
  async getProperties() { return { metadata: blobs.get(this.name).meta }; }
  async download() { return { readableStreamBody: Readable.from([blobs.get(this.name).data]) }; }
  async upload(data, length, options) { blobs.set(this.name, { data, meta: options && options.metadata }); }
}
const fakeStorage = {
  BlobServiceClient: {
    fromConnectionString: () => ({
      getContainerClient: () => ({
        createIfNotExists: async () => {},
        getBlockBlobClient: (name) => new FakeBlob(name)
      })
    })
  }
};
const originalLoad = Module._load;
Module._load = function (request, ...rest) {
  if (request === '@azure/storage-blob') return fakeStorage;
  return originalLoad.call(this, request, ...rest);
};

const ROOT = 'EMME/LEADERSHIP TEAM/EMME Processes/Referral Process/Provider Schedules';
const FILE_PATH = ROOT + '/ndelgado/Natalie Delgado Schedule.xlsx';
const drive = { bytes: null, version: 0, beforePut: null, putCount: 0, latencyMs: 0, failPuts: 0 };

function etag() { return '"{' + drive.version + '},1"'; }
function jsonResponse(status, body) {
  return new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
}

global.fetch = async function (url, options) {
  if (drive.latencyMs) await new Promise(function (resolve) { setTimeout(resolve, drive.latencyMs); });
  const method = (options && options.method) || 'GET';
  const text = decodeURIComponent(String(url));
  if (text.includes('/oauth2/v2.0/token')) return jsonResponse(200, { access_token: 't', expires_in: 3600 });
  if (text.includes('/users/')) return jsonResponse(200, { displayName: 'Natalie Delgado' });
  if (text.includes('/personal/jenriat_emmeco_org')) return jsonResponse(200, { id: 'site1' });
  if (!text.includes('/sites/site1/drive/root:/')) return jsonResponse(404, { error: { code: 'itemNotFound', message: 'no' } });
  const after = text.split('/drive/root:/')[1];
  const pathOnly = after.split(':')[0].split('?')[0];
  if (after.includes(':/children')) {
    if (pathOnly === ROOT) return jsonResponse(200, { value: [{ name: 'ndelgado', folder: {} }] });
    if (pathOnly === ROOT + '/ndelgado') return jsonResponse(200, { value: [{ name: 'Natalie Delgado Schedule.xlsx', file: {} }] });
    return jsonResponse(404, { error: { code: 'itemNotFound', message: 'no' } });
  }
  if (pathOnly === ROOT + '/ndelgado/emme-provider.json') return jsonResponse(404, { error: { code: 'itemNotFound', message: 'no' } });
  if (pathOnly !== FILE_PATH) return jsonResponse(404, { error: { code: 'itemNotFound', message: 'no' } });
  if (after.includes(':/content')) {
    if (method === 'PUT') {
      if (drive.failPuts > 0) { drive.failPuts -= 1; return jsonResponse(423, { error: { code: 'resourceLocked', message: 'The file is locked' } }); }
      if (drive.beforePut) { const hook = drive.beforePut; drive.beforePut = null; await hook(); }
      const match = options.headers['If-Match'];
      if (match && match !== etag()) return jsonResponse(412, { error: { code: 'resourceModified', message: 'etag' } });
      drive.bytes = Buffer.from(options.body);
      drive.version += 1;
      drive.putCount += 1;
      return jsonResponse(200, { eTag: etag(), lastModifiedDateTime: new Date().toISOString() });
    }
    return new Response(drive.bytes, { status: 200 });
  }
  return jsonResponse(200, { eTag: etag(), lastModifiedDateTime: new Date().toISOString() });
};


module.exports = { drive, blobs, FILE_PATH };
