/**
 * Saves a file into the signed-in provider's own OneDrive.
 * Files stay in the emmeco.org Microsoft 365 tenant, in a folder named Emme Dashboard.
 */
const DEFAULT_CLIENT_ID = 'b1e3bd77-c545-4716-afdc-8b0ec14ddcdf';
const DEFAULT_TENANT_ID = '4da01280-a61e-44cc-b537-9b21b7f17c68';
const FOLDER = 'Emme Dashboard';
const MAX_BYTES = 4 * 1024 * 1024;

let tokenCache = { value: '', expires: 0 };

function emmeEmail(value) {
  const email = String(value || '').trim().toLowerCase();
  return /@emmeco\.org$/.test(email) ? email : '';
}

function safeFileName(value) {
  const name = String(value || '').trim().replace(/\\/g, '/').split('/').pop() || '';
  if (!/^[A-Za-z0-9][A-Za-z0-9 ._()\-]{0,120}$/.test(name)) return '';
  if (!/\.(csv|pdf|html|json|txt|xlsx)$/i.test(name)) return '';
  if (name.indexOf('..') !== -1) return '';
  return name;
}

async function appToken() {
  if (!process.env.AZURE_CLIENT_SECRET) {
    throw new Error('Add AZURE_CLIENT_SECRET so files can be saved to the provider’s OneDrive.');
  }
  if (tokenCache.value && tokenCache.expires > Date.now() + 60000) return tokenCache.value;
  const tenant = process.env.AZURE_TENANT_ID || DEFAULT_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID || DEFAULT_CLIENT_ID;
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: process.env.AZURE_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });
  const res = await fetch('https://login.microsoftonline.com/' + tenant + '/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString()
  });
  const data = await res.json().catch(function () { return {}; });
  if (!res.ok || !data.access_token) {
    throw new Error('Microsoft app sign-in failed. Check the client secret and admin consent.');
  }
  tokenCache = {
    value: data.access_token,
    expires: Date.now() + (Number(data.expires_in) || 3600) * 1000
  };
  return tokenCache.value;
}

async function graphError(res) {
  const data = await res.json().catch(function () { return {}; });
  const message = (data.error && (data.error.message || data.error.code)) || ('Graph ' + res.status);
  if (res.status === 401 || res.status === 403) {
    return message + ' Jenny needs to grant Files.ReadWrite.All so files save into each provider’s OneDrive.';
  }
  return message;
}

async function ensureFolder(token, email) {
  const headers = { Authorization: 'Bearer ' + token };
  const existing = await fetch(
    'https://graph.microsoft.com/v1.0/users/' + encodeURIComponent(email) + '/drive/root:/' + encodeURIComponent(FOLDER),
    { headers: headers }
  );
  if (existing.ok) return;
  if (existing.status !== 404) throw new Error(await graphError(existing));
  const created = await fetch(
    'https://graph.microsoft.com/v1.0/users/' + encodeURIComponent(email) + '/drive/root/children',
    {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        name: FOLDER,
        folder: {},
        '@microsoft.graph.conflictBehavior': 'fail'
      })
    }
  );
  if (created.ok || created.status === 409) return;
  throw new Error(await graphError(created));
}

async function saveProviderFile(email, fileName, bytes, contentType) {
  const provider = emmeEmail(email);
  if (!provider) throw new Error('Files can only be saved for an emmeco.org account.');
  const name = safeFileName(fileName);
  if (!name) throw new Error('That file name cannot be saved to OneDrive.');
  const payload = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes || []);
  if (!payload.length || payload.length > MAX_BYTES) {
    throw new Error('The file must be between 1 byte and 4 MB.');
  }
  const token = await appToken();
  await ensureFolder(token, provider);
  const path = '/users/' + encodeURIComponent(provider) + '/drive/root:/' + encodeURIComponent(FOLDER) + '/' + encodeURIComponent(name) + ':/content';
  const res = await fetch('https://graph.microsoft.com/v1.0' + path, {
    method: 'PUT',
    headers: {
      Authorization: 'Bearer ' + token,
      'Content-Type': contentType || 'application/octet-stream'
    },
    body: payload
  });
  const data = await res.json().catch(function () { return {}; });
  if (!res.ok) {
    const message = (data.error && (data.error.message || data.error.code)) || ('Graph ' + res.status);
    if (res.status === 401 || res.status === 403) {
      throw new Error(message + ' Jenny needs to grant Files.ReadWrite.All so files save into each provider’s OneDrive.');
    }
    throw new Error(message);
  }
  return {
    ok: true,
    email: provider,
    name: name,
    path: '/' + FOLDER + '/' + name,
    webUrl: data.webUrl || ''
  };
}

module.exports = { saveProviderFile, safeFileName, emmeEmail };
