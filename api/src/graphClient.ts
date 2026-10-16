import { getConfig, graphConfigured } from './config';

const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';

type TokenCache = { value: string; expires: number };
let tokenCache: TokenCache = { value: '', expires: 0 };
let siteIdCache = '';

export class GraphError extends Error {
  status: number;
  constructor(status: number, message: string) {
    super(message);
    this.name = 'GraphError';
    this.status = status;
  }
}

export function graphIsConfigured(): boolean {
  return graphConfigured();
}

async function appToken(): Promise<string> {
  const cfg = getConfig();
  if (!cfg.graphClientSecret) {
    throw new GraphError(503, 'Microsoft Graph is not configured. Set GRAPH_CLIENT_SECRET.');
  }
  if (tokenCache.value && tokenCache.expires > Date.now() + 60000) return tokenCache.value;
  const body = new URLSearchParams({
    client_id: cfg.graphClientId,
    client_secret: cfg.graphClientSecret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });
  const res = await fetch('https://login.microsoftonline.com/' + cfg.graphTenantId + '/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString(),
  });
  const data = await res.json().catch(function () { return {}; }) as { access_token?: string; expires_in?: number };
  if (!res.ok || !data.access_token) {
    throw new GraphError(res.status || 502, 'Microsoft app sign-in failed. Check the client secret and admin consent.');
  }
  tokenCache = {
    value: data.access_token,
    expires: Date.now() + (Number(data.expires_in) || 3600) * 1000,
  };
  return tokenCache.value;
}

function toUrl(path: string): string {
  if (path.startsWith('https://')) return path;
  return GRAPH_ROOT + (path.startsWith('/') ? path : '/' + path);
}

export async function graphGet(path: string, extraHeaders?: Record<string, string>): Promise<any> {
  const token = await appToken();
  const res = await fetch(toUrl(path), {
    headers: Object.assign({ Authorization: 'Bearer ' + token, Accept: 'application/json' }, extraHeaders || {}),
  });
  if (res.status === 204) return null;
  const data = await res.json().catch(function () { return {}; });
  if (!res.ok) {
    const message = (data && data.error && (data.error.message || data.error.code)) || ('Graph ' + res.status);
    throw new GraphError(res.status, String(message).slice(0, 300));
  }
  return data;
}

export async function graphGetBytes(path: string): Promise<Uint8Array> {
  const token = await appToken();
  const res = await fetch(toUrl(path), {
    headers: { Authorization: 'Bearer ' + token },
  });
  if (!res.ok) {
    throw new GraphError(res.status, 'Graph file read failed (' + res.status + ')');
  }
  return new Uint8Array(await res.arrayBuffer());
}

export async function graphGetText(path: string): Promise<string> {
  const token = await appToken();
  const res = await fetch(toUrl(path), {
    headers: { Authorization: 'Bearer ' + token },
  });
  if (!res.ok) {
    throw new GraphError(res.status, 'Graph file read failed (' + res.status + ')');
  }
  return res.text();
}

export async function graphSend(path: string, options: {
  method: string;
  jsonBody?: unknown;
  rawBody?: Uint8Array;
  contentType?: string;
}): Promise<any> {
  const token = await appToken();
  const headers: Record<string, string> = { Authorization: 'Bearer ' + token };
  let body: BodyInit | undefined;
  if (options.rawBody) {
    headers['Content-Type'] = options.contentType || 'application/octet-stream';
    const copy = new ArrayBuffer(options.rawBody.byteLength);
    new Uint8Array(copy).set(options.rawBody);
    body = copy;
  } else if (options.jsonBody !== undefined) {
    headers['Content-Type'] = 'application/json';
    body = JSON.stringify(options.jsonBody);
  }
  const res = await fetch(toUrl(path), { method: options.method, headers, body });
  if (res.status === 204) return null;
  const text = await res.text();
  let data: any = {};
  if (text) {
    try { data = JSON.parse(text); } catch (_) { data = {}; }
  }
  if (!res.ok) {
    const message = (data && data.error && (data.error.message || data.error.code)) || ('Graph ' + res.status);
    throw new GraphError(res.status, String(message).slice(0, 300));
  }
  return data;
}

export async function getSharePointSiteId(): Promise<string> {
  if (siteIdCache) return siteIdCache;
  const cfg = getConfig();
  if (!cfg.sharePointHostname || !cfg.sharePointSitePath) {
    throw new GraphError(503, 'SharePoint is not configured. Set SHAREPOINT_HOSTNAME and SHAREPOINT_SITE_PATH.');
  }
  const sitePath = cfg.sharePointSitePath.startsWith('/') ? cfg.sharePointSitePath : '/' + cfg.sharePointSitePath;
  const data = await graphGet('/sites/' + cfg.sharePointHostname + ':' + sitePath);
  if (!data || !data.id) throw new GraphError(502, 'SharePoint site was not found.');
  siteIdCache = String(data.id);
  return siteIdCache;
}

const nameCache = new Map<string, { name: string; expires: number }>();

export async function getUserDisplayName(email: string): Promise<string> {
  const key = email.trim().toLowerCase();
  if (!key) return '';
  const cached = nameCache.get(key);
  if (cached && cached.expires > Date.now()) return cached.name;
  try {
    const user = await graphGet('/users/' + encodeURIComponent(key) + '?$select=displayName,mail');
    const name = String((user && user.displayName) || key);
    nameCache.set(key, { name, expires: Date.now() + 10 * 60 * 1000 });
    return name;
  } catch (_) {
    return key;
  }
}

/**
 * Returns true when the signed-in user is a member of the Emme Admin group.
 * Uses application permissions so the check does not depend on a delegated token.
 */
export async function userIsAdmin(email: string): Promise<boolean> {
  const cfg = getConfig();
  const key = email.trim().toLowerCase();
  if (!key || !graphConfigured()) return false;
  const wantedName = cfg.adminGroupName.toLowerCase();
  let next: string | null = '/users/' + encodeURIComponent(key) +
    '/transitiveMemberOf/microsoft.graph.group?$select=id,displayName&$top=200';
  for (let page = 0; page < 5 && next; page++) {
    const data = await graphGet(next);
    const groups = (data && data.value) || [];
    for (let i = 0; i < groups.length; i++) {
      const id = String(groups[i].id || '');
      const name = String(groups[i].displayName || '').toLowerCase();
      if ((cfg.adminGroupId && id === cfg.adminGroupId) || name === wantedName) return true;
    }
    next = data && data['@odata.nextLink'] ? String(data['@odata.nextLink']) : null;
  }
  return false;
}
