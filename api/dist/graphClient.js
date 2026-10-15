"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.GraphError = void 0;
exports.graphIsConfigured = graphIsConfigured;
exports.graphGet = graphGet;
exports.graphGetText = graphGetText;
exports.graphSend = graphSend;
exports.getSharePointSiteId = getSharePointSiteId;
exports.getUserDisplayName = getUserDisplayName;
exports.userIsAdmin = userIsAdmin;
const config_1 = require("./config");
const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';
let tokenCache = { value: '', expires: 0 };
let siteIdCache = '';
class GraphError extends Error {
    status;
    constructor(status, message) {
        super(message);
        this.name = 'GraphError';
        this.status = status;
    }
}
exports.GraphError = GraphError;
function graphIsConfigured() {
    return (0, config_1.graphConfigured)();
}
async function appToken() {
    const cfg = (0, config_1.getConfig)();
    if (!cfg.graphClientSecret) {
        throw new GraphError(503, 'Microsoft Graph is not configured. Set GRAPH_CLIENT_SECRET.');
    }
    if (tokenCache.value && tokenCache.expires > Date.now() + 60000)
        return tokenCache.value;
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
    const data = await res.json().catch(function () { return {}; });
    if (!res.ok || !data.access_token) {
        throw new GraphError(res.status || 502, 'Microsoft app sign-in failed. Check the client secret and admin consent.');
    }
    tokenCache = {
        value: data.access_token,
        expires: Date.now() + (Number(data.expires_in) || 3600) * 1000,
    };
    return tokenCache.value;
}
function toUrl(path) {
    if (path.startsWith('https://'))
        return path;
    return GRAPH_ROOT + (path.startsWith('/') ? path : '/' + path);
}
async function graphGet(path, extraHeaders) {
    const token = await appToken();
    const res = await fetch(toUrl(path), {
        headers: Object.assign({ Authorization: 'Bearer ' + token, Accept: 'application/json' }, extraHeaders || {}),
    });
    if (res.status === 204)
        return null;
    const data = await res.json().catch(function () { return {}; });
    if (!res.ok) {
        const message = (data && data.error && (data.error.message || data.error.code)) || ('Graph ' + res.status);
        throw new GraphError(res.status, String(message).slice(0, 300));
    }
    return data;
}
async function graphGetText(path) {
    const token = await appToken();
    const res = await fetch(toUrl(path), {
        headers: { Authorization: 'Bearer ' + token },
    });
    if (!res.ok) {
        throw new GraphError(res.status, 'Graph file read failed (' + res.status + ')');
    }
    return res.text();
}
async function graphSend(path, options) {
    const token = await appToken();
    const headers = { Authorization: 'Bearer ' + token };
    let body;
    if (options.rawBody) {
        headers['Content-Type'] = options.contentType || 'application/octet-stream';
        const copy = new ArrayBuffer(options.rawBody.byteLength);
        new Uint8Array(copy).set(options.rawBody);
        body = copy;
    }
    else if (options.jsonBody !== undefined) {
        headers['Content-Type'] = 'application/json';
        body = JSON.stringify(options.jsonBody);
    }
    const res = await fetch(toUrl(path), { method: options.method, headers, body });
    if (res.status === 204)
        return null;
    const text = await res.text();
    let data = {};
    if (text) {
        try {
            data = JSON.parse(text);
        }
        catch (_) {
            data = {};
        }
    }
    if (!res.ok) {
        const message = (data && data.error && (data.error.message || data.error.code)) || ('Graph ' + res.status);
        throw new GraphError(res.status, String(message).slice(0, 300));
    }
    return data;
}
async function getSharePointSiteId() {
    if (siteIdCache)
        return siteIdCache;
    const cfg = (0, config_1.getConfig)();
    if (!cfg.sharePointHostname || !cfg.sharePointSitePath) {
        throw new GraphError(503, 'SharePoint is not configured. Set SHAREPOINT_HOSTNAME and SHAREPOINT_SITE_PATH.');
    }
    const sitePath = cfg.sharePointSitePath.startsWith('/') ? cfg.sharePointSitePath : '/' + cfg.sharePointSitePath;
    const data = await graphGet('/sites/' + cfg.sharePointHostname + ':' + sitePath);
    if (!data || !data.id)
        throw new GraphError(502, 'SharePoint site was not found.');
    siteIdCache = String(data.id);
    return siteIdCache;
}
const nameCache = new Map();
async function getUserDisplayName(email) {
    const key = email.trim().toLowerCase();
    if (!key)
        return '';
    const cached = nameCache.get(key);
    if (cached && cached.expires > Date.now())
        return cached.name;
    try {
        const user = await graphGet('/users/' + encodeURIComponent(key) + '?$select=displayName,mail');
        const name = String((user && user.displayName) || key);
        nameCache.set(key, { name, expires: Date.now() + 10 * 60 * 1000 });
        return name;
    }
    catch (_) {
        return key;
    }
}
/**
 * Returns true when the signed-in user is a member of the Emme Admin group.
 * Uses application permissions so the check does not depend on a delegated token.
 */
async function userIsAdmin(email) {
    const cfg = (0, config_1.getConfig)();
    const key = email.trim().toLowerCase();
    if (!key || !(0, config_1.graphConfigured)())
        return false;
    const wantedName = cfg.adminGroupName.toLowerCase();
    let next = '/users/' + encodeURIComponent(key) +
        '/transitiveMemberOf/microsoft.graph.group?$select=id,displayName&$top=200';
    for (let page = 0; page < 5 && next; page++) {
        const data = await graphGet(next);
        const groups = (data && data.value) || [];
        for (let i = 0; i < groups.length; i++) {
            const id = String(groups[i].id || '');
            const name = String(groups[i].displayName || '').toLowerCase();
            if ((cfg.adminGroupId && id === cfg.adminGroupId) || name === wantedName)
                return true;
        }
        next = data && data['@odata.nextLink'] ? String(data['@odata.nextLink']) : null;
    }
    return false;
}
