"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.GraphError = void 0;
exports.graphIsConfigured = graphIsConfigured;
exports.resetGraphBreaker = resetGraphBreaker;
exports.graphGet = graphGet;
exports.graphGetBytes = graphGetBytes;
exports.graphGetText = graphGetText;
exports.graphSend = graphSend;
exports.getSharePointSiteId = getSharePointSiteId;
exports.knownSites = knownSites;
exports.candidateSiteIds = candidateSiteIds;
exports.forgetFolderSites = forgetFolderSites;
exports.siteHoldingFolder = siteHoldingFolder;
exports.getUserDisplayName = getUserDisplayName;
exports.userExists = userExists;
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
/**
 * Every call to Microsoft goes through here. A call that takes too long is cut off,
 * and after a call fails to connect, further calls fail at once for a short while.
 * Without this, one slow or unreachable Microsoft service made every schedule save
 * wait until the web host gave up, so nothing was saved.
 */
let graphDownUntil = 0;
function graphTimeoutMs() {
    return Number(process.env.GRAPH_TIMEOUT_MS) || 10000;
}
function graphPauseMs() {
    const set = process.env.GRAPH_PAUSE_MS;
    return set === undefined || set === '' ? 30000 : Number(set) || 0;
}
/** Forgets a recent connection failure. Used by tests. */
function resetGraphBreaker() {
    graphDownUntil = 0;
}
async function graphFetch(url, init) {
    if (Date.now() < graphDownUntil) {
        throw new GraphError(503, 'Microsoft 365 is not answering right now. It will be tried again shortly.');
    }
    try {
        return await fetch(url, Object.assign({}, init || {}, { signal: AbortSignal.timeout(graphTimeoutMs()) }));
    }
    catch (_) {
        graphDownUntil = Date.now() + graphPauseMs();
        throw new GraphError(503, 'Microsoft 365 did not answer in time.');
    }
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
    const res = await graphFetch('https://login.microsoftonline.com/' + cfg.graphTenantId + '/oauth2/v2.0/token', {
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
    const res = await graphFetch(toUrl(path), {
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
async function graphGetBytes(path) {
    const token = await appToken();
    const res = await graphFetch(toUrl(path), {
        headers: { Authorization: 'Bearer ' + token },
    });
    if (!res.ok) {
        throw new GraphError(res.status, 'Graph file read failed (' + res.status + ')');
    }
    return new Uint8Array(await res.arrayBuffer());
}
async function graphGetText(path) {
    const token = await appToken();
    const res = await graphFetch(toUrl(path), {
        headers: { Authorization: 'Bearer ' + token },
    });
    if (!res.ok) {
        throw new GraphError(res.status, 'Graph file read failed (' + res.status + ')');
    }
    return res.text();
}
async function graphSend(path, options) {
    const token = await appToken();
    const headers = Object.assign({ Authorization: 'Bearer ' + token }, options.headers || {});
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
    const res = await graphFetch(toUrl(path), { method: options.method, headers, body });
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
const LEGACY_PERSONAL_SITE = '/sites/emmeco031-my.sharepoint.com:/personal/jenriat_emmeco_org';
/** The shared Emme site from the Azure settings, then Jenny's old OneDrive, which is only used until the folders are moved. */
async function knownSites() {
    let shared = '';
    let personal = '';
    try {
        shared = await getSharePointSiteId();
    }
    catch (_) { /* not configured or not reachable */ }
    try {
        const found = await graphGet(LEGACY_PERSONAL_SITE + '?$select=id');
        personal = found && found.id ? String(found.id) : '';
    }
    catch (_) { /* the legacy library is optional */ }
    return { shared, personal };
}
async function candidateSiteIds() {
    const sites = await knownSites();
    const ids = [];
    if (sites.shared)
        ids.push(sites.shared);
    if (sites.personal && ids.indexOf(sites.personal) === -1)
        ids.push(sites.personal);
    return ids;
}
const SITE_MEMORY_MS = 5 * 60 * 1000;
const folderSites = new Map();
/** Forgets every remembered answer to "which site holds this folder". Used by tests. */
function forgetFolderSites() {
    folderSites.clear();
}
/**
 * Finds which site holds a fixed folder: the shared Emme site first, then the
 * old personal library. The answer is remembered for a few minutes only, so a
 * folder that is moved is found again without restarting anything. When the
 * folder is on neither site, the first site is returned but never remembered,
 * so the next request looks again.
 */
async function siteHoldingFolder(segments, notConfigured) {
    const key = segments.join('/');
    const hit = folderSites.get(key);
    if (hit && hit.expires > Date.now())
        return hit.siteId;
    const tries = await candidateSiteIds();
    if (!tries.length)
        throw new GraphError(503, notConfigured);
    const folder = segments.map(encodeURIComponent).join('/');
    for (let i = 0; i < tries.length; i++) {
        try {
            await graphGet('/sites/' + tries[i] + '/drive/root:/' + folder + ':/children?$select=name&$top=1');
            folderSites.set(key, { siteId: tries[i], expires: Date.now() + SITE_MEMORY_MS });
            return tries[i];
        }
        catch (err) {
            if (err instanceof GraphError && (err.status === 401 || err.status === 403 || err.status === 404))
                continue;
            throw err;
        }
    }
    return tries[0];
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
 * False only when Microsoft says there is no such account. If Microsoft cannot be asked (not set up, or a
 * temporary failure), the answer is true so a provider can still be added.
 */
async function userExists(email) {
    const key = email.trim().toLowerCase();
    if (!key)
        return false;
    if (!(0, config_1.graphConfigured)())
        return true;
    try {
        await graphGet('/users/' + encodeURIComponent(key) + '?$select=id');
        return true;
    }
    catch (err) {
        return !(err instanceof GraphError && err.status === 404);
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
