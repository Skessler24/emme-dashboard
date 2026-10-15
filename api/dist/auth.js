"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.readPrincipal = readPrincipal;
exports.automationAuthorized = automationAuthorized;
exports.resolveIsAdmin = resolveIsAdmin;
exports.requireUser = requireUser;
exports.requireAdmin = requireAdmin;
const crypto_1 = require("crypto");
const config_1 = require("./config");
const graphClient_1 = require("./graphClient");
const http_1 = require("./http");
const adminCache = new Map();
function readPrincipal(req) {
    const directId = (0, http_1.header)(req, 'x-ms-client-principal-id');
    const directName = (0, http_1.header)(req, 'x-ms-client-principal-name');
    const encoded = (0, http_1.header)(req, 'x-ms-client-principal');
    let parsed = null;
    if (encoded) {
        try {
            parsed = JSON.parse(Buffer.from(encoded, 'base64').toString('utf8'));
        }
        catch (_) {
            parsed = null;
        }
    }
    const claims = (parsed && parsed.claims) || [];
    let email = '';
    let name = '';
    function consider(value, prefer) {
        const text = String(value || '').trim();
        if (!text)
            return;
        if (!name && text.indexOf('@') === -1)
            name = text;
        if (text.indexOf('@') === -1)
            return;
        if (prefer || !email)
            email = text.toLowerCase();
    }
    claims.forEach(function (claim) {
        const typ = String((claim && claim.typ) || '').toLowerCase();
        const val = String((claim && claim.val) || '').trim();
        if (typ.indexOf('name') !== -1 && typ.indexOf('upn') === -1 && val.indexOf('@') === -1)
            name = val;
        const prefer = typ.indexOf('email') !== -1 || typ.indexOf('preferred_username') !== -1 || typ.endsWith('/upn');
        if (prefer || val.indexOf('@') !== -1)
            consider(val, prefer);
    });
    consider(directName, directName.indexOf('@') !== -1);
    const userId = directId || (parsed && parsed.userId) || '';
    if (!email && !userId)
        return null;
    return {
        userId: String(userId || email),
        name: name || email,
        email: email,
    };
}
function automationAuthorized(req) {
    const expected = (0, config_1.getConfig)().automationKey;
    if (!expected)
        return false;
    const got = (0, http_1.header)(req, 'x-emme-automation-key');
    const a = Buffer.from(got);
    const b = Buffer.from(expected);
    if (a.length !== b.length || a.length === 0)
        return false;
    return (0, crypto_1.timingSafeEqual)(a, b);
}
async function resolveIsAdmin(email) {
    if (!email)
        return { isAdmin: false, check: 'signed-out' };
    // Temporary access until Entra admin consent lets the Emme Admin group check succeed.
    // Remove EMME_ADMIN_EMAILS after that group lookup works.
    if ((0, config_1.getConfig)().adminEmails.indexOf(email.trim().toLowerCase()) !== -1) {
        return { isAdmin: true, check: 'allowlist' };
    }
    if (!(0, config_1.graphConfigured)())
        return { isAdmin: false, check: 'not-configured' };
    const cached = adminCache.get(email);
    if (cached && cached.expires > Date.now()) {
        return { isAdmin: cached.value, check: 'cache' };
    }
    try {
        const isAdmin = await (0, graphClient_1.userIsAdmin)(email);
        adminCache.set(email, { value: isAdmin, expires: Date.now() + 5 * 60 * 1000 });
        return { isAdmin, check: 'graph' };
    }
    catch (err) {
        const message = err instanceof graphClient_1.GraphError ? err.message : 'Admin group lookup failed';
        return { isAdmin: false, check: message };
    }
}
async function requireUser(context, req) {
    const principal = readPrincipal(req);
    if (!principal || !principal.email) {
        (0, http_1.json)(context, 401, { error: 'Sign in required' });
        return null;
    }
    return principal;
}
async function requireAdmin(context, req) {
    if (automationAuthorized(req)) {
        return { userId: 'automation', name: 'Power Automate', email: '' };
    }
    const principal = await requireUser(context, req);
    if (!principal)
        return null;
    const result = await resolveIsAdmin(principal.email);
    if (!result.isAdmin) {
        (0, http_1.json)(context, 403, { error: 'Emme Admin access is required' });
        return null;
    }
    return principal;
}
