"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.json = json;
exports.header = header;
exports.methodOf = methodOf;
exports.queryOf = queryOf;
exports.restParts = restParts;
exports.readBody = readBody;
exports.asText = asText;
function json(context, status, body) {
    context.res = {
        status,
        headers: {
            'Content-Type': 'application/json',
            'Cache-Control': 'no-store',
        },
        body,
    };
}
function header(req, name) {
    const headers = (req && req.headers) || {};
    const want = name.toLowerCase();
    if (headers[want])
        return String(headers[want]);
    const key = Object.keys(headers).find(function (k) { return k.toLowerCase() === want; });
    return key ? String(headers[key]) : '';
}
function methodOf(req) {
    return String(req.method || 'GET').toUpperCase();
}
function queryOf(req, name) {
    const query = (req && req.query) || {};
    const want = name.toLowerCase();
    if (query[name] != null)
        return String(query[name]);
    const key = Object.keys(query).find(function (k) { return k.toLowerCase() === want; });
    return key && query[key] != null ? String(query[key]) : '';
}
function restParts(req) {
    const raw = String((req.params && (req.params.rest || req.params.id)) || '').replace(/^\/+|\/+$/g, '');
    if (!raw)
        return [];
    return raw.split('/').filter(Boolean).map(function (part) {
        try {
            return decodeURIComponent(part);
        }
        catch (_) {
            return part;
        }
    });
}
function readBody(req) {
    let body = req.body;
    if (typeof body === 'string') {
        try {
            body = JSON.parse(body);
        }
        catch (_) {
            body = {};
        }
    }
    if (!body || typeof body !== 'object' || Array.isArray(body))
        return {};
    return body;
}
function asText(value, max = 4000) {
    return String(value == null ? '' : value).trim().slice(0, max);
}
