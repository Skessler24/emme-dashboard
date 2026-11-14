"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.MAX_FILES = exports.MAX_FILE_BYTES = void 0;
exports.safeFileName = safeFileName;
exports.handleBillingIntake = handleBillingIntake;
const auth_1 = require("./auth");
const http_1 = require("./http");
const billingQuestions_1 = require("./billingQuestions");
const billingIntakeStore_1 = require("./billingIntakeStore");
const billingPrompt_1 = require("./billingPrompt");
const crypto_1 = require("crypto");
exports.MAX_FILE_BYTES = 6 * 1024 * 1024;
exports.MAX_FILES = 100;
const MAX_SOURCES = 40;
const MAX_TEXT = 6000;
const FILE_TYPES = {
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12',
    xls: 'application/vnd.ms-excel',
    csv: 'text/csv',
    pdf: 'application/pdf',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    doc: 'application/msword',
    txt: 'text/plain',
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
};
/** Billing setup holds sensitive notes, so only a signed-in human administrator may use it (the automation key is not enough). */
async function requireHumanAdmin(context, req) {
    const principal = (0, auth_1.readPrincipal)(req);
    if (!principal || !principal.email) {
        (0, http_1.json)(context, 401, { error: 'Sign in required' });
        return null;
    }
    if (!(0, auth_1.emailAllowed)(principal.email)) {
        (0, http_1.json)(context, 403, { error: auth_1.NOT_EMME_MESSAGE });
        return null;
    }
    const admin = await (0, auth_1.resolveIsAdmin)(principal.email);
    if (!admin.isAdmin) {
        (0, http_1.json)(context, 403, { error: 'Emme Admin access is required' });
        return null;
    }
    return { email: principal.email, name: principal.name };
}
function cleanAnswers(raw, questions) {
    const out = {};
    const source = raw && typeof raw === 'object' ? raw : {};
    questions.forEach(function (question) {
        const given = source[question.id];
        if (!given || typeof given !== 'object')
            return;
        const valid = (question.options || []).map(function (option) { return option.id; });
        let picked = Array.isArray(given.choice) ? given.choice.map(String) : [];
        picked = picked.filter(function (id, index) { return valid.indexOf(id) !== -1 && picked.indexOf(id) === index; });
        if (question.kind === 'one')
            picked = picked.slice(0, 1);
        if (question.kind === 'text')
            picked = [];
        const answer = { choice: picked, text: (0, http_1.asText)(given.text, MAX_TEXT), unsure: given.unsure === true };
        if (answer.choice.length || answer.text || answer.unsure)
            out[question.id] = answer;
    });
    return out;
}
function cleanSources(raw) {
    if (!Array.isArray(raw))
        return null;
    const seen = new Set();
    const out = [];
    for (let i = 0; i < raw.length && out.length < MAX_SOURCES; i++) {
        const item = raw[i];
        if (!item || typeof item !== 'object')
            continue;
        const kind = String(item.kind || '');
        if (!billingQuestions_1.SOURCE_KINDS.some(function (k) { return k.id === kind; }))
            continue;
        let id = String(item.id || '');
        if (!/^[a-z0-9-]{3,40}$/.test(id) || id === 'general' || seen.has(id))
            id = 'src-' + (0, crypto_1.randomBytes)(5).toString('hex');
        seen.add(id);
        out.push({
            id,
            kind,
            name: (0, http_1.asText)(item.name, 120) || 'Unnamed funding source',
            answers: cleanAnswers(item.answers, (0, billingQuestions_1.questionsForKind)(kind)),
        });
    }
    return out;
}
function safeFileName(name) {
    const base = String(name || '').split(/[\\/]/).pop() || '';
    const cleaned = base.replace(/[^A-Za-z0-9._ ()-]/g, '_').replace(/\.{2,}/g, '.').trim().slice(0, 120);
    return cleaned || 'file';
}
function extensionOf(name) {
    const dot = name.lastIndexOf('.');
    return dot === -1 ? '' : name.slice(dot + 1).toLowerCase();
}
function questionBank() {
    return {
        generalSections: billingQuestions_1.GENERAL_SECTIONS,
        generalQuestions: billingQuestions_1.GENERAL_QUESTIONS,
        sourceSections: billingQuestions_1.SOURCE_SECTIONS,
        sourceQuestions: billingQuestions_1.SOURCE_QUESTIONS,
        sourceKinds: billingQuestions_1.SOURCE_KINDS,
        fileKinds: billingQuestions_1.FILE_KINDS,
        connectionGuide: billingQuestions_1.CONNECTION_GUIDE,
        limits: { maxFileBytes: exports.MAX_FILE_BYTES, fileTypes: Object.keys(FILE_TYPES) },
    };
}
async function readState(context) {
    const [answers, files] = await Promise.all([(0, billingIntakeStore_1.readAnswers)(), (0, billingIntakeStore_1.listFiles)()]);
    (0, http_1.json)(context, 200, Object.assign({ answers, files }, questionBank()));
}
async function save(context, req, who) {
    const body = (0, http_1.readBody)(req);
    const sources = cleanSources(body.sources);
    if (!sources) {
        (0, http_1.json)(context, 400, { error: 'Send the list of funding sources' });
        return;
    }
    const general = cleanAnswers(body.general, billingQuestions_1.GENERAL_QUESTIONS);
    try {
        const saved = await (0, billingIntakeStore_1.saveAnswers)({ general, sources }, (0, http_1.asText)(body.baseUpdatedAt, 40), who.email);
        (0, http_1.json)(context, 200, { ok: true, updatedAt: saved.updatedAt, updatedBy: saved.updatedBy, sources: saved.sources.map(function (s) { return { id: s.id, name: s.name }; }) });
    }
    catch (err) {
        if (err instanceof billingIntakeStore_1.ConflictError) {
            (0, http_1.json)(context, 409, { error: err.message, updatedBy: err.updatedBy, updatedAt: err.updatedAt });
            return;
        }
        throw err;
    }
}
async function upload(context, req, who) {
    const body = (0, http_1.readBody)(req);
    if (body.confirmedNoClientInfo !== true) {
        (0, http_1.json)(context, 400, { error: 'Confirm that the file has no real client information in it. Replace names, birth dates, and member numbers with fake ones first.' });
        return;
    }
    const name = safeFileName((0, http_1.asText)(body.name, 300));
    const ext = extensionOf(name);
    if (!FILE_TYPES[ext]) {
        (0, http_1.json)(context, 400, { error: 'That kind of file is not allowed. Use one of: ' + Object.keys(FILE_TYPES).join(', ') + '.' });
        return;
    }
    const kind = (0, http_1.asText)(body.kind, 40);
    if (!billingQuestions_1.FILE_KINDS.some(function (item) { return item.id === kind; })) {
        (0, http_1.json)(context, 400, { error: 'Choose what kind of file this is' });
        return;
    }
    const encoded = typeof body.contentBase64 === 'string' ? body.contentBase64.replace(/\s+/g, '') : '';
    if (!encoded || encoded.length > Math.ceil(exports.MAX_FILE_BYTES * 4 / 3) + 8 || !/^[A-Za-z0-9+/]+={0,2}$/.test(encoded)) {
        (0, http_1.json)(context, 400, { error: 'The file is missing, damaged, or larger than ' + Math.round(exports.MAX_FILE_BYTES / 1024 / 1024) + ' megabytes' });
        return;
    }
    const data = Buffer.from(encoded, 'base64');
    if (!data.length || data.length > exports.MAX_FILE_BYTES) {
        (0, http_1.json)(context, 400, { error: 'The file is empty or larger than ' + Math.round(exports.MAX_FILE_BYTES / 1024 / 1024) + ' megabytes' });
        return;
    }
    const scope = (0, http_1.asText)(body.scope, 40);
    if (scope !== 'general') {
        const answers = await (0, billingIntakeStore_1.readAnswers)();
        if (!answers.sources.some(function (source) { return source.id === scope; })) {
            (0, http_1.json)(context, 400, { error: 'Save the funding source first, then upload its files' });
            return;
        }
    }
    const existing = await (0, billingIntakeStore_1.listFiles)();
    if (existing.length >= exports.MAX_FILES) {
        (0, http_1.json)(context, 400, { error: 'The file limit was reached. Remove files you no longer need.' });
        return;
    }
    const file = await (0, billingIntakeStore_1.addFile)({
        scope,
        kind,
        note: (0, http_1.asText)(body.note, 300),
        name,
        contentType: FILE_TYPES[ext],
        uploadedBy: who.email,
    }, data);
    (0, http_1.json)(context, 200, { ok: true, file });
}
function contentDisposition(name) {
    return 'attachment; filename="' + name.replace(/"/g, '') + '"; filename*=UTF-8\'\'' + encodeURIComponent(name);
}
async function download(context, id) {
    const found = await (0, billingIntakeStore_1.readFile)(id);
    if (!found) {
        (0, http_1.json)(context, 404, { error: 'File not found' });
        return;
    }
    context.res = {
        status: 200,
        headers: {
            // Always a download, never shown in the browser, so a hostile file cannot run on the dashboard.
            'Content-Type': 'application/octet-stream',
            'Content-Disposition': contentDisposition(found.file.name),
            'X-Content-Type-Options': 'nosniff',
            'Cache-Control': 'no-store',
        },
        body: found.data,
        isRaw: true,
    };
}
async function prompt(context) {
    const [answers, files] = await Promise.all([(0, billingIntakeStore_1.readAnswers)(), (0, billingIntakeStore_1.listFiles)()]);
    const known = new Set(['general'].concat(answers.sources.map(function (s) { return s.id; })));
    (0, http_1.json)(context, 200, {
        prompt: (0, billingPrompt_1.buildPrompt)(answers, files.filter(function (file) { return known.has(file.scope); }), new Date()),
        assessments: answers.sources.map(billingPrompt_1.assessSource),
        missing: (0, billingPrompt_1.missingItems)(answers),
    });
}
async function handleBillingIntake(context, req) {
    if ((0, http_1.methodOf)(req) === 'OPTIONS') {
        context.res = { status: 204, headers: { 'Content-Type': 'application/json' } };
        return;
    }
    const who = await requireHumanAdmin(context, req);
    if (!who)
        return;
    try {
        const verb = (0, http_1.methodOf)(req);
        const parts = (0, http_1.restParts)(req);
        if (!parts.length && verb === 'GET')
            return await readState(context);
        if (!parts.length && verb === 'PUT')
            return await save(context, req, who);
        if (parts[0] === 'prompt' && verb === 'GET')
            return await prompt(context);
        if (parts[0] === 'files' && !parts[1] && verb === 'POST')
            return await upload(context, req, who);
        if (parts[0] === 'files' && parts[1] && verb === 'GET')
            return await download(context, parts[1]);
        if (parts[0] === 'files' && parts[1] && verb === 'DELETE') {
            const removed = await (0, billingIntakeStore_1.removeFile)(parts[1]);
            (0, http_1.json)(context, removed ? 200 : 404, removed ? { ok: true } : { error: 'File not found' });
            return;
        }
        (0, http_1.json)(context, 404, { error: 'Not found' });
    }
    catch (err) {
        // Do not echo storage details back to the browser.
        const message = err instanceof Error && /not configured/i.test(err.message) ? err.message : 'Billing setup could not be saved or read. Try again in a moment.';
        (0, http_1.json)(context, 500, { error: message });
    }
}
