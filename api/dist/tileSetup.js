"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.MAX_TILE_FILES = void 0;
exports.handleTileSetup = handleTileSetup;
const http_1 = require("./http");
const billingIntake_1 = require("./billingIntake");
const tileSetupQuestions_1 = require("./tileSetupQuestions");
const tileSetupStore_1 = require("./tileSetupStore");
const tilePrompt_1 = require("./tilePrompt");
exports.MAX_TILE_FILES = 40;
/** What the browser needs to draw the form. The first-draft answers are sent as answers, not as part of the questions. */
function bank(tile) {
    return {
        tile: { id: tile.id, title: tile.title },
        sections: tile.sections,
        questions: tile.questions.map(function (question) {
            const copy = Object.assign({}, question);
            delete copy.seed;
            return copy;
        }),
        fileKinds: tile.fileKinds,
        limits: { maxFileBytes: billingIntake_1.MAX_FILE_BYTES, fileTypes: Object.keys(billingIntake_1.FILE_TYPES) },
    };
}
async function readState(context, tile) {
    const [saved, files] = await Promise.all([(0, tileSetupStore_1.readAnswers)(tile.id), (0, tileSetupStore_1.listFiles)(tile.id)]);
    (0, http_1.json)(context, 200, Object.assign({
        answers: saved
            ? { updatedAt: saved.updatedAt, updatedBy: saved.updatedBy, answers: saved.answers, firstDraft: false }
            : { updatedAt: '', updatedBy: '', answers: (0, tilePrompt_1.seededAnswers)(tile), firstDraft: true },
        files,
    }, bank(tile)));
}
async function save(context, req, tile, who) {
    const body = (0, http_1.readBody)(req);
    const answers = (0, billingIntake_1.cleanAnswers)(body.answers, tile.questions);
    try {
        const saved = await (0, tileSetupStore_1.saveAnswers)(tile.id, answers, (0, http_1.asText)(body.baseUpdatedAt, 40), who.email);
        (0, http_1.json)(context, 200, { ok: true, updatedAt: saved.updatedAt, updatedBy: saved.updatedBy });
    }
    catch (err) {
        if (err instanceof tileSetupStore_1.ConflictError) {
            (0, http_1.json)(context, 409, { error: err.message, updatedBy: err.updatedBy, updatedAt: err.updatedAt });
            return;
        }
        throw err;
    }
}
async function upload(context, req, tile, who) {
    const body = (0, http_1.readBody)(req);
    if (body.confirmedNoClientInfo !== true) {
        (0, http_1.json)(context, 400, { error: 'Confirm that the file has no real client information in it. Replace names, birth dates, and member numbers with fake ones first.' });
        return;
    }
    const name = (0, billingIntake_1.safeFileName)((0, http_1.asText)(body.name, 300));
    const ext = (0, billingIntake_1.extensionOf)(name);
    if (!billingIntake_1.FILE_TYPES[ext]) {
        (0, http_1.json)(context, 400, { error: 'That kind of file is not allowed. Use one of: ' + Object.keys(billingIntake_1.FILE_TYPES).join(', ') + '.' });
        return;
    }
    const kind = (0, http_1.asText)(body.kind, 40);
    if (!tile.fileKinds.some(function (item) { return item.id === kind; })) {
        (0, http_1.json)(context, 400, { error: 'Choose what kind of file this is' });
        return;
    }
    const encoded = typeof body.contentBase64 === 'string' ? body.contentBase64.replace(/\s+/g, '') : '';
    if (!encoded || encoded.length > Math.ceil(billingIntake_1.MAX_FILE_BYTES * 4 / 3) + 8 || !/^[A-Za-z0-9+/]+={0,2}$/.test(encoded)) {
        (0, http_1.json)(context, 400, { error: 'The file is missing, damaged, or larger than ' + Math.round(billingIntake_1.MAX_FILE_BYTES / 1024 / 1024) + ' megabytes' });
        return;
    }
    const data = Buffer.from(encoded, 'base64');
    if (!data.length || data.length > billingIntake_1.MAX_FILE_BYTES) {
        (0, http_1.json)(context, 400, { error: 'The file is empty or larger than ' + Math.round(billingIntake_1.MAX_FILE_BYTES / 1024 / 1024) + ' megabytes' });
        return;
    }
    if ((await (0, tileSetupStore_1.listFiles)(tile.id)).length >= exports.MAX_TILE_FILES) {
        (0, http_1.json)(context, 400, { error: 'The file limit was reached. Remove files you no longer need.' });
        return;
    }
    const file = await (0, tileSetupStore_1.addFile)(tile.id, {
        kind,
        note: (0, http_1.asText)(body.note, 300),
        name,
        contentType: billingIntake_1.FILE_TYPES[ext],
        uploadedBy: who.email,
    }, data);
    (0, http_1.json)(context, 200, { ok: true, file });
}
async function download(context, tile, id) {
    const found = await (0, tileSetupStore_1.readFile)(tile.id, id);
    if (!found) {
        (0, http_1.json)(context, 404, { error: 'File not found' });
        return;
    }
    context.res = {
        status: 200,
        headers: {
            // Always a download, never shown in the browser, so a hostile file cannot run on the dashboard.
            'Content-Type': 'application/octet-stream',
            'Content-Disposition': (0, billingIntake_1.contentDisposition)(found.file.name),
            'X-Content-Type-Options': 'nosniff',
            'Cache-Control': 'no-store',
        },
        body: found.data,
        isRaw: true,
    };
}
async function prompt(context, tile) {
    const [saved, files] = await Promise.all([(0, tileSetupStore_1.readAnswers)(tile.id), (0, tileSetupStore_1.listFiles)(tile.id)]);
    const answers = saved ? saved.answers : (0, tilePrompt_1.seededAnswers)(tile);
    (0, http_1.json)(context, 200, {
        prompt: (0, tilePrompt_1.buildTilePrompt)(tile, saved, files, new Date()),
        missing: (0, tilePrompt_1.missingItems)(tile, answers),
    });
}
async function handleTileSetup(context, req) {
    if ((0, http_1.methodOf)(req) === 'OPTIONS') {
        context.res = { status: 204, headers: { 'Content-Type': 'application/json' } };
        return;
    }
    const who = await (0, billingIntake_1.requireHumanAdmin)(context, req);
    if (!who)
        return;
    try {
        const verb = (0, http_1.methodOf)(req);
        const parts = (0, http_1.restParts)(req);
        if (!parts.length && verb === 'GET') {
            (0, http_1.json)(context, 200, { tiles: tileSetupQuestions_1.TILE_SETUPS.map(function (tile) { return { id: tile.id, title: tile.title }; }) });
            return;
        }
        const tile = (0, tileSetupQuestions_1.tileSetup)(parts[0]);
        if (!tile) {
            (0, http_1.json)(context, 404, { error: 'Not found' });
            return;
        }
        const rest = parts.slice(1);
        if (!rest.length && verb === 'GET')
            return await readState(context, tile);
        if (!rest.length && verb === 'PUT')
            return await save(context, req, tile, who);
        if (rest[0] === 'prompt' && verb === 'GET')
            return await prompt(context, tile);
        if (rest[0] === 'files' && !rest[1] && verb === 'POST')
            return await upload(context, req, tile, who);
        if (rest[0] === 'files' && rest[1] && verb === 'GET')
            return await download(context, tile, rest[1]);
        if (rest[0] === 'files' && rest[1] && verb === 'DELETE') {
            const removed = await (0, tileSetupStore_1.removeFile)(tile.id, rest[1]);
            (0, http_1.json)(context, removed ? 200 : 404, removed ? { ok: true } : { error: 'File not found' });
            return;
        }
        (0, http_1.json)(context, 404, { error: 'Not found' });
    }
    catch (err) {
        // Do not echo storage details back to the browser.
        const message = err instanceof Error && /not configured/i.test(err.message) ? err.message : 'The setup could not be saved or read. Try again in a moment.';
        (0, http_1.json)(context, 500, { error: message });
    }
}
