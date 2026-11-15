"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ConflictError = void 0;
exports.emptyAnswers = emptyAnswers;
exports.readAnswers = readAnswers;
exports.listFiles = listFiles;
exports.saveAnswers = saveAnswers;
exports.addFile = addFile;
exports.readFile = readFile;
exports.removeFile = removeFile;
const storage_blob_1 = require("@azure/storage-blob");
const crypto_1 = require("crypto");
/**
 * Setup answers and the files people upload with them, for every topic. They
 * live in the same storage account as schedules and messages. Each topic has
 * its own prefix (for example billing-intake/ or intake/onboarding/), so no
 * topic can read or overwrite another, and no SharePoint folder is needed.
 */
const CONTAINER = process.env.AZURE_STORAGE_CONTAINER || 'emme-reports';
class ConflictError extends Error {
    updatedBy;
    updatedAt;
    constructor(updatedBy, updatedAt) {
        super('Someone else saved changes to this setup after you opened it. Reload the page to see their changes, then make your edits again.');
        this.updatedBy = updatedBy;
        this.updatedAt = updatedAt;
    }
}
exports.ConflictError = ConflictError;
let writeChain = Promise.resolve();
function locked(work) {
    const run = writeChain.then(work, work);
    writeChain = run.then(function () { return undefined; }, function () { return undefined; });
    return run;
}
function container() {
    const conn = process.env.AZURE_STORAGE_CONNECTION_STRING || '';
    if (!conn)
        throw new Error('Setup storage is not configured.');
    return storage_blob_1.BlobServiceClient.fromConnectionString(conn).getContainerClient(CONTAINER);
}
function answersName(prefix) { return prefix + 'answers.json'; }
function indexName(prefix) { return prefix + 'files.json'; }
function fileName(prefix, id) { return prefix + 'files/' + id; }
async function readJson(name, fallback) {
    const blob = container().getBlockBlobClient(name);
    if (!(await blob.exists()))
        return fallback;
    try {
        return JSON.parse((await blob.downloadToBuffer()).toString('utf8'));
    }
    catch (_) {
        return fallback;
    }
}
async function writeJson(name, value) {
    const target = container();
    await target.createIfNotExists();
    const payload = JSON.stringify(value);
    await target.getBlockBlobClient(name).upload(payload, Buffer.byteLength(payload), {
        blobHTTPHeaders: { blobContentType: 'application/json' },
    });
}
function emptyAnswers() {
    return { updatedAt: '', updatedBy: '', general: {}, sources: [] };
}
async function readAnswers(prefix) {
    const stored = await readJson(answersName(prefix), null);
    if (!stored || typeof stored !== 'object')
        return emptyAnswers();
    return {
        updatedAt: String(stored.updatedAt || ''),
        updatedBy: String(stored.updatedBy || ''),
        general: stored.general || {},
        sources: Array.isArray(stored.sources) ? stored.sources : [],
    };
}
async function listFiles(prefix) {
    const rows = await readJson(indexName(prefix), []);
    return Array.isArray(rows) ? rows : [];
}
/**
 * Saves the answers. When the answers were already saved by someone, the caller
 * must pass the updatedAt it last saw, so two people never overwrite each other.
 * Files that belong to a group entry that was removed are deleted too.
 */
function saveAnswers(prefix, next, baseUpdatedAt, by) {
    return locked(async function () {
        const current = await readAnswers(prefix);
        if (current.updatedAt && current.updatedAt !== baseUpdatedAt) {
            throw new ConflictError(current.updatedBy, current.updatedAt);
        }
        const saved = {
            updatedAt: new Date().toISOString(),
            updatedBy: by,
            general: next.general,
            sources: next.sources,
        };
        await writeJson(answersName(prefix), saved);
        const keep = new Set(['general'].concat(saved.sources.map(function (s) { return s.id; })));
        const files = await listFiles(prefix);
        const orphans = files.filter(function (file) { return !keep.has(file.scope); });
        if (orphans.length) {
            const target = container();
            for (let i = 0; i < orphans.length; i++) {
                await target.getBlockBlobClient(fileName(prefix, orphans[i].id)).deleteIfExists();
            }
            await writeJson(indexName(prefix), files.filter(function (file) { return keep.has(file.scope); }));
        }
        return saved;
    });
}
function addFile(prefix, meta, data) {
    return locked(async function () {
        const file = Object.assign({}, meta, {
            id: (0, crypto_1.randomBytes)(8).toString('hex'),
            size: data.length,
            uploadedAt: new Date().toISOString(),
        });
        const target = container();
        await target.createIfNotExists();
        await target.getBlockBlobClient(fileName(prefix, file.id)).upload(data, data.length, {
            blobHTTPHeaders: { blobContentType: 'application/octet-stream' },
        });
        const files = await listFiles(prefix);
        files.push(file);
        await writeJson(indexName(prefix), files);
        return file;
    });
}
async function readFile(prefix, id) {
    const files = await listFiles(prefix);
    const file = files.find(function (row) { return row.id === id; });
    if (!file)
        return null;
    const blob = container().getBlockBlobClient(fileName(prefix, file.id));
    if (!(await blob.exists()))
        return null;
    return { file, data: await blob.downloadToBuffer() };
}
function removeFile(prefix, id) {
    return locked(async function () {
        const files = await listFiles(prefix);
        const file = files.find(function (row) { return row.id === id; });
        if (!file)
            return false;
        await container().getBlockBlobClient(fileName(prefix, file.id)).deleteIfExists();
        await writeJson(indexName(prefix), files.filter(function (row) { return row.id !== id; }));
        return true;
    });
}
