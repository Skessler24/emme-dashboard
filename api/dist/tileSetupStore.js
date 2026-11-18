"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ConflictError = void 0;
exports.readAnswers = readAnswers;
exports.listFiles = listFiles;
exports.saveAnswers = saveAnswers;
exports.addFile = addFile;
exports.readFile = readFile;
exports.removeFile = removeFile;
const storage_blob_1 = require("@azure/storage-blob");
const crypto_1 = require("crypto");
/**
 * Setup answers and uploaded files for the Onboarding, Referrals, Staff
 * Information, and Staff Expectations tiles. They live in the same storage
 * account as schedules and messages, under tile-setup/<tile>/, one folder per
 * tile, so the tiles never see each other's answers.
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
function answersName(tile) { return 'tile-setup/' + tile + '/answers.json'; }
function indexName(tile) { return 'tile-setup/' + tile + '/files.json'; }
function fileName(tile, id) { return 'tile-setup/' + tile + '/files/' + id; }
function container() {
    const conn = process.env.AZURE_STORAGE_CONNECTION_STRING || '';
    if (!conn)
        throw new Error('Setup storage is not configured.');
    return storage_blob_1.BlobServiceClient.fromConnectionString(conn).getContainerClient(CONTAINER);
}
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
/** Returns null when nobody has saved this tile's answers yet. */
async function readAnswers(tile) {
    const stored = await readJson(answersName(tile), null);
    if (!stored || typeof stored !== 'object' || !stored.updatedAt)
        return null;
    return {
        updatedAt: String(stored.updatedAt),
        updatedBy: String(stored.updatedBy || ''),
        answers: stored.answers && typeof stored.answers === 'object' ? stored.answers : {},
    };
}
async function listFiles(tile) {
    const rows = await readJson(indexName(tile), []);
    return Array.isArray(rows) ? rows : [];
}
/**
 * Saves the answers. When answers were already saved by someone, the caller
 * must pass the updatedAt it last saw, so two people never overwrite each other.
 */
function saveAnswers(tile, answers, baseUpdatedAt, by) {
    return locked(async function () {
        const current = await readAnswers(tile);
        if (current && current.updatedAt !== baseUpdatedAt) {
            throw new ConflictError(current.updatedBy, current.updatedAt);
        }
        const saved = { updatedAt: new Date().toISOString(), updatedBy: by, answers };
        await writeJson(answersName(tile), saved);
        return saved;
    });
}
function addFile(tile, meta, data) {
    return locked(async function () {
        const file = Object.assign({}, meta, {
            id: (0, crypto_1.randomBytes)(8).toString('hex'),
            size: data.length,
            uploadedAt: new Date().toISOString(),
        });
        const target = container();
        await target.createIfNotExists();
        await target.getBlockBlobClient(fileName(tile, file.id)).upload(data, data.length, {
            blobHTTPHeaders: { blobContentType: 'application/octet-stream' },
        });
        const files = await listFiles(tile);
        files.push(file);
        await writeJson(indexName(tile), files);
        return file;
    });
}
async function readFile(tile, id) {
    const files = await listFiles(tile);
    const file = files.find(function (row) { return row.id === id; });
    if (!file)
        return null;
    const blob = container().getBlockBlobClient(fileName(tile, file.id));
    if (!(await blob.exists()))
        return null;
    return { file, data: await blob.downloadToBuffer() };
}
function removeFile(tile, id) {
    return locked(async function () {
        const files = await listFiles(tile);
        const file = files.find(function (row) { return row.id === id; });
        if (!file)
            return false;
        await container().getBlockBlobClient(fileName(tile, file.id)).deleteIfExists();
        await writeJson(indexName(tile), files.filter(function (row) { return row.id !== id; }));
        return true;
    });
}
