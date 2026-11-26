"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ExpectationsStoreError = void 0;
exports.emptyBoard = emptyBoard;
exports.newId = newId;
exports.updateBoard = updateBoard;
exports.loadBoard = loadBoard;
const storage_blob_1 = require("@azure/storage-blob");
const crypto_1 = require("crypto");
/**
 * Live mileage board, travel approvals, and reminder checks.
 * Stored under expectations/ so schedules and other tiles are left alone.
 * Note text is never written here.
 */
const CONTAINER = process.env.AZURE_STORAGE_CONTAINER || 'emme-reports';
const BLOB = 'expectations/staff-board.json';
class ExpectationsStoreError extends Error {
    status;
    constructor(message, status) {
        super(message);
        this.status = status;
    }
}
exports.ExpectationsStoreError = ExpectationsStoreError;
let writeChain = Promise.resolve();
function locked(work) {
    const run = writeChain.then(work, work);
    writeChain = run.then(function () { return undefined; }, function () { return undefined; });
    return run;
}
function container() {
    const conn = process.env.AZURE_STORAGE_CONNECTION_STRING || '';
    if (!conn)
        throw new ExpectationsStoreError('Mileage storage is not configured.', 503);
    return storage_blob_1.BlobServiceClient.fromConnectionString(conn).getContainerClient(CONTAINER);
}
function emptyBoard() {
    return { homes: {}, trips: [], hours: {}, requests: [], emailChecks: {}, frequency: {}, followUps: [], outreach: {} };
}
function newId(prefix) {
    return prefix + '-' + (0, crypto_1.randomBytes)(4).toString('hex');
}
async function readBoard() {
    const blob = container().getBlockBlobClient(BLOB);
    if (!(await blob.exists()))
        return emptyBoard();
    try {
        const parsed = JSON.parse((await blob.downloadToBuffer()).toString('utf8'));
        const blank = emptyBoard();
        return {
            homes: parsed.homes || blank.homes,
            trips: Array.isArray(parsed.trips) ? parsed.trips : [],
            hours: parsed.hours || blank.hours,
            requests: Array.isArray(parsed.requests) ? parsed.requests : [],
            emailChecks: parsed.emailChecks || blank.emailChecks,
            frequency: parsed.frequency || blank.frequency,
            followUps: Array.isArray(parsed.followUps) ? parsed.followUps : [],
            outreach: parsed.outreach || blank.outreach,
        };
    }
    catch (_) {
        return emptyBoard();
    }
}
async function writeBoard(state) {
    const target = container();
    await target.createIfNotExists();
    const blob = target.getBlockBlobClient(BLOB);
    const body = JSON.stringify(state);
    await blob.upload(body, Buffer.byteLength(body), {
        blobHTTPHeaders: { blobContentType: 'application/json' },
        metadata: { kind: 'expectations' },
    });
}
function updateBoard(change) {
    return locked(async function () {
        const state = await readBoard();
        change(state);
        await writeBoard(state);
        return state;
    });
}
function loadBoard() {
    return readBoard();
}
