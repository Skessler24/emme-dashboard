"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.StoreNotReadyError = exports.MapConflictError = void 0;
exports.mapBlobName = mapBlobName;
exports.readMap = readMap;
exports.saveMap = saveMap;
exports.readScheduleWorkspace = readScheduleWorkspace;
exports.listProviderEmails = listProviderEmails;
exports.listAppointments = listAppointments;
exports.listAllAppointments = listAllAppointments;
exports.addAppointment = addAppointment;
exports.updateAppointment = updateAppointment;
const storage_blob_1 = require("@azure/storage-blob");
const crypto_1 = require("crypto");
const reflectiveMap_1 = require("./reflectiveMap");
/**
 * Live growth maps and reflective supervision appointments.
 * Stored beside schedules, under reflective-supervision/, so other tiles never read them.
 * The words typed into a map are not written to the application log.
 */
const CONTAINER = process.env.AZURE_STORAGE_CONTAINER || 'emme-reports';
class MapConflictError extends Error {
    updatedBy;
    updatedAt;
    constructor(updatedBy, updatedAt) {
        super('Someone else saved this growth map after you opened it. Reload the page, then make your edits again.');
        this.updatedBy = updatedBy;
        this.updatedAt = updatedAt;
    }
}
exports.MapConflictError = MapConflictError;
class StoreNotReadyError extends Error {
    status;
    constructor() {
        super('Reflective supervision storage is not configured.');
        this.status = 503;
    }
}
exports.StoreNotReadyError = StoreNotReadyError;
let writeChain = Promise.resolve();
function locked(work) {
    const run = writeChain.then(work, work);
    writeChain = run.then(function () { return undefined; }, function () { return undefined; });
    return run;
}
function container() {
    const conn = process.env.AZURE_STORAGE_CONNECTION_STRING || '';
    if (!conn)
        throw new StoreNotReadyError();
    return storage_blob_1.BlobServiceClient.fromConnectionString(conn).getContainerClient(CONTAINER);
}
function mapBlobName(email) {
    const key = email.trim().toLowerCase().replace(/[^a-z0-9._-]/g, '_').slice(0, 120) || 'user';
    return 'reflective-supervision/maps/' + key + '.json';
}
function scheduleBlobName(email) {
    const key = email.trim().toLowerCase().replace(/[^a-z0-9._-]/g, '_').slice(0, 120) || 'user';
    return 'schedules/by-email/' + key + '.json';
}
const APPOINTMENTS = 'reflective-supervision/appointments.json';
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
async function writeJson(name, value, owner) {
    const target = container();
    await target.createIfNotExists();
    const payload = JSON.stringify(value);
    await target.getBlockBlobClient(name).upload(payload, Buffer.byteLength(payload), {
        blobHTTPHeaders: { blobContentType: 'application/json' },
        metadata: owner ? { owner: owner.slice(0, 200) } : undefined,
    });
}
async function readMap(email, name) {
    const stored = await readJson(mapBlobName(email), null);
    if (!stored || !stored.topics)
        return (0, reflectiveMap_1.blankMap)(email, name);
    return (0, reflectiveMap_1.cleanMap)(email, stored.name || name, stored, stored.updatedBy || '', stored.updatedAt || '');
}
async function saveMap(email, name, incoming, updatedBy, baseUpdatedAt) {
    return locked(async function () {
        const current = await readJson(mapBlobName(email), null);
        const currentAt = current && current.updatedAt ? current.updatedAt : '';
        if (baseUpdatedAt !== currentAt)
            throw new MapConflictError((current && current.updatedBy) || '', currentAt);
        const saved = (0, reflectiveMap_1.cleanMap)(email, name, incoming, updatedBy, new Date().toISOString());
        await writeJson(mapBlobName(email), saved, email);
        return saved;
    });
}
async function readScheduleWorkspace(email) {
    return readJson(scheduleBlobName(email), null);
}
async function listProviderEmails() {
    const found = [];
    const box = container();
    if (!(await box.exists().catch(function () { return false; })))
        return found;
    for await (const item of box.listBlobsFlat({ prefix: 'schedules/by-email/', includeMetadata: true })) {
        const owner = String((item.metadata && (item.metadata.owner || item.metadata.Owner)) || '').trim().toLowerCase();
        if (/^[a-z0-9._+-]+@emmeco\.org$/.test(owner) && found.indexOf(owner) === -1)
            found.push(owner);
    }
    for await (const item of box.listBlobsFlat({ prefix: 'reflective-supervision/maps/', includeMetadata: true })) {
        const owner = String((item.metadata && (item.metadata.owner || item.metadata.Owner)) || '').trim().toLowerCase();
        if (/^[a-z0-9._+-]+@emmeco\.org$/.test(owner) && found.indexOf(owner) === -1)
            found.push(owner);
    }
    found.sort();
    return found;
}
function blankAppointment() {
    return {
        id: '',
        providerEmail: '',
        providerName: '',
        adminEmail: '',
        adminName: '',
        start: '',
        end: '',
        suggestedStart: '',
        suggestedEnd: '',
        status: 'proposed',
        providerEventId: '',
        adminEventId: '',
        createdAt: '',
        updatedAt: '',
    };
}
function cleanAppointment(value) {
    if (!value || typeof value !== 'object')
        return null;
    const row = value;
    const item = blankAppointment();
    item.id = String(row.id || '').replace(/[^a-z0-9]/g, '').slice(0, 40);
    item.providerEmail = String(row.providerEmail || '').trim().toLowerCase();
    item.providerName = String(row.providerName || '').slice(0, 120);
    item.adminEmail = String(row.adminEmail || '').trim().toLowerCase();
    item.adminName = String(row.adminName || '').slice(0, 120);
    item.start = String(row.start || '');
    item.end = String(row.end || '');
    item.suggestedStart = String(row.suggestedStart || '');
    item.suggestedEnd = String(row.suggestedEnd || '');
    const status = String(row.status || '');
    const allowed = ['proposed', 'declined', 'suggested', 'accepted', 'confirmed', 'withdrawn'];
    item.status = allowed.indexOf(status) === -1 ? 'proposed' : status;
    item.providerEventId = String(row.providerEventId || '').slice(0, 200);
    item.adminEventId = String(row.adminEventId || '').slice(0, 200);
    item.createdAt = String(row.createdAt || '');
    item.updatedAt = String(row.updatedAt || '');
    if (!item.id || !item.providerEmail)
        return null;
    return item;
}
async function readAppointments() {
    const stored = await readJson(APPOINTMENTS, []);
    if (!Array.isArray(stored))
        return [];
    return stored.map(cleanAppointment).filter(function (item) { return !!item; });
}
async function writeAppointments(items) {
    await writeJson(APPOINTMENTS, items.slice(-400), '');
}
async function listAppointments(email) {
    const all = await readAppointments();
    return all.filter(function (item) { return item.providerEmail === email; });
}
async function listAllAppointments() {
    return readAppointments();
}
async function addAppointment(item) {
    return locked(async function () {
        const all = await readAppointments();
        const now = new Date().toISOString();
        const saved = Object.assign({}, item, {
            id: (0, crypto_1.randomBytes)(8).toString('hex'),
            createdAt: now,
            updatedAt: now,
        });
        all.push(saved);
        await writeAppointments(all);
        return saved;
    });
}
async function updateAppointment(id, change) {
    return locked(async function () {
        const all = await readAppointments();
        const item = all.find(function (row) { return row.id === id; });
        if (!item)
            return null;
        change(item);
        item.updatedAt = new Date().toISOString();
        await writeAppointments(all);
        return item;
    });
}
