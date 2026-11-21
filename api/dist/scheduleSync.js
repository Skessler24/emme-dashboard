"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.pullSchedule = pullSchedule;
exports.pushSchedule = pushSchedule;
exports.importInitialSchedule = importInitialSchedule;
exports.readProviderWorkbook = readProviderWorkbook;
exports.needsPush = needsPush;
const crypto_1 = __importDefault(require("crypto"));
const graphClient_1 = require("./graphClient");
const scheduleWorkbook_1 = require("./scheduleWorkbook");
const providerSetup_1 = require("./providerSetup");
/**
 * Two-way sync between a provider's dashboard schedule and their SharePoint workbook.
 *
 * `meta` is saved next to the provider's schedule and remembers what the workbook looked like
 * the last time both sides agreed (eTag plus the rows with their dashboard ids). That lets us tell
 * "someone edited the workbook" from "nothing changed", and line workbook rows up with dashboard slots.
 */
// eslint-disable-next-line @typescript-eslint/no-var-requires
const scheduleMerge = require('../schedule/merge');
const LOCATION_TTL_MS = 10 * 60 * 1000;
const locationCache = new Map();
function newId() {
    return 'slot_' + crypto_1.default.randomBytes(5).toString('hex');
}
function errorText(err, fallback) {
    const message = err instanceof Error ? err.message : fallback;
    if (/403|401|accessdenied|authorization/i.test(message))
        return 'Microsoft denied access to the Provider Schedules folder.';
    return message.slice(0, 200);
}
async function findLocation(email, meta) {
    const key = email.toLowerCase();
    const cached = locationCache.get(key);
    if (cached && cached.expires > Date.now())
        return cached.value;
    if (meta.location) {
        const info = await (0, scheduleWorkbook_1.scheduleItemInfo)(meta.location.segments.concat([meta.location.fileName]));
        if (info) {
            locationCache.set(key, { value: meta.location, expires: Date.now() + LOCATION_TTL_MS });
            return meta.location;
        }
    }
    const staffName = await (0, graphClient_1.getUserDisplayName)(email);
    const found = await (0, scheduleWorkbook_1.locateSchedule)(staffName, email);
    if (found)
        locationCache.set(key, { value: found, expires: Date.now() + LOCATION_TTL_MS });
    return found;
}
function dropLocation(email) {
    locationCache.delete(email.toLowerCase());
}
function hasRows(days) {
    return Object.keys(days || {}).some(function (key) { return (days[key] || []).length > 0; });
}
/**
 * When the workbook is somewhere else than last time (the folder was moved, or a different file was found),
 * what was remembered about the old file no longer applies. The next sync treats it as a first contact.
 */
function forgetOldFile(meta, location) {
    if (!meta.location || meta.location.path === location.path)
        return meta;
    return { location, syncedAt: meta.syncedAt };
}
function fileSegments(location) {
    return location.segments.concat([location.fileName]);
}
/**
 * Checks the provider's workbook. If someone edited it since the last sync, their changes are merged
 * into a copy of `workspace` (dashboard edits that are not in the workbook yet are kept) and returned.
 * The caller saves the returned workspace and meta.
 */
async function pullSchedule(workspace, email, meta) {
    if (!(0, graphClient_1.graphIsConfigured)() || !email)
        return { status: 'skipped', meta };
    try {
        const location = await findLocation(email, meta);
        if (!location)
            return { status: 'no-file', meta };
        meta = forgetOldFile(meta, location);
        const segments = fileSegments(location);
        const info = await (0, scheduleWorkbook_1.scheduleItemInfo)(segments);
        if (!info) {
            dropLocation(email);
            return { status: 'no-file', meta };
        }
        const nextMeta = Object.assign({}, meta, { location });
        if (meta.baseDays && meta.eTag && info.eTag && info.eTag === meta.eTag)
            return { status: 'unchanged', meta: nextMeta };
        const bytes = await (0, scheduleWorkbook_1.readScheduleBytes)(segments);
        if (!bytes)
            return { status: 'no-file', meta };
        const sheetDays = await (0, scheduleWorkbook_1.readScheduleDays)(bytes);
        if (!meta.baseDays) {
            // First contact. The dashboard is the source of truth, so the workbook is overwritten on the next push.
            const mine = (0, scheduleWorkbook_1.exportedDays)(workspace);
            if (!hasRows(mine) && hasRows(sheetDays)) {
                // Unless the dashboard has nothing yet: then the workbook's schedule is brought in, never wiped.
                const withIds = {};
                Object.keys(sheetDays).forEach(function (key) {
                    withIds[key] = sheetDays[key].map(function (row) { return Object.assign({}, row, { id: newId() }); });
                });
                nextMeta.baseDays = withIds;
                nextMeta.eTag = info.eTag;
                nextMeta.syncedAt = new Date().toISOString();
                return { status: 'imported', meta: nextMeta, workspace: (0, scheduleWorkbook_1.applyDaysToWorkspace)(workspace, withIds, newId) };
            }
            const base = {};
            Object.keys(sheetDays).forEach(function (key) {
                base[key] = (0, scheduleWorkbook_1.alignRows)(mine[key] || [], sheetDays[key], newId);
            });
            nextMeta.baseDays = base;
            nextMeta.eTag = info.eTag;
            return { status: 'unchanged', meta: nextMeta };
        }
        const known = meta.baseDays;
        const theirs = {};
        Object.keys(sheetDays).forEach(function (key) {
            theirs[key] = (0, scheduleWorkbook_1.alignRows)(known[key] || [], sheetDays[key], newId);
        });
        nextMeta.eTag = info.eTag;
        if ((0, scheduleWorkbook_1.fingerprintDays)(theirs) === (0, scheduleWorkbook_1.fingerprintDays)(known))
            return { status: 'unchanged', meta: nextMeta };
        const base = (0, scheduleWorkbook_1.applyDaysToWorkspace)(workspace, known, newId);
        const remote = (0, scheduleWorkbook_1.applyDaysToWorkspace)(workspace, theirs, newId);
        const merged = scheduleMerge.merge(base, workspace, remote);
        nextMeta.baseDays = theirs;
        nextMeta.syncedAt = new Date().toISOString();
        return { status: 'imported', meta: nextMeta, workspace: merged };
    }
    catch (err) {
        return { status: 'error', meta, error: errorText(err, 'Could not read the SharePoint schedule') };
    }
}
/**
 * Replaces the provider's read-only "<Name> - Dashboard Copy.xlsx" in Provider Schedules with the schedule
 * as it is on the dashboard now. A dashboard with nothing in it writes nothing.
 */
async function pushDashboardCopy(workspace, email, meta) {
    try {
        const desired = (0, scheduleWorkbook_1.exportedDays)(workspace);
        if (!hasRows(desired))
            return { ok: true, unchanged: true, meta };
        const fileName = (0, scheduleWorkbook_1.dashboardCopyName)(await (0, graphClient_1.getUserDisplayName)(email), email);
        const segments = providerSetup_1.SCHEDULE_ROOT.concat([fileName]);
        const path = '/' + segments.join('/');
        const fingerprint = (0, scheduleWorkbook_1.fingerprintDays)(desired);
        const info = await (0, scheduleWorkbook_1.scheduleItemInfo)(segments);
        if (info && meta.copyFingerprint === fingerprint && meta.copyPath === path) {
            return { ok: true, unchanged: true, path, meta };
        }
        const bytes = await (0, scheduleWorkbook_1.buildScheduleWorkbook)(workspace, { readOnly: true });
        await (0, scheduleWorkbook_1.writeScheduleBytes)(segments, new Uint8Array(bytes));
        return {
            ok: true,
            path,
            meta: Object.assign({}, meta, { copyFingerprint: fingerprint, copyPath: path, syncedAt: new Date().toISOString() }),
        };
    }
    catch (err) {
        return { ok: false, meta, error: errorText(err, 'Could not update the Dashboard Copy in SharePoint') };
    }
}
/**
 * Writes the dashboard schedule into the provider's workbook. Returns conflict when the workbook changed
 * after the last pull, so the caller can pull again and retry.
 */
async function pushSchedule(workspace, email, meta) {
    if (!(0, graphClient_1.graphIsConfigured)() || !email)
        return { ok: false, skipped: true, meta };
    // The default: a read-only "Dashboard Copy" file that every save replaces. No provider workbook is edited.
    if (process.env.SCHEDULE_TWO_WAY_SYNC !== '1')
        return pushDashboardCopy(workspace, email, meta);
    try {
        const location = await findLocation(email, meta);
        if (!location) {
            return { ok: false, meta, error: 'No existing schedule folder for ' + email + '. Nothing new was created.' };
        }
        meta = forgetOldFile(meta, location);
        const segments = fileSegments(location);
        const info = await (0, scheduleWorkbook_1.scheduleItemInfo)(segments);
        if (!info) {
            dropLocation(email);
            return { ok: false, meta, error: 'The schedule file for ' + email + ' could not be found.' };
        }
        const nextMeta = Object.assign({}, meta, { location });
        const twoWay = process.env.SCHEDULE_TWO_WAY_SYNC === '1';
        // One-way (the default): the dashboard wins, so a changed workbook is simply rewritten, never a conflict.
        if (twoWay && meta.eTag && info.eTag && info.eTag !== meta.eTag)
            return { ok: false, conflict: true, meta };
        const desired = (0, scheduleWorkbook_1.exportedDays)(workspace);
        if (meta.baseDays && meta.eTag && (0, scheduleWorkbook_1.fingerprintDays)(desired) === (0, scheduleWorkbook_1.fingerprintDays)(meta.baseDays) && (twoWay || info.eTag === meta.eTag)) {
            return { ok: true, unchanged: true, path: location.path, meta: nextMeta };
        }
        const current = await (0, scheduleWorkbook_1.readScheduleBytes)(segments);
        if (!current)
            return { ok: false, meta, error: 'The schedule file could not be read.' };
        if (!meta.baseDays && !hasRows(desired) && hasRows(await (0, scheduleWorkbook_1.readScheduleDays)(current))) {
            // An empty dashboard never blanks a workbook that has a schedule in it.
            return { ok: true, unchanged: true, path: location.path, meta: nextMeta };
        }
        const next = await (0, scheduleWorkbook_1.applyScheduleWorkbook)(current, workspace);
        let saved;
        try {
            saved = await (0, scheduleWorkbook_1.writeScheduleBytes)(segments, new Uint8Array(next), info.eTag || undefined);
        }
        catch (err) {
            if (err instanceof graphClient_1.GraphError && err.status === 412)
                return { ok: false, conflict: true, meta };
            throw err;
        }
        nextMeta.eTag = saved.eTag || ((await (0, scheduleWorkbook_1.scheduleItemInfo)(segments)) || { eTag: '' }).eTag;
        nextMeta.baseDays = desired;
        nextMeta.syncedAt = new Date().toISOString();
        return { ok: true, path: location.path, meta: nextMeta };
    }
    catch (err) {
        return { ok: false, meta, error: errorText(err, 'Could not update the SharePoint schedule') };
    }
}
/**
 * For a provider who has no dashboard schedule yet: builds one from their existing SharePoint workbook,
 * so opening the dashboard for the first time does not replace a schedule that is already in use.
 */
async function importInitialSchedule(email, meta) {
    if (!(0, graphClient_1.graphIsConfigured)() || !email)
        return { status: 'skipped', meta };
    try {
        const location = await findLocation(email, meta);
        if (!location)
            return { status: 'no-file', meta };
        const segments = fileSegments(location);
        const info = await (0, scheduleWorkbook_1.scheduleItemInfo)(segments);
        if (!info) {
            dropLocation(email);
            return { status: 'no-file', meta };
        }
        const bytes = await (0, scheduleWorkbook_1.readScheduleBytes)(segments);
        if (!bytes)
            return { status: 'no-file', meta };
        const sheetDays = await (0, scheduleWorkbook_1.readScheduleDays)(bytes);
        const keys = Object.keys(sheetDays);
        if (!keys.some(function (key) { return sheetDays[key].length > 0; }))
            return { status: 'empty', meta };
        const withIds = {};
        keys.forEach(function (key) {
            withIds[key] = sheetDays[key].map(function (row) { return Object.assign({}, row, { id: newId() }); });
        });
        const skeleton = {
            version: 1,
            boards: [{ id: 'board_' + crypto_1.default.randomBytes(4).toString('hex'), name: 'Current Schedule', days: [] }],
            clients: [],
            dxCodes: [],
        };
        const workspace = (0, scheduleWorkbook_1.applyDaysToWorkspace)(skeleton, withIds, newId);
        const nextMeta = { location, eTag: info.eTag, baseDays: withIds, syncedAt: new Date().toISOString() };
        return { status: 'imported', workspace, meta: nextMeta };
    }
    catch (err) {
        return { status: 'error', meta, error: errorText(err, 'Could not read the SharePoint schedule') };
    }
}
/**
 * The Import from folder button: the provider's existing workbook in Provider Schedules, as it is now.
 * Nothing is written to SharePoint and no sync state is kept.
 */
async function readProviderWorkbook(email) {
    if (!(0, graphClient_1.graphIsConfigured)() || !email)
        return { status: 'skipped', error: 'SharePoint is not connected.' };
    try {
        dropLocation(email);
        const location = await findLocation(email, {});
        if (!location)
            return { status: 'no-file' };
        const bytes = await (0, scheduleWorkbook_1.readScheduleBytes)(fileSegments(location));
        if (!bytes)
            return { status: 'no-file' };
        return { status: 'found', bytes, fileName: location.fileName, path: location.path };
    }
    catch (err) {
        return { status: 'error', error: errorText(err, 'Could not read the SharePoint schedule') };
    }
}
/** True when the dashboard has changes the workbook does not have yet (for example after a failed save). */
function needsPush(workspace, meta) {
    if (!meta || !meta.baseDays)
        return false;
    return (0, scheduleWorkbook_1.fingerprintDays)((0, scheduleWorkbook_1.exportedDays)(workspace)) !== (0, scheduleWorkbook_1.fingerprintDays)(meta.baseDays);
}
