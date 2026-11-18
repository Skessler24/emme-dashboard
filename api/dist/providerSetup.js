"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.RS_OBSERVATION_ROOT = exports.SCHEDULE_ROOT = exports.MILEAGE_ROOT = void 0;
exports.emmeEmail = emmeEmail;
exports.providerMatchScore = providerMatchScore;
exports.looseFileScore = looseFileScore;
exports.resolveMileageRoot = resolveMileageRoot;
exports.findExistingProviderFile = findExistingProviderFile;
exports.searchProviderLibraries = searchProviderLibraries;
exports.listObservationProcess = listObservationProcess;
const graphClient_1 = require("./graphClient");
/**
 * Finds a provider's existing SharePoint folder by their emmeco.org email.
 * This only reads folders that are already there. It does not create folders or files.
 * An optional emme-provider.json in the folder can name the email directly.
 * Otherwise the folder or workbook name is matched to the provider's name and email.
 */
exports.MILEAGE_ROOT = ['EMME', 'PAYROLL & MILEAGE', 'Mileage'];
// The leadership files now sit at the top of the Documents library on the Emme team site (no "EMME" folder above them).
exports.SCHEDULE_ROOT = ['LEADERSHIP TEAM', 'EMME Processes', 'Referral Process', 'Provider Schedules'];
exports.RS_OBSERVATION_ROOT = ['LEADERSHIP TEAM', 'EMME Processes', 'RS_Observation Process'];
const MARKER = 'emme-provider.json';
function emmeEmail(value) {
    const email = String(value || '').trim().toLowerCase();
    return /@emmeco\.org$/.test(email) ? email : '';
}
function normName(value) {
    return String(value || '').toLowerCase().replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
}
function compact(value) {
    return normName(value).replace(/\s/g, '');
}
function shortCode(name) {
    const parts = normName(name).split(' ').filter(Boolean);
    if (parts.length < 2)
        return parts[0] || '';
    return parts[0].slice(0, 3) + parts[parts.length - 1].slice(0, 3);
}
function labelOf(value) {
    return compact(String(value || '')
        .replace(/\.xlsx$/i, '')
        .replace(/\b(schedule|mileage|reimbursement|form|20\d{2})\b/gi, ' '));
}
function providerMatchScore(candidate, staffName, email) {
    const folder = labelOf(candidate);
    const name = compact(staffName);
    const code = shortCode(staffName);
    const local = compact((email || '').split('@')[0] || '');
    if (!folder)
        return 0;
    if (local && folder === local)
        return 100;
    if (name && folder === name)
        return 95;
    if (code && folder === code)
        return 90;
    return 0;
}
/** Looser match for files named by first name ("Kelly Schedule.xlsx") or initials ("SB Schedule.xlsx"). */
function looseFileScore(candidate, staffName) {
    const label = labelOf(candidate);
    const parts = normName(staffName).split(' ').filter(Boolean);
    if (!label || !parts.length)
        return 0;
    const first = parts[0];
    const last = parts.length > 1 ? parts[parts.length - 1] : '';
    if (label === first)
        return 85;
    if (last && (label === first + last[0] || label === first[0] + last[0]))
        return 85;
    return 0;
}
/**
 * When the MILEAGE_FOLDER_URL setting holds a sharing link to the Mileage folder, the folder's real place is
 * read from SharePoint and used in place of the fixed path above. The answer is kept for ten minutes.
 */
let mileageResolvedAt = 0;
async function resolveMileageRoot() {
    const url = String(process.env.MILEAGE_FOLDER_URL || '').trim();
    if (!url || !(0, graphClient_1.graphIsConfigured)() || Date.now() - mileageResolvedAt < 10 * 60 * 1000)
        return exports.MILEAGE_ROOT;
    try {
        const id = 'u!' + Buffer.from(url, 'utf8').toString('base64').replace(/=+$/g, '').replace(/\//g, '_').replace(/\+/g, '-');
        const item = await (0, graphClient_1.graphGet)('/shares/' + encodeURIComponent(id) + '/driveItem?$select=name,parentReference');
        const parent = String((item && item.parentReference && item.parentReference.path) || '');
        const match = parent.match(/root:(.*)$/);
        if (match && item && item.name) {
            const segments = match[1].split('/').filter(Boolean).map(function (part) { return decodeURIComponent(part); }).concat([String(item.name)]);
            exports.MILEAGE_ROOT.splice(0, exports.MILEAGE_ROOT.length, ...segments);
            mileageResolvedAt = Date.now();
        }
    }
    catch (_) { /* keep the folder path that is already set */ }
    return exports.MILEAGE_ROOT;
}
async function librarySiteId() {
    await resolveMileageRoot();
    return (0, graphClient_1.siteHoldingFolder)(exports.MILEAGE_ROOT, 'SharePoint is not configured.');
}
function drivePath(segments) {
    return segments.map(encodeURIComponent).join('/');
}
async function listChildren(segments) {
    const siteId = await librarySiteId();
    const path = '/sites/' + siteId + '/drive/root:/' + drivePath(segments) + ':/children?$select=name,folder,file';
    try {
        const data = await (0, graphClient_1.graphGet)(path);
        return ((data && data.value) || []).map(function (item) {
            return { name: String(item.name || ''), isFolder: !!item.folder };
        });
    }
    catch (err) {
        if (err instanceof graphClient_1.GraphError && (err.status === 404 || /itemNotFound/i.test(err.message)))
            return [];
        throw err;
    }
}
async function markerEmail(segments) {
    const siteId = await librarySiteId();
    const path = '/sites/' + siteId + '/drive/root:/' + drivePath(segments.concat([MARKER])) + ':/content';
    try {
        const text = await (0, graphClient_1.graphGetText)(path);
        const parsed = JSON.parse(text);
        return emmeEmail(String(parsed && parsed.email || ''));
    }
    catch (_) {
        return '';
    }
}
function workbookName(files, kind) {
    const books = files.filter(function (file) {
        return !file.isFolder && /\.xlsx$/i.test(file.name) && file.name.indexOf('~$') !== 0;
    });
    const preferred = books.find(function (file) {
        return kind === 'mileage' ? /mileage/i.test(file.name) : /schedule/i.test(file.name);
    });
    return (preferred || books[0] || { name: '' }).name;
}
async function findExistingProviderFile(root, email, displayName, kind) {
    const providerEmail = emmeEmail(email);
    const entries = await listChildren(root);
    let emailFolder = '';
    for (let i = 0; i < entries.length; i++) {
        const item = entries[i];
        if (!item.isFolder)
            continue;
        const marked = await markerEmail(root.concat([item.name]));
        if (marked && marked === providerEmail) {
            emailFolder = item.name;
            break;
        }
    }
    let bestFolder = emailFolder;
    let bestScore = emailFolder ? 100 : 0;
    let matchedBy = emailFolder ? 'email' : 'name';
    if (!bestFolder) {
        entries.forEach(function (item) {
            if (!item.isFolder)
                return;
            const score = Math.max(providerMatchScore(item.name, displayName, providerEmail), providerMatchScore(item.name, displayName, ''));
            if (score > bestScore) {
                bestScore = score;
                bestFolder = item.name;
                matchedBy = providerEmail && compact(item.name) === compact(providerEmail.split('@')[0] || '') ? 'email' : 'name';
            }
        });
    }
    if (bestFolder && (emailFolder || bestScore >= 90)) {
        const inside = await listChildren(root.concat([bestFolder]));
        const fileName = workbookName(inside, kind);
        const segments = fileName ? root.concat([bestFolder, fileName]) : root.concat([bestFolder]);
        return { folder: bestFolder, fileName, path: '/' + segments.join('/'), matchedBy };
    }
    const rootBooks = entries.filter(function (item) { return !item.isFolder && /\.xlsx$/i.test(item.name) && item.name.indexOf('~$') !== 0; });
    let bestFile = '';
    let fileScore = 0;
    let tied = false;
    rootBooks.forEach(function (item) {
        const score = Math.max(providerMatchScore(item.name, displayName, providerEmail), looseFileScore(item.name, displayName));
        if (score > fileScore) {
            fileScore = score;
            bestFile = item.name;
            tied = false;
        }
        else if (score === fileScore && score > 0) {
            tied = true;
        }
    });
    // Files named only by first name ("Kelly Schedule.xlsx") count, but only when exactly one file fits.
    if (!bestFile || fileScore < 85 || (fileScore < 90 && tied) || !/\.xlsx$/i.test(bestFile))
        return null;
    return {
        folder: '',
        fileName: bestFile,
        path: '/' + root.concat([bestFile]).join('/'),
        matchedBy: 'name',
    };
}
async function searchProviderLibraries(email, displayName) {
    const providerEmail = emmeEmail(email);
    if (!providerEmail)
        throw new graphClient_1.GraphError(400, 'Use the provider’s emmeco.org email.');
    let name = String(displayName || '').trim();
    if (!name && (0, graphClient_1.graphIsConfigured)())
        name = await (0, graphClient_1.getUserDisplayName)(providerEmail);
    if (!name)
        name = providerEmail;
    await resolveMileageRoot();
    const mileage = await findExistingProviderFile(exports.MILEAGE_ROOT, providerEmail, name, 'mileage');
    const schedule = await findExistingProviderFile(exports.SCHEDULE_ROOT, providerEmail, name, 'schedule');
    return { email: providerEmail, name, mileage, schedule };
}
function assertFolderSegment(segment) {
    const name = String(segment || '').trim();
    if (!name || name === '.' || name === '..' || /[\\/]/.test(name) || name.length > 180) {
        throw new graphClient_1.GraphError(400, 'Folder path is not valid');
    }
    return name;
}
let observationSiteCache = { id: '', expires: 0 };
async function folderOnSite(siteId, segments) {
    try {
        const hit = await (0, graphClient_1.graphGet)('/sites/' + siteId + '/drive/root:/' + drivePath(segments) + '?$select=id,webUrl,folder');
        if (!hit || !hit.folder)
            return null;
        return { webUrl: String(hit.webUrl || '') };
    }
    catch (err) {
        if (err instanceof graphClient_1.GraphError && (err.status === 404 || /itemNotFound/i.test(err.message)))
            return null;
        throw err;
    }
}
/**
 * Reflective Supervision files live in
 * EMME/LEADERSHIP TEAM/EMME Processes/RS_Observation Process.
 * The shared Emme site is checked first, then the legacy personal library.
 */
async function observationSiteId() {
    if (observationSiteCache.id && observationSiteCache.expires > Date.now())
        return observationSiteCache.id;
    const tries = await (0, graphClient_1.candidateSiteIds)();
    let last = null;
    for (let i = 0; i < tries.length; i++) {
        try {
            const found = await folderOnSite(tries[i], exports.RS_OBSERVATION_ROOT);
            if (found) {
                // Remembered for a few minutes only, so a moved folder is found again without a restart.
                observationSiteCache = { id: tries[i], expires: Date.now() + 5 * 60 * 1000 };
                return tries[i];
            }
        }
        catch (err) {
            last = err;
            if (err instanceof graphClient_1.GraphError && (err.status === 401 || err.status === 403 || err.status === 404))
                continue;
            throw err;
        }
    }
    if (last instanceof graphClient_1.GraphError && last.status !== 404)
        throw last;
    throw new graphClient_1.GraphError(404, 'The Reflective Supervision folder was not found in SharePoint.');
}
async function listDriveChildren(siteId, segments) {
    let next = '/sites/' + siteId + '/drive/root:/' + drivePath(segments) +
        ':/children?$select=name,lastModifiedDateTime,folder,file,webUrl,size&$top=200';
    const items = [];
    for (let page = 0; page < 10 && next; page++) {
        const data = await (0, graphClient_1.graphGet)(next);
        const value = (data && data.value) || [];
        value.forEach(function (item) {
            const name = String(item.name || '');
            if (!name || name.indexOf('~$') === 0)
                return;
            const isFolder = !!item.folder;
            items.push({
                name,
                kind: isFolder ? 'folder' : 'file',
                lastModified: String(item.lastModifiedDateTime || ''),
                webUrl: String(item.webUrl || ''),
                size: isFolder || typeof item.size !== 'number' ? null : item.size,
            });
        });
        next = data && data['@odata.nextLink'] ? String(data['@odata.nextLink']) : null;
    }
    items.sort(function (a, b) {
        if (a.kind !== b.kind)
            return a.kind === 'folder' ? -1 : 1;
        return a.name.localeCompare(b.name, undefined, { sensitivity: 'base' });
    });
    return items;
}
async function listObservationProcess(relative) {
    const extra = (relative || []).map(assertFolderSegment);
    const siteId = await observationSiteId();
    const segments = exports.RS_OBSERVATION_ROOT.concat(extra);
    const folder = await folderOnSite(siteId, segments);
    if (!folder)
        throw new graphClient_1.GraphError(404, 'That Reflective Supervision folder was not found.');
    const items = await listDriveChildren(siteId, segments);
    return { path: '/' + segments.join('/'), webUrl: folder.webUrl, items };
}
