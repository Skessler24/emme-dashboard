"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.SCHEDULE_ROOT = exports.MILEAGE_ROOT = void 0;
exports.emmeEmail = emmeEmail;
exports.providerMatchScore = providerMatchScore;
exports.findExistingProviderFile = findExistingProviderFile;
exports.searchProviderLibraries = searchProviderLibraries;
const graphClient_1 = require("./graphClient");
/**
 * Finds a provider's existing SharePoint folder by their emmeco.org email.
 * This only reads folders that are already there. It does not create folders or files.
 * An optional emme-provider.json in the folder can name the email directly.
 * Otherwise the folder or workbook name is matched to the provider's name and email.
 */
exports.MILEAGE_ROOT = ['EMME', 'PAYROLL & MILEAGE', 'Mileage'];
exports.SCHEDULE_ROOT = ['EMME', 'LEADERSHIP TEAM', 'EMME Processes', 'Referral Process', 'Provider Schedules'];
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
let librarySiteCache = '';
async function librarySiteId() {
    if (librarySiteCache)
        return librarySiteCache;
    const tries = [];
    try {
        const personal = await (0, graphClient_1.graphGet)('/sites/emmeco031-my.sharepoint.com:/personal/jenriat_emmeco_org?$select=id');
        if (personal && personal.id)
            tries.push(String(personal.id));
    }
    catch (_) { /* the configured site is the fallback */ }
    try {
        const configured = await (0, graphClient_1.getSharePointSiteId)();
        if (tries.indexOf(configured) === -1)
            tries.push(configured);
    }
    catch (_) { /* the personal library may still be reachable */ }
    if (!tries.length)
        throw new graphClient_1.GraphError(503, 'SharePoint is not configured.');
    const folder = exports.MILEAGE_ROOT.map(encodeURIComponent).join('/');
    for (let i = 0; i < tries.length; i++) {
        try {
            await (0, graphClient_1.graphGet)('/sites/' + tries[i] + '/drive/root:/' + folder + ':/children?$select=name&$top=1');
            librarySiteCache = tries[i];
            return librarySiteCache;
        }
        catch (err) {
            if (err instanceof graphClient_1.GraphError && (err.status === 401 || err.status === 403 || err.status === 404))
                continue;
            throw err;
        }
    }
    librarySiteCache = tries[0];
    return librarySiteCache;
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
    const rootBooks = entries.filter(function (item) { return !item.isFolder; });
    let bestFile = '';
    let fileScore = 0;
    rootBooks.forEach(function (item) {
        const score = providerMatchScore(item.name, displayName, providerEmail);
        if (score > fileScore) {
            fileScore = score;
            bestFile = item.name;
        }
    });
    if (!bestFile || fileScore < 90 || !/\.xlsx$/i.test(bestFile))
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
    const mileage = await findExistingProviderFile(exports.MILEAGE_ROOT, providerEmail, name, 'mileage');
    const schedule = await findExistingProviderFile(exports.SCHEDULE_ROOT, providerEmail, name, 'schedule');
    return { email: providerEmail, name, mileage, schedule };
}
