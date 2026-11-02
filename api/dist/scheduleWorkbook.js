"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.spreadsheetStatus = spreadsheetStatus;
exports.lastVisitDate = lastVisitDate;
exports.buildScheduleWorkbook = buildScheduleWorkbook;
exports.openingsFromWorkbook = openingsFromWorkbook;
exports.applyScheduleWorkbook = applyScheduleWorkbook;
exports.listScheduleOpenings = listScheduleOpenings;
exports.syncScheduleWorkbook = syncScheduleWorkbook;
const exceljs_1 = __importDefault(require("exceljs"));
const graphClient_1 = require("./graphClient");
const providerSetup_1 = require("./providerSetup");
const schedulingService_1 = require("./schedulingService");
/**
 * Updates each provider's existing schedule workbook in
 * EMME/LEADERSHIP TEAM/EMME Processes/Referral Process/Provider Schedules.
 * The sheet stays in the Sam Schedule layout: one table per weekday, with
 * appointment time, status, client name, and Date of Last Visit.
 * Date of Last Visit is the day the client turns 3, or the day their
 * Medicaid or insurance PAR ends, whichever comes first.
 */
const STATUSES = [
    'Open 4x',
    'Open 3x (1,2,3)',
    'Open 3x (2,3,4)',
    'Open 3x (1,3,4)',
    'Open 2x (1&3)',
    'Open 2x (2&4)',
    'Open 1x (1)',
    'Open 1x (2)',
    'Open 1x (3)',
    'Open 1x (4)',
    'Full',
    'Scheduling in Progress',
    'Lunch/Paperwork',
    'Makeup',
    'Not Open',
];
const DAY_ORDER = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];
const DAY_HEADER = {
    monday: 'Mon Appointment Time ',
    tuesday: 'Tues Appointment Time ',
    wednesday: 'Wed Appointment Time ',
    thursday: 'Thurs Appointment Time ',
    friday: 'Fri Appointment Time ',
};
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
function dayKey(name) {
    const text = String(name || '').trim().toLowerCase();
    const found = DAY_ORDER.find(function (day) { return text.startsWith(day.slice(0, 3)); });
    return found || text || 'day';
}
function dayLabel(key) {
    return key.charAt(0).toUpperCase() + key.slice(1);
}
function timeHeader(key) {
    return DAY_HEADER[key] || (dayLabel(key).slice(0, 3) + ' Appointment Time ');
}
function tableName(key, used) {
    const base = dayLabel(key).replace(/[^A-Za-z0-9]/g, '') || 'Day';
    let name = base;
    let n = 2;
    while (used.has(name)) {
        name = base + n;
        n += 1;
    }
    used.add(name);
    return name;
}
function spreadsheetStatus(status) {
    const text = String(status || '').trim();
    const key = text.toLowerCase();
    if (key === 'lunch')
        return 'Lunch/Paperwork';
    if (key === 'scheduling in progress')
        return 'Scheduling in Progress';
    return text;
}
function timeFraction(text) {
    const match = String(text || '').trim().match(/^(\d{1,2})(?::(\d{2}))?\s*(am|pm)?$/i);
    if (!match)
        return null;
    let hour = Number(match[1]);
    const minute = Number(match[2] || '0');
    const suffix = String(match[3] || '').toLowerCase();
    if (suffix === 'pm' && hour < 12)
        hour += 12;
    if (suffix === 'am' && hour === 12)
        hour = 0;
    if (hour > 23 || minute > 59)
        return null;
    return (hour * 60 + minute) / (24 * 60);
}
function formatExcelTime(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        const hour = value.getUTCHours();
        const minute = value.getUTCMinutes();
        const suffix = hour >= 12 ? 'PM' : 'AM';
        const hour12 = hour % 12 || 12;
        return hour12 + ':' + String(minute).padStart(2, '0') + ' ' + suffix;
    }
    if (typeof value === 'number' && Number.isFinite(value)) {
        const minutes = Math.round((value % 1) * 24 * 60);
        const hour = Math.floor(minutes / 60) % 24;
        const minute = minutes % 60;
        const suffix = hour >= 12 ? 'PM' : 'AM';
        const hour12 = hour % 12 || 12;
        return hour12 + ':' + String(minute).padStart(2, '0') + ' ' + suffix;
    }
    return String(value || '').trim();
}
function parseLooseDate(value) {
    const text = String(value || '').trim();
    const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso)
        return new Date(Date.UTC(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])));
    const us = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
    if (!us)
        return null;
    let year = Number(us[3]);
    if (year < 100)
        year += 2000;
    const month = Number(us[1]);
    const day = Number(us[2]);
    if (month < 1 || month > 12 || day < 1 || day > 31)
        return null;
    return new Date(Date.UTC(year, month - 1, day));
}
function excelDateSerial(value) {
    return (Date.UTC(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate()) - Date.UTC(1899, 11, 30)) / 86400000;
}
function formatLooseDate(value) {
    if (!value || Number.isNaN(value.getTime()))
        return '';
    const month = String(value.getUTCMonth() + 1).padStart(2, '0');
    const day = String(value.getUTCDate()).padStart(2, '0');
    return month + '/' + day + '/' + value.getUTCFullYear();
}
function lastVisitDate(client) {
    if (!client)
        return '';
    const third = parseLooseDate(String(client.dob || ''));
    if (third)
        third.setUTCFullYear(third.getUTCFullYear() + 3);
    const par = parseLooseDate(String(client.parExpires || ''));
    let chosen = null;
    if (third && par)
        chosen = third.getTime() <= par.getTime() ? third : par;
    else
        chosen = third || par;
    return formatLooseDate(chosen);
}
function normName(value) {
    return String(value || '').toLowerCase().replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
}
function namesMatch(left, right) {
    const a = normName(left).split(' ').filter(Boolean);
    const b = normName(right).split(' ').filter(Boolean);
    if (!a.length || !b.length)
        return false;
    if (a.join(' ') === b.join(' '))
        return true;
    if (a.length >= 2 && b.length >= 2 && a.every(function (token) { return b.indexOf(token) !== -1; }))
        return true;
    if (a.length >= 2 && b.length >= 2 && b.every(function (token) { return a.indexOf(token) !== -1; }))
        return true;
    return false;
}
function clientForNotes(clients, notes) {
    const text = String(notes || '').trim();
    if (!text)
        return null;
    const list = clients || [];
    const exact = list.find(function (client) { return namesMatch(String(client.name || ''), text); });
    return exact || null;
}
function visitForSlot(slot, clients) {
    const computed = lastVisitDate(clientForNotes(clients, String(slot.notes || '')));
    return computed || String(slot.lastDay || '').trim();
}
function weekdayFromHeader(text) {
    const raw = String(text || '').trim();
    const lowered = raw.toLowerCase();
    if (lowered.indexOf('appointment') !== -1) {
        const prefix = lowered.replace(/appointment time/gi, '').trim();
        const found = DAY_ORDER.find(function (day) {
            return prefix === day || prefix.startsWith(day.slice(0, 3));
        });
        return found ? dayLabel(found) : '';
    }
    return DAY_ORDER.includes(lowered) ? dayLabel(lowered) : '';
}
async function buildScheduleWorkbook(workspace) {
    const board = (0, schedulingService_1.currentBoard)(workspace.boards);
    const days = ((board && board.days) || []).slice().sort(function (a, b) {
        const ia = DAY_ORDER.indexOf(dayKey(String(a.name || '')));
        const ib = DAY_ORDER.indexOf(dayKey(String(b.name || '')));
        return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);
    });
    const workbook = new exceljs_1.default.Workbook();
    workbook.creator = 'Emme Dashboard';
    const sheet = workbook.addWorksheet('Current Schedule');
    const lists = workbook.addWorksheet('Lists', { state: 'hidden' });
    STATUSES.forEach(function (status, index) {
        lists.getCell(index + 1, 1).value = status;
    });
    sheet.getColumn(1).width = 28;
    sheet.getColumn(2).width = 26;
    sheet.getColumn(3).width = 36;
    sheet.getColumn(4).width = 22;
    sheet.getColumn(6).width = 32;
    const usedNames = new Set();
    let row = 1;
    days.forEach(function (day) {
        const key = dayKey(String(day.name || ''));
        const slots = (day.slots || []).filter(function (slot) {
            return String(slot.time || '').trim() || String(slot.status || '').trim() || String(slot.notes || '').trim();
        });
        const data = (slots.length ? slots : [{ time: '', status: '', notes: '', lastDay: '' }]).map(function (slot) {
            const fraction = timeFraction(String(slot.time || ''));
            const visitText = visitForSlot(slot, workspace.clients);
            const visit = parseLooseDate(visitText);
            return [
                fraction == null ? String(slot.time || '') : fraction,
                spreadsheetStatus(String(slot.status || '')),
                String(slot.notes || ''),
                visit ? excelDateSerial(visit) : visitText,
            ];
        });
        const headerRow = row;
        sheet.getCell(headerRow, 6).value = 'Notes, Upcoming Changes';
        sheet.addTable({
            name: tableName(key, usedNames),
            ref: 'A' + headerRow,
            headerRow: true,
            columns: [
                { name: timeHeader(key), filterButton: true },
                { name: 'Status ', filterButton: true },
                { name: 'Notes (i.e. client names)', filterButton: true },
                { name: 'Date of Last Visit', filterButton: true },
            ],
            rows: data,
        });
        for (let i = 0; i < data.length; i++) {
            const timeCell = sheet.getCell(headerRow + 1 + i, 1);
            if (typeof timeCell.value === 'number')
                timeCell.numFmt = 'h:mm AM/PM';
            const visitCell = sheet.getCell(headerRow + 1 + i, 4);
            if (typeof visitCell.value === 'number' && visitCell.value > 1000)
                visitCell.numFmt = 'mm/dd/yyyy';
        }
        const statusStart = headerRow + 1;
        const statusEnd = headerRow + data.length;
        const validations = sheet;
        validations.dataValidations.add('B' + statusStart + ':B' + statusEnd, {
            type: 'list',
            allowBlank: true,
            formulae: ['Lists!$A$1:$A$' + STATUSES.length],
            showErrorMessage: true,
            errorTitle: 'Status',
            error: 'Choose a status from the list.',
        });
        row = statusEnd + 2;
    });
    const buffer = await workbook.xlsx.writeBuffer();
    return Buffer.from(buffer);
}
async function openingsFromWorkbook(bytes, staffName) {
    const workbook = new exceljs_1.default.Workbook();
    await workbook.xlsx.load(Buffer.from(bytes));
    const sheet = workbook.getWorksheet('Current Schedule') || workbook.worksheets[0];
    const openings = [];
    if (!sheet)
        return { staffName, source: 'sharepoint-schedule-workbook', openings };
    let day = '';
    sheet.eachRow(function (sheetRow) {
        const first = String(sheetRow.getCell(1).text || sheetRow.getCell(1).value || '');
        const headerDay = weekdayFromHeader(first);
        if (headerDay) {
            day = headerDay;
            return;
        }
        const status = String(sheetRow.getCell(2).text || sheetRow.getCell(2).value || '').trim();
        const time = formatExcelTime(sheetRow.getCell(1).value);
        if (!day || !status || !/^open\s+[1-4]x/i.test(status))
            return;
        openings.push({ day, time, status });
    });
    return { staffName, source: 'sharepoint-schedule-workbook', openings };
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
function scoreName(candidate, staffName, email) {
    const folder = compact(candidate);
    const name = compact(staffName);
    const code = shortCode(staffName);
    const local = compact((email || '').split('@')[0] || '');
    const first = compact(normName(staffName).split(' ')[0] || '');
    if (!folder)
        return 0;
    if (name && folder === name)
        return 100;
    if (code && folder === code)
        return 90;
    if (first && (folder === first || folder.indexOf(first) === 0))
        return 75;
    const tokens = normName(staffName).split(' ').filter(function (token) { return token.length > 2; });
    if (tokens.length >= 2 && tokens.every(function (token) { return folder.includes(token); }))
        return 80;
    if (local && (folder === local || folder.includes(local)))
        return 70;
    if (code && folder.includes(code))
        return 60;
    return 0;
}
let scheduleSiteCache = '';
async function scheduleSiteId() {
    if (scheduleSiteCache)
        return scheduleSiteCache;
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
    catch (_) { /* personal site may still work */ }
    if (!tries.length)
        throw new graphClient_1.GraphError(503, 'SharePoint is not configured for schedules.');
    const folder = providerSetup_1.SCHEDULE_ROOT.map(encodeURIComponent).join('/');
    for (let i = 0; i < tries.length; i++) {
        try {
            await (0, graphClient_1.graphGet)('/sites/' + tries[i] + '/drive/root:/' + folder + ':/children?$select=name&$top=1');
            scheduleSiteCache = tries[i];
            return scheduleSiteCache;
        }
        catch (err) {
            if (err instanceof graphClient_1.GraphError && (err.status === 401 || err.status === 403 || err.status === 404))
                continue;
            throw err;
        }
    }
    scheduleSiteCache = tries[0];
    return scheduleSiteCache;
}
function drivePath(segments) {
    return segments.map(encodeURIComponent).join('/');
}
async function listScheduleFolder(segments) {
    const siteId = await scheduleSiteId();
    const path = '/sites/' + siteId + '/drive/root:/' + drivePath(segments) + ':/children?$select=name,lastModifiedDateTime,folder,file';
    try {
        const data = await (0, graphClient_1.graphGet)(path);
        return ((data && data.value) || []).map(function (item) {
            return {
                name: String(item.name || ''),
                lastModified: String(item.lastModifiedDateTime || ''),
                isFolder: !!item.folder,
            };
        });
    }
    catch (err) {
        if (err instanceof graphClient_1.GraphError && (err.status === 404 || /itemNotFound/i.test(err.message)))
            return [];
        throw err;
    }
}
async function readScheduleBytes(segments) {
    try {
        const siteId = await scheduleSiteId();
        return await (0, graphClient_1.graphGetBytes)('/sites/' + siteId + '/drive/root:/' + drivePath(segments) + ':/content');
    }
    catch (err) {
        if (err instanceof graphClient_1.GraphError && (err.status === 404 || /not found|itemNotFound/i.test(err.message)))
            return null;
        throw err;
    }
}
async function writeScheduleBytes(segments, bytes) {
    const siteId = await scheduleSiteId();
    await (0, graphClient_1.graphSend)('/sites/' + siteId + '/drive/root:/' + drivePath(segments) + ':/content', {
        method: 'PUT',
        rawBody: bytes,
        contentType: XLSX_TYPE,
    });
}
function chooseScheduleFile(files, staffName, email) {
    const books = files.filter(function (file) {
        return !file.isFolder && /\.xlsx$/i.test(file.name) && file.name.indexOf('~$') !== 0;
    });
    if (!books.length)
        return '';
    const ranked = books.map(function (file) {
        return {
            name: file.name,
            modified: file.lastModified,
            score: scoreName(file.name, staffName, email) + (/schedule/i.test(file.name) ? 15 : 0),
        };
    });
    ranked.sort(function (a, b) { return b.score - a.score || b.modified.localeCompare(a.modified); });
    return ranked[0].name;
}
async function locateSchedule(staffName, email) {
    const found = await (0, providerSetup_1.findExistingProviderFile)(providerSetup_1.SCHEDULE_ROOT, email, staffName, 'schedule');
    if (!found || !found.fileName)
        return null;
    const segments = found.folder ? providerSetup_1.SCHEDULE_ROOT.concat([found.folder]) : providerSetup_1.SCHEDULE_ROOT.slice();
    return { segments, fileName: found.fileName, path: found.path };
}
function cellPlain(cell) {
    return String(cell.text || '').trim();
}
function dayBlocks(sheet) {
    const blocks = [];
    let block = null;
    const max = Math.max(sheet.rowCount || 0, 1);
    for (let index = 1; index <= max; index++) {
        const row = sheet.getRow(index);
        const header = weekdayFromHeader(cellPlain(row.getCell(1)));
        if (header) {
            block = { headerRow: index, rows: [] };
            blocks.push(block);
            continue;
        }
        if (!block)
            continue;
        const filled = [1, 2, 3, 4].some(function (col) { return cellPlain(row.getCell(col)) !== ''; });
        if (!filled) {
            block = null;
            continue;
        }
        block.rows.push(index);
    }
    return blocks;
}
function slotsForBlock(workspace, headerText) {
    const board = (0, schedulingService_1.currentBoard)(workspace.boards);
    const wanted = dayKey(weekdayFromHeader(headerText) || headerText);
    const day = ((board && board.days) || []).find(function (item) { return dayKey(String(item.name || '')) === wanted; });
    return ((day && day.slots) || []).filter(function (slot) {
        return String(slot.time || '').trim() || String(slot.status || '').trim() || String(slot.notes || '').trim();
    });
}
function writeVisitCell(cell, text) {
    const parsed = parseLooseDate(text);
    if (!parsed) {
        cell.value = text;
        return;
    }
    cell.value = excelDateSerial(parsed);
    cell.numFmt = 'mm/dd/yyyy';
}
async function applyScheduleWorkbook(bytes, workspace) {
    const workbook = new exceljs_1.default.Workbook();
    await workbook.xlsx.load(Buffer.from(bytes));
    const sheet = workbook.getWorksheet('Current Schedule') || workbook.worksheets[0];
    if (!sheet)
        throw new Error('The schedule workbook has no worksheet.');
    const blocks = dayBlocks(sheet);
    if (!blocks.length)
        throw new Error('The schedule form needs a weekday appointment table.');
    for (let b = blocks.length - 1; b >= 0; b--) {
        const block = blocks[b];
        const headerText = cellPlain(sheet.getRow(block.headerRow).getCell(1));
        const slots = slotsForBlock(workspace, headerText);
        while (block.rows.length < slots.length) {
            const insertAt = (block.rows.length ? block.rows[block.rows.length - 1] : block.headerRow) + 1;
            sheet.insertRow(insertAt, []);
            block.rows.push(insertAt);
        }
        slots.forEach(function (slot, index) {
            const row = sheet.getRow(block.rows[index]);
            const fraction = timeFraction(String(slot.time || ''));
            row.getCell(1).value = fraction == null ? String(slot.time || '') : fraction;
            if (typeof row.getCell(1).value === 'number')
                row.getCell(1).numFmt = 'h:mm AM/PM';
            row.getCell(2).value = spreadsheetStatus(String(slot.status || ''));
            row.getCell(3).value = String(slot.notes || '');
            writeVisitCell(row.getCell(4), visitForSlot(slot, workspace.clients));
        });
        for (let index = slots.length; index < block.rows.length; index++) {
            const row = sheet.getRow(block.rows[index]);
            row.getCell(1).value = null;
            row.getCell(2).value = null;
            row.getCell(3).value = null;
            row.getCell(4).value = null;
        }
    }
    return Buffer.from(await workbook.xlsx.writeBuffer());
}
async function listScheduleOpenings() {
    const entries = await listScheduleFolder(providerSetup_1.SCHEDULE_ROOT);
    const found = [];
    for (let i = 0; i < entries.length; i++) {
        const entry = entries[i];
        if (!entry.isFolder)
            continue;
        const inside = await listScheduleFolder(providerSetup_1.SCHEDULE_ROOT.concat([entry.name]));
        const fileName = chooseScheduleFile(inside, entry.name, '');
        if (!fileName)
            continue;
        const segments = providerSetup_1.SCHEDULE_ROOT.concat([entry.name, fileName]);
        const bytes = await readScheduleBytes(segments);
        if (!bytes)
            continue;
        const parsed = await openingsFromWorkbook(bytes, entry.name);
        found.push({
            staffName: entry.name,
            fileName,
            path: '/' + segments.join('/'),
            summary: parsed,
        });
    }
    return found;
}
async function syncScheduleWorkbook(workspace, email) {
    try {
        const staffEmail = String(email || '').trim().toLowerCase();
        const staffName = (0, graphClient_1.graphIsConfigured)() && staffEmail ? await (0, graphClient_1.getUserDisplayName)(staffEmail) : (staffEmail || 'Provider');
        const form = await locateSchedule(staffName, staffEmail);
        if (!form) {
            return { ok: false, error: 'No existing schedule folder for ' + staffEmail + '. Nothing new was created.' };
        }
        const current = await readScheduleBytes(form.segments.concat([form.fileName]));
        if (!current)
            return { ok: false, error: 'The schedule file for ' + staffName + ' could not be read.' };
        const next = await applyScheduleWorkbook(current, workspace);
        await writeScheduleBytes(form.segments.concat([form.fileName]), new Uint8Array(next));
        return { ok: true, path: form.path };
    }
    catch (err) {
        const message = err instanceof Error ? err.message : 'Could not update the SharePoint schedule';
        if (/not configured/i.test(message))
            return { ok: false, skipped: true };
        if (/403|401|accessdenied|authorization/i.test(message)) {
            return { ok: false, error: 'Microsoft denied access to the Provider Schedules folder.' };
        }
        return { ok: false, error: message.slice(0, 200) };
    }
}
