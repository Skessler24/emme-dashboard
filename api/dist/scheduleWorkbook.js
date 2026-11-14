"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.spreadsheetStatus = spreadsheetStatus;
exports.lastVisitDate = lastVisitDate;
exports.buildScheduleWorkbook = buildScheduleWorkbook;
exports.openingsFromWorkbook = openingsFromWorkbook;
exports.scheduleSiteId = scheduleSiteId;
exports.drivePath = drivePath;
exports.readScheduleBytes = readScheduleBytes;
exports.scheduleItemInfo = scheduleItemInfo;
exports.writeScheduleBytes = writeScheduleBytes;
exports.locateSchedule = locateSchedule;
exports.readScheduleDays = readScheduleDays;
exports.fingerprintDays = fingerprintDays;
exports.exportedDays = exportedDays;
exports.alignRows = alignRows;
exports.applyDaysToWorkspace = applyDaysToWorkspace;
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
        const ms = value.getUTCHours() * 3600000 + value.getUTCMinutes() * 60000 + value.getUTCSeconds() * 1000 + value.getUTCMilliseconds();
        const total = Math.round(ms / 60000) % 1440;
        const hour = Math.floor(total / 60);
        const minute = total % 60;
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
function isFilledSlot(slot) {
    return !!(String(slot.time || '').trim() || String(slot.status || '').trim() || String(slot.notes || '').trim());
}
/**
 * Writes one weekday table starting at headerRow and returns the last row used.
 * listCount is the number of rows in the hidden Lists sheet, or 0 to skip the status drop-down.
 */
function addDayTable(sheet, key, headerRow, slots, clients, usedNames, listCount) {
    const data = (slots.length ? slots : [{ time: '', status: '', notes: '', lastDay: '' }]).map(function (slot) {
        const fraction = timeFraction(String(slot.time || ''));
        const visitText = visitForSlot(slot, clients);
        const visit = parseLooseDate(visitText);
        return [
            fraction == null ? String(slot.time || '') : fraction,
            spreadsheetStatus(String(slot.status || '')),
            String(slot.notes || ''),
            visit ? excelDateSerial(visit) : visitText,
        ];
    });
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
    const statusEnd = headerRow + data.length;
    if (listCount > 0) {
        const validations = sheet;
        validations.dataValidations.add('B' + (headerRow + 1) + ':B' + statusEnd, {
            type: 'list',
            allowBlank: true,
            formulae: ['Lists!$A$1:$A$' + listCount],
            showErrorMessage: true,
            errorTitle: 'Status',
            error: 'Choose a status from the list.',
        });
    }
    return statusEnd;
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
        const slots = (day.slots || []).filter(isFilledSlot);
        row = addDayTable(sheet, key, row, slots, workspace.clients, usedNames, STATUSES.length) + 2;
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
async function scheduleItemInfo(segments) {
    const siteId = await scheduleSiteId();
    try {
        const item = await (0, graphClient_1.graphGet)('/sites/' + siteId + '/drive/root:/' + drivePath(segments) + '?$select=id,eTag,lastModifiedDateTime');
        return { eTag: String((item && item.eTag) || ''), lastModified: String((item && item.lastModifiedDateTime) || '') };
    }
    catch (err) {
        if (err instanceof graphClient_1.GraphError && (err.status === 404 || /itemNotFound/i.test(err.message)))
            return null;
        throw err;
    }
}
/** Writes the workbook. When ifMatch is given, SharePoint refuses the write if someone saved the file since. */
async function writeScheduleBytes(segments, bytes, ifMatch) {
    const siteId = await scheduleSiteId();
    const saved = await (0, graphClient_1.graphSend)('/sites/' + siteId + '/drive/root:/' + drivePath(segments) + ':/content', {
        method: 'PUT',
        rawBody: bytes,
        contentType: XLSX_TYPE,
        headers: ifMatch ? { 'If-Match': ifMatch } : undefined,
    });
    return { eTag: String((saved && saved.eTag) || ''), lastModified: String((saved && saved.lastModifiedDateTime) || '') };
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
function columnNumber(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++)
        n = n * 26 + (letters.charCodeAt(i) - 64);
    return n;
}
function columnLetters(n) {
    let out = '';
    let rest = n;
    while (rest > 0) {
        const mod = (rest - 1) % 26;
        out = String.fromCharCode(65 + mod) + out;
        rest = Math.floor((rest - 1) / 26);
    }
    return out;
}
function sheetTables(sheet) {
    const raw = sheet.tables || {};
    return Object.keys(raw).map(function (name) { return raw[name]; });
}
function tableExtent(table) {
    const model = table.table || {};
    if (model.totalsRow)
        return null;
    const ref = String(model.tableRef || model.ref || '');
    const match = ref.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
    if (!match)
        return null;
    const row = Number(match[2]);
    return { row, col: columnNumber(match[1]), endRow: match[4] ? Number(match[4]) : row };
}
function dayBlocks(sheet) {
    const blocks = [];
    const extents = sheetTables(sheet).map(function (table) {
        return { table, extent: tableExtent(table) };
    });
    let block = null;
    const max = Math.max(sheet.rowCount || 0, 1);
    for (let index = 1; index <= max; index++) {
        const row = sheet.getRow(index);
        const header = weekdayFromHeader(cellPlain(row.getCell(1)));
        if (header) {
            const found = extents.find(function (entry) {
                return entry.extent && entry.extent.row === index && entry.extent.col === 1;
            });
            if (found && found.extent) {
                const rows = [];
                for (let r = index + 1; r <= found.extent.endRow; r++)
                    rows.push(r);
                blocks.push({ headerRow: index, rows, table: found.table, column: found.extent.col });
                index = found.extent.endRow;
                block = null;
                continue;
            }
            block = { headerRow: index, rows: [], table: null, column: 1 };
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
    return ((day && day.slots) || []).filter(isFilledSlot);
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
function cellDateText(cell) {
    const value = cell.value;
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        return formatLooseDate(new Date(Date.UTC(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate())));
    }
    return cellPlain(cell);
}
/** Reads every weekday table in the provider's workbook. Blank rows are skipped. */
async function readScheduleDays(bytes) {
    const workbook = new exceljs_1.default.Workbook();
    await workbook.xlsx.load(Buffer.from(bytes));
    const sheet = workbook.getWorksheet('Current Schedule') || workbook.worksheets[0];
    const days = {};
    if (!sheet)
        return days;
    dayBlocks(sheet).forEach(function (block) {
        const key = dayKey(weekdayFromHeader(cellPlain(sheet.getRow(block.headerRow).getCell(1))));
        const rows = days[key] || [];
        block.rows.forEach(function (index) {
            const row = sheet.getRow(index);
            const entry = {
                time: formatExcelTime(row.getCell(1).value),
                status: cellPlain(row.getCell(2)),
                notes: cellPlain(row.getCell(3)),
                lastDay: cellDateText(row.getCell(4)),
            };
            if (entry.time || entry.status || entry.notes)
                rows.push(entry);
        });
        days[key] = rows;
    });
    return days;
}
function normTime(text) {
    const fraction = timeFraction(String(text || ''));
    return fraction == null ? String(text || '').trim().toLowerCase() : String(Math.round(fraction * 1440));
}
function canonicalStatus(text) {
    return spreadsheetStatus(String(text || '')).trim().toLowerCase();
}
function normDate(text) {
    const parsed = parseLooseDate(String(text || ''));
    return parsed ? parsed.toISOString().slice(0, 10) : String(text || '').trim().toLowerCase();
}
function rowKey(row) {
    return [normTime(row.time), canonicalStatus(row.status), String(row.notes || '').trim().toLowerCase()].join('|');
}
function rowFingerprint(row) {
    return rowKey(row) + '|' + normDate(row.lastDay);
}
function fingerprintDays(days) {
    const out = {};
    Object.keys(days || {}).sort().forEach(function (key) {
        out[key] = (days[key] || []).map(rowFingerprint);
    });
    return JSON.stringify(out);
}
/** What the dashboard writes to SharePoint for the current board: filled slots only, keyed by weekday. */
function exportedDays(workspace) {
    const board = (0, schedulingService_1.currentBoard)(workspace.boards);
    const out = {};
    ((board && board.days) || []).forEach(function (day) {
        const key = dayKey(String(day.name || ''));
        if (out[key])
            return;
        out[key] = (day.slots || []).filter(isFilledSlot).map(function (slot) {
            return {
                id: String(slot.id || ''),
                time: String(slot.time || ''),
                status: String(slot.status || ''),
                notes: String(slot.notes || ''),
                lastDay: visitForSlot(slot, workspace.clients),
            };
        });
    });
    return out;
}
/**
 * Gives rows read from the workbook the ids of the rows they replace, so edits made in
 * Excel line up with the dashboard's slots instead of looking like deletes and adds.
 */
function alignRows(previous, current, makeId) {
    const n = previous.length;
    const m = current.length;
    const table = [];
    for (let i = 0; i <= n; i++)
        table.push(new Array(m + 1).fill(0));
    for (let i = n - 1; i >= 0; i--) {
        for (let j = m - 1; j >= 0; j--) {
            table[i][j] = rowKey(previous[i]) === rowKey(current[j])
                ? table[i + 1][j + 1] + 1
                : Math.max(table[i + 1][j], table[i][j + 1]);
        }
    }
    const matches = [];
    let a = 0;
    let b = 0;
    while (a < n && b < m) {
        if (rowKey(previous[a]) === rowKey(current[b])) {
            matches.push([a, b]);
            a++;
            b++;
        }
        else if (table[a + 1][b] >= table[a][b + 1]) {
            a++;
        }
        else {
            b++;
        }
    }
    matches.push([n, m]);
    const out = new Array(m);
    let pi = 0;
    let ci = 0;
    matches.forEach(function (match) {
        const gapPrev = match[0] - pi;
        const gapCurrent = match[1] - ci;
        for (let k = 0; k < gapCurrent; k++) {
            const source = k < gapPrev ? previous[pi + k] : null;
            out[ci + k] = Object.assign({}, current[ci + k], { id: (source && source.id) || makeId() });
        }
        if (match[0] < n) {
            out[match[1]] = Object.assign({}, current[match[1]], { id: previous[match[0]].id || makeId() });
        }
        pi = match[0] + 1;
        ci = match[1] + 1;
    });
    return out;
}
const DASHBOARD_STATUS = {
    'lunch/paperwork': 'Lunch',
    'scheduling in progress': 'Scheduling In Progress',
};
/**
 * Applies rows read from the workbook onto a copy of the workspace's current board.
 * Slots that are blank on the dashboard are not in the workbook, so they stay at the end of their day.
 */
function applyDaysToWorkspace(workspace, days, makeId) {
    const copy = JSON.parse(JSON.stringify(workspace));
    const board = (0, schedulingService_1.currentBoard)(copy.boards);
    if (!board)
        return copy;
    board.days = board.days || [];
    const boardDays = board.days;
    Object.keys(days || {}).forEach(function (key) {
        let day = boardDays.find(function (item) { return dayKey(String(item.name || '')) === key; });
        if (!day) {
            day = { id: 'day_sp_' + key, name: dayLabel(key), color: boardDays.length % 2 ? 'maroon' : 'blue', slots: [] };
            boardDays.push(day);
        }
        const existing = (day.slots || []);
        const byId = {};
        existing.forEach(function (slot) { if (slot.id)
            byId[slot.id] = slot; });
        const rows = (days[key] || []).map(function (row) {
            const prev = row.id ? byId[row.id] : undefined;
            const next = Object.assign({}, prev || {}, { id: row.id || makeId() });
            next.time = prev && normTime(prev.time || '') === normTime(row.time) ? prev.time : row.time;
            next.status = prev && canonicalStatus(prev.status || '') === canonicalStatus(row.status)
                ? prev.status
                : (DASHBOARD_STATUS[canonicalStatus(row.status)] || String(row.status || '').trim());
            next.notes = row.notes;
            const computed = prev ? visitForSlot(prev, workspace.clients) : '';
            if (normDate(row.lastDay) === normDate(computed))
                next.lastDay = prev ? prev.lastDay : '';
            else
                next.lastDay = row.lastDay;
            return next;
        });
        day.slots = rows.concat(existing.filter(function (slot) { return !isFilledSlot(slot); }));
    });
    return copy;
}
function lastUsedRow(sheet) {
    let last = 0;
    const max = sheet.rowCount || 0;
    for (let index = 1; index <= max; index++) {
        const row = sheet.getRow(index);
        if ([1, 2, 3, 4, 5, 6].some(function (col) { return cellPlain(row.getCell(col)) !== ''; }))
            last = index;
    }
    return last;
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
    const originalHeaders = blocks.map(function (block) { return block.headerRow; });
    const added = blocks.map(function () { return 0; });
    const seenKeys = new Set();
    blocks.forEach(function (block) {
        seenKeys.add(dayKey(weekdayFromHeader(cellPlain(sheet.getRow(block.headerRow).getCell(1)))));
    });
    const dashboardKeys = new Set(Object.keys(exportedDays(workspace)));
    for (let b = blocks.length - 1; b >= 0; b--) {
        const block = blocks[b];
        const headerText = cellPlain(sheet.getRow(block.headerRow).getCell(1));
        if (!dashboardKeys.has(dayKey(weekdayFromHeader(headerText) || headerText)))
            continue;
        const slots = slotsForBlock(workspace, headerText);
        const want = Math.max(slots.length, 1);
        while (block.rows.length < want) {
            const insertAt = (block.rows.length ? block.rows[block.rows.length - 1] : block.headerRow) + 1;
            sheet.insertRow(insertAt, [], block.rows.length ? 'i' : 'n');
            block.rows.push(insertAt);
            added[b] += 1;
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
    // insertRow moves cells but not tables or drop-downs, so put them back where the rows now are.
    const validationModel = (sheet.dataValidations || { model: {} }).model;
    let rule = null;
    Object.keys(validationModel).forEach(function (key) {
        if (/^B\d+(:B\d+)?$/.test(key) && validationModel[key]) {
            rule = rule || validationModel[key];
            delete validationModel[key];
        }
    });
    const listsSheet = workbook.getWorksheet('Lists');
    if (!rule && listsSheet && listsSheet.rowCount > 0) {
        rule = {
            type: 'list',
            allowBlank: true,
            formulae: ['Lists!$A$1:$A$' + listsSheet.rowCount],
            showErrorMessage: true,
            errorTitle: 'Status',
            error: 'Choose a status from the list.',
        };
    }
    let shift = 0;
    blocks.forEach(function (block, b) {
        shift += b > 0 ? added[b - 1] : 0;
        const header = originalHeaders[b] + shift;
        if (block.table) {
            const model = block.table.table;
            model.ref = columnLetters(block.column) + header;
            model.rows = new Array(block.rows.length).fill([]);
            block.table.validate();
        }
        if (rule && block.rows.length) {
            validationModel['B' + (header + 1) + ':B' + (header + block.rows.length)] = rule;
        }
    });
    const missing = Object.keys(exportedDays(workspace)).filter(function (key) {
        return !seenKeys.has(key) && exportedDays(workspace)[key].length > 0;
    });
    if (missing.length) {
        const usedNames = new Set(Object.keys(sheet.tables || {}));
        const order = missing.slice().sort(function (a, b) {
            const ia = DAY_ORDER.indexOf(a);
            const ib = DAY_ORDER.indexOf(b);
            return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);
        });
        let next = lastUsedRow(sheet) + 2;
        order.forEach(function (key) {
            const slots = slotsForBlock(workspace, dayLabel(key));
            next = addDayTable(sheet, key, next, slots, workspace.clients, usedNames, listsSheet ? listsSheet.rowCount : 0) + 2;
        });
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
