"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.rowsFromWorkbook = rowsFromWorkbook;
exports.mergeMileageRows = mergeMileageRows;
exports.syncMileageWorkbook = syncMileageWorkbook;
exports.syncAllMileageLogs = syncAllMileageLogs;
exports.confirmMileageRow = confirmMileageRow;
const storage_blob_1 = require("@azure/storage-blob");
const exceljs_1 = __importDefault(require("exceljs"));
const graphClient_1 = require("./graphClient");
const intakeQService_1 = require("./intakeQService");
const providerSetup_1 = require("./providerSetup");
/**
 * Adds a row to each provider's existing mileage reimbursement workbook.
 * The files live in EMME/PAYROLL & MILEAGE/Mileage, one folder per provider.
 * A trip is logged only when the client tile is In person with average miles,
 * Outlook has that client on the provider's calendar that day, and IntakeQ
 * shows the appointment was kept.
 */
const RATE = 0.67;
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const ID_HEADER = 'Dashboard Id';
function denverToday() {
    return new Intl.DateTimeFormat('en-CA', {
        timeZone: 'America/Denver',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
    }).format(new Date());
}
function windowStart(today) {
    const parts = today.split('-').map(Number);
    const prev = new Date(Date.UTC(parts[0], parts[1] - 2, 1));
    return prev.toISOString().slice(0, 10);
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
function inPerson(client) {
    return client.inPerson === true || client.inPerson === 'true';
}
function milesOf(client) {
    const raw = Number(String(client.milesPerTrip == null ? '' : client.milesPerTrip).replace(/[^0-9.]/g, ''));
    if (!Number.isFinite(raw) || raw <= 0 || raw > 400)
        return 0;
    return Math.round(raw * 10) / 10;
}
function appointmentDate(when) {
    const text = String(when || '').trim();
    const iso = text.match(/^(\d{4}-\d{2}-\d{2})/);
    if (iso)
        return iso[1];
    const ms = Date.parse(text);
    if (Number.isNaN(ms))
        return '';
    return new Intl.DateTimeFormat('en-CA', {
        timeZone: 'America/Denver',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
    }).format(new Date(ms));
}
function excelDate(iso) {
    const parts = iso.split('-').map(Number);
    return new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
}
function dateText(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        const year = value.getUTCFullYear();
        const month = String(value.getUTCMonth() + 1).padStart(2, '0');
        const day = String(value.getUTCDate()).padStart(2, '0');
        return year + '-' + month + '-' + day;
    }
    const text = String(value || '').trim();
    const iso = text.match(/^(\d{4}-\d{2}-\d{2})/);
    return iso ? iso[1] : text.slice(0, 20);
}
function money(miles) {
    return Math.round(miles * RATE * 100) / 100;
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
    if (!folder)
        return 0;
    if (name && folder === name)
        return 100;
    if (code && folder === code)
        return 90;
    const tokens = normName(staffName).split(' ').filter(function (token) { return token.length > 2; });
    if (tokens.length >= 2 && tokens.every(function (token) { return folder.includes(token); }))
        return 80;
    if (local && (folder === local || folder.includes(local)))
        return 70;
    if (code && folder.includes(code))
        return 60;
    return 0;
}
function columnKind(header) {
    const text = header.toLowerCase().replace(/\s+/g, ' ').trim();
    if (!text)
        return '';
    if (text === ID_HEADER.toLowerCase() || text === 'appointment id' || text === 'appointmentid')
        return 'appointmentId';
    if (text === 'date' || text.startsWith('date ') || text.includes('date of'))
        return 'date';
    if (text.includes('reimburse') || text.includes('amount') || text.includes('$'))
        return 'amount';
    if (text.includes('mile') || text === 'mi')
        return 'miles';
    if (text.includes('client') || text.includes('patient'))
        return 'client';
    if (text.includes('service') || text.includes('purpose') || text.includes('reason') || text.includes('description'))
        return 'purpose';
    if (text.includes('intake'))
        return 'intakeq';
    if (text.includes('billing') || text === 'status')
        return 'billing';
    return '';
}
function cellResult(value) {
    if (value && typeof value === 'object' && 'result' in value)
        return value.result;
    return value;
}
function cellDate(cell) {
    const value = cellResult(cell.value);
    if (typeof value === 'number' && value > 20000 && value < 80000) {
        return dateText(new Date(Math.round((value - 25569) * 86400000)));
    }
    const fromValue = dateText(value);
    if (/^\d{4}-\d{2}-\d{2}$/.test(fromValue))
        return fromValue;
    return dateText(cell.text);
}
function cellNumber(cell) {
    const value = cellResult(cell.value);
    if (typeof value === 'number' && Number.isFinite(value))
        return value;
    const parsed = Number(String(value == null ? cell.text : value).replace(/[^0-9.-]/g, ''));
    return Number.isFinite(parsed) ? parsed : NaN;
}
function cellString(cell) {
    const value = cellResult(cell.value);
    if (value == null || value === '')
        return String(cell.text || '').trim();
    return String(value).trim();
}
let mileageSiteCache = '';
async function mileageSiteId() {
    if (mileageSiteCache)
        return mileageSiteCache;
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
        throw new graphClient_1.GraphError(503, 'SharePoint is not configured for mileage.');
    const folder = providerSetup_1.MILEAGE_ROOT.map(encodeURIComponent).join('/');
    for (let i = 0; i < tries.length; i++) {
        try {
            await (0, graphClient_1.graphGet)('/sites/' + tries[i] + '/drive/root:/' + folder + ':/children?$select=name&$top=1');
            mileageSiteCache = tries[i];
            return mileageSiteCache;
        }
        catch (err) {
            if (err instanceof graphClient_1.GraphError && (err.status === 401 || err.status === 403 || err.status === 404))
                continue;
            throw err;
        }
    }
    mileageSiteCache = tries[0];
    return mileageSiteCache;
}
function drivePath(segments) {
    return segments.map(encodeURIComponent).join('/');
}
async function listMileage(segments) {
    const siteId = await mileageSiteId();
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
async function readWorkbookBytes(segments) {
    try {
        const siteId = await mileageSiteId();
        const path = '/sites/' + siteId + '/drive/root:/' + drivePath(segments) + ':/content';
        return await (0, graphClient_1.graphGetBytes)(path);
    }
    catch (err) {
        if (err instanceof graphClient_1.GraphError && (err.status === 404 || /not found|itemNotFound/i.test(err.message)))
            return null;
        throw err;
    }
}
async function writeWorkbook(segments, bytes) {
    const siteId = await mileageSiteId();
    const path = '/sites/' + siteId + '/drive/root:/' + drivePath(segments) + ':/content';
    await (0, graphClient_1.graphSend)(path, { method: 'PUT', rawBody: bytes, contentType: XLSX_TYPE });
}
function findHeader(sheet) {
    let best = null;
    let bestCount = 0;
    const max = Math.min(Math.max(sheet.rowCount || 0, 1), 40);
    for (let index = 1; index <= max; index++) {
        const columns = {};
        sheet.getRow(index).eachCell({ includeEmpty: false }, function (cell, col) {
            const kind = columnKind(String(cell.text || '').trim());
            if (kind && !columns[kind])
                columns[kind] = col;
        });
        const count = Object.keys(columns).length;
        if (columns.date && columns.miles && count > bestCount) {
            best = { headerRow: index, columns };
            bestCount = count;
        }
    }
    return best;
}
async function openWorkbook(bytes) {
    const workbook = new exceljs_1.default.Workbook();
    await workbook.xlsx.load(Buffer.from(bytes));
    const sheet = workbook.worksheets[0];
    if (!sheet)
        throw new Error('The mileage form has no worksheet.');
    const map = findHeader(sheet);
    if (!map)
        throw new Error('The mileage form needs a header row with Date and Miles.');
    return { workbook, sheet, map };
}
function chooseWorkbook(files, staffName, email) {
    const year = String(new Date().getFullYear());
    const books = files.filter(function (file) {
        return !file.isFolder && /\.xlsx$/i.test(file.name) && file.name.indexOf('~$') !== 0;
    });
    if (!books.length)
        return '';
    const ranked = books.map(function (file) {
        return {
            name: file.name,
            modified: file.lastModified,
            score: scoreName(file.name, staffName, email) + (file.name.indexOf(year) !== -1 ? 15 : 0) + (/mileage/i.test(file.name) ? 10 : 0),
        };
    });
    ranked.sort(function (a, b) { return b.score - a.score || b.modified.localeCompare(a.modified); });
    if (books.length === 1 || ranked[0].score >= 10)
        return ranked[0].name;
    return ranked[0].name;
}
async function locateForm(staffName, email) {
    const found = await (0, providerSetup_1.findExistingProviderFile)(providerSetup_1.MILEAGE_ROOT, email, staffName, 'mileage');
    if (!found || !found.fileName)
        return null;
    const segments = found.folder ? providerSetup_1.MILEAGE_ROOT.concat([found.folder]) : providerSetup_1.MILEAGE_ROOT.slice();
    return { segments, fileName: found.fileName, path: found.path };
}
async function rowsFromWorkbook(bytes, staffName, staffEmail, path) {
    const opened = await openWorkbook(bytes);
    const columns = opened.map.columns;
    const rows = [];
    opened.sheet.eachRow(function (row, index) {
        if (index <= opened.map.headerRow)
            return;
        const date = columns.date ? cellDate(row.getCell(columns.date)) : '';
        const client = columns.client ? cellString(row.getCell(columns.client)) : '';
        const appointmentId = columns.appointmentId ? cellString(row.getCell(columns.appointmentId)) : '';
        if (!date && !client && !appointmentId)
            return;
        const miles = columns.miles ? cellNumber(row.getCell(columns.miles)) : 0;
        const safeMiles = Number.isFinite(miles) ? miles : 0;
        const amount = columns.amount ? cellNumber(row.getCell(columns.amount)) : NaN;
        rows.push({
            id: appointmentId || staffEmail + ':' + date + ':' + client,
            staffName,
            staffEmail,
            date,
            client,
            service: columns.purpose ? cellString(row.getCell(columns.purpose)) : '',
            miles: safeMiles,
            amount: Number.isFinite(amount) ? amount : money(safeMiles),
            intakeq: columns.intakeq ? cellString(row.getCell(columns.intakeq)) : '',
            billing: (columns.billing ? cellString(row.getCell(columns.billing)) : '') || 'Pending',
            appointmentId,
            path,
        });
    });
    return rows;
}
function ensureIdColumn(sheet, map) {
    if (map.columns.appointmentId)
        return map.columns.appointmentId;
    let last = 1;
    sheet.getRow(map.headerRow).eachCell({ includeEmpty: false }, function (_cell, col) {
        if (col > last)
            last = col;
    });
    const col = last + 1;
    sheet.getRow(map.headerRow).getCell(col).value = ID_HEADER;
    sheet.getColumn(col).hidden = true;
    map.columns.appointmentId = col;
    return col;
}
function sampleFormat(sheet, map, kind, fallback) {
    const col = map.columns[kind];
    if (!col)
        return fallback;
    return sheet.getRow(map.headerRow + 1).getCell(col).numFmt || fallback;
}
async function appendRows(bytes, additions) {
    const opened = await openWorkbook(bytes);
    const idCol = ensureIdColumn(opened.sheet, opened.map);
    let last = opened.map.headerRow;
    opened.sheet.eachRow(function (row, index) {
        if (index <= opened.map.headerRow)
            return;
        const date = opened.map.columns.date ? cellDate(row.getCell(opened.map.columns.date)) : '';
        const client = opened.map.columns.client ? cellString(row.getCell(opened.map.columns.client)) : '';
        const id = cellString(row.getCell(idCol));
        if (date || client || id)
            last = index;
    });
    const dateFmt = sampleFormat(opened.sheet, opened.map, 'date', 'mm/dd/yyyy');
    const milesFmt = sampleFormat(opened.sheet, opened.map, 'miles', '0.0');
    const amountFmt = sampleFormat(opened.sheet, opened.map, 'amount', '"$"#,##0.00');
    additions.forEach(function (row, offset) {
        const line = opened.sheet.getRow(last + 1 + offset);
        const cols = opened.map.columns;
        if (cols.date && /^\d{4}-\d{2}-\d{2}$/.test(row.date)) {
            line.getCell(cols.date).value = excelDate(row.date);
            line.getCell(cols.date).numFmt = dateFmt;
        }
        if (cols.client)
            line.getCell(cols.client).value = row.client;
        if (cols.purpose)
            line.getCell(cols.purpose).value = row.service;
        if (cols.miles) {
            line.getCell(cols.miles).value = row.miles;
            line.getCell(cols.miles).numFmt = milesFmt;
        }
        if (cols.amount) {
            const sample = opened.sheet.getRow(opened.map.headerRow + 1).getCell(cols.amount).value;
            const formula = sample && typeof sample === 'object' && 'formula' in sample && typeof sample.formula === 'string' ? sample.formula : '';
            const targetRow = last + 1 + offset;
            if (formula) {
                const sourceRow = opened.map.headerRow + 1;
                line.getCell(cols.amount).value = { formula: formula.replace(new RegExp('\\b' + sourceRow + '\\b', 'g'), String(targetRow)) };
            }
            else {
                line.getCell(cols.amount).value = row.amount;
                line.getCell(cols.amount).numFmt = amountFmt;
            }
        }
        if (cols.intakeq)
            line.getCell(cols.intakeq).value = row.intakeq;
        if (cols.billing)
            line.getCell(cols.billing).value = row.billing || 'Pending';
        line.getCell(idCol).value = row.appointmentId;
    });
    return Buffer.from(await opened.workbook.xlsx.writeBuffer());
}
function calendarConfirms(events, date, clientName) {
    return events.some(function (event) {
        return event.date === date && namesMatch(clientName, event.subject);
    });
}
function mergeMileageRows(options) {
    const start = windowStart(options.today);
    const known = new Set(options.existing.map(function (row) { return row.appointmentId; }).filter(Boolean));
    const knownTrips = new Set(options.existing.map(function (row) { return row.date + '|' + normName(row.client); }));
    const rows = options.existing.slice();
    let added = 0;
    options.appointments.forEach(function (appointment) {
        if (!appointment.id || known.has(appointment.id))
            return;
        if (!options.staffEmail || appointment.practitionerEmail !== options.staffEmail)
            return;
        if ((0, intakeQService_1.isCancellationStatus)(appointment.status))
            return;
        const date = appointmentDate(appointment.when);
        if (!date || date < start || date > options.today)
            return;
        const client = (options.clients || []).find(function (item) {
            return inPerson(item) && namesMatch(String(item.name || ''), appointment.clientName);
        });
        if (!client)
            return;
        const miles = milesOf(client);
        if (!miles)
            return;
        const clientName = String(client.name || appointment.clientName);
        const tripKey = date + '|' + normName(clientName);
        if (knownTrips.has(tripKey))
            return;
        const onCalendar = calendarConfirms(options.events, date, clientName) || calendarConfirms(options.events, date, appointment.clientName);
        if (!onCalendar)
            return;
        known.add(appointment.id);
        knownTrips.add(tripKey);
        added += 1;
        rows.push({
            id: appointment.id,
            staffName: options.staffName,
            staffEmail: options.staffEmail,
            date,
            client: clientName,
            service: String(client.mileagePurpose || '').trim() || appointment.serviceName || 'Client session',
            miles,
            amount: money(miles),
            intakeq: appointment.status || 'Confirmed',
            billing: 'Pending',
            appointmentId: appointment.id,
            path: options.path,
        });
    });
    rows.sort(function (a, b) {
        return a.date.localeCompare(b.date) || a.client.localeCompare(b.client);
    });
    return { rows, added };
}
async function staffNameFor(email) {
    if ((0, graphClient_1.graphIsConfigured)() && email)
        return await (0, graphClient_1.getUserDisplayName)(email);
    return email || 'Provider';
}
async function calendarEvents(email, start, end) {
    const select = encodeURIComponent('subject,start,isCancelled');
    let path = '/users/' + encodeURIComponent(email) + '/calendarView?startDateTime=' + encodeURIComponent(start + 'T00:00:00') + '&endDateTime=' + encodeURIComponent(end + 'T23:59:59') + '&$select=' + select + '&$top=100';
    const events = [];
    for (let page = 0; page < 5 && path; page++) {
        const data = await (0, graphClient_1.graphGet)(path, { Prefer: 'outlook.timezone="America/Denver"' });
        ((data && data.value) || []).forEach(function (item) {
            if (item.isCancelled)
                return;
            const date = dateText(item.start && item.start.dateTime);
            const subject = String(item.subject || '').trim();
            if (date && subject)
                events.push({ date, subject });
        });
        path = data && data['@odata.nextLink'] ? String(data['@odata.nextLink']) : '';
    }
    return events;
}
async function syncMileageWorkbook(workspace, email, appointments) {
    try {
        const staffEmail = String(email || workspace.ownerEmail || '').trim().toLowerCase();
        const staffName = await staffNameFor(staffEmail);
        const hasInPerson = (workspace.clients || []).some(function (client) { return inPerson(client) && milesOf(client) > 0; });
        if (!hasInPerson)
            return { ok: true, skipped: true, added: 0 };
        const form = await locateForm(staffName, staffEmail);
        if (!form) {
            return { ok: false, error: 'No existing mileage folder for ' + staffEmail + '. Nothing new was created.' };
        }
        const currentBytes = await readWorkbookBytes(form.segments.concat([form.fileName]));
        if (!currentBytes)
            return { ok: false, error: 'The mileage form for ' + staffName + ' could not be read.' };
        const existing = await rowsFromWorkbook(currentBytes, staffName, staffEmail, form.path);
        const today = denverToday();
        const start = windowStart(today);
        const visits = appointments || await (0, intakeQService_1.getAppointments)(start, today);
        const events = await calendarEvents(staffEmail, start, today);
        const merged = mergeMileageRows({
            existing,
            appointments: visits,
            clients: workspace.clients || [],
            events,
            staffName,
            staffEmail,
            path: form.path,
            today,
        });
        if (!merged.added)
            return { ok: true, path: form.path, added: 0 };
        const existingIds = new Set(existing.map(function (row) { return row.appointmentId; }).filter(Boolean));
        const existingTrips = new Set(existing.map(function (row) { return row.date + '|' + normName(row.client); }));
        const additions = merged.rows.filter(function (row) {
            if (row.appointmentId && existingIds.has(row.appointmentId))
                return false;
            return !existingTrips.has(row.date + '|' + normName(row.client));
        });
        const next = await appendRows(currentBytes, additions);
        await writeWorkbook(form.segments.concat([form.fileName]), new Uint8Array(next));
        return { ok: true, path: form.path, added: merged.added };
    }
    catch (err) {
        const message = err instanceof Error ? err.message : 'Could not update the mileage workbook';
        if (/not configured/i.test(message))
            return { ok: false, skipped: true, error: message.slice(0, 200) };
        if (/403|401|accessdenied|authorization/i.test(message)) {
            return { ok: false, error: 'Microsoft denied calendar or mileage-folder access. Jenny needs to grant Calendars.Read and access to the payroll mileage folder.' };
        }
        return { ok: false, error: message.slice(0, 200) };
    }
}
async function listProviderWorkspaces() {
    const conn = process.env.AZURE_STORAGE_CONNECTION_STRING || '';
    if (!conn)
        return [];
    const container = storage_blob_1.BlobServiceClient.fromConnectionString(conn).getContainerClient(process.env.AZURE_STORAGE_CONTAINER || 'emme-reports');
    const found = [];
    for await (const blob of container.listBlobsFlat({ prefix: 'schedules/by-email/' })) {
        if (!blob.name.endsWith('.json'))
            continue;
        const client = container.getBlockBlobClient(blob.name);
        const text = (await client.downloadToBuffer()).toString('utf8');
        let workspace = {};
        try {
            workspace = JSON.parse(text);
        }
        catch (_) {
            workspace = {};
        }
        const email = String(workspace.ownerEmail || '').trim().toLowerCase();
        if (email)
            found.push({ email, workspace });
    }
    return found;
}
async function syncAllMileageLogs() {
    try {
        const providers = await listProviderWorkspaces();
        const today = denverToday();
        let appointments = [];
        let intakeError = '';
        try {
            appointments = await (0, intakeQService_1.getAppointments)(windowStart(today), today);
        }
        catch (err) {
            intakeError = err instanceof Error ? err.message : 'IntakeQ could not be read';
        }
        const rows = [];
        for (let i = 0; i < providers.length; i++) {
            const provider = providers[i];
            try {
                const result = await syncMileageWorkbook(provider.workspace, provider.email, appointments);
                if (!result.ok && result.error && !intakeError)
                    intakeError = result.error;
                if (!result.path)
                    continue;
                const staffName = await staffNameFor(provider.email);
                const form = await locateForm(staffName, provider.email);
                if (!form)
                    continue;
                const bytes = await readWorkbookBytes(form.segments.concat([form.fileName]));
                if (!bytes)
                    continue;
                const parsed = await rowsFromWorkbook(bytes, staffName, provider.email, form.path);
                parsed.forEach(function (row) { rows.push(row); });
            }
            catch (err) {
                if (!intakeError)
                    intakeError = err instanceof Error ? err.message : 'Could not update one mileage spreadsheet';
            }
        }
        rows.sort(function (a, b) { return b.date.localeCompare(a.date) || a.staffName.localeCompare(b.staffName); });
        return intakeError ? { rows, error: intakeError.slice(0, 200) } : { rows };
    }
    catch (err) {
        const message = err instanceof Error ? err.message : 'Could not read mileage workbooks';
        return { rows: [], error: message.slice(0, 200) };
    }
}
async function confirmMileageRow(staffEmail, appointmentId) {
    const email = String(staffEmail || '').trim().toLowerCase();
    const id = String(appointmentId || '').trim();
    if (!email || !id)
        return { ok: false, error: 'Staff and appointment are required' };
    try {
        const staffName = await staffNameFor(email);
        const form = await locateForm(staffName, email);
        if (!form)
            return { ok: false, error: 'No mileage form for this provider yet' };
        const bytes = await readWorkbookBytes(form.segments.concat([form.fileName]));
        if (!bytes)
            return { ok: false, error: 'No mileage form for this provider yet' };
        const opened = await openWorkbook(bytes);
        const idCol = opened.map.columns.appointmentId;
        const billingCol = opened.map.columns.billing;
        if (!idCol)
            return { ok: false, error: 'That session is not on the mileage form' };
        if (!billingCol)
            return { ok: false, error: 'This form has no billing column to confirm' };
        let found = false;
        opened.sheet.eachRow(function (line, index) {
            if (found || index <= opened.map.headerRow)
                return;
            if (cellString(line.getCell(idCol)) !== id)
                return;
            line.getCell(billingCol).value = 'Confirmed';
            found = true;
        });
        if (!found)
            return { ok: false, error: 'That session is not on the mileage form' };
        const next = Buffer.from(await opened.workbook.xlsx.writeBuffer());
        await writeWorkbook(form.segments.concat([form.fileName]), new Uint8Array(next));
        return { ok: true };
    }
    catch (err) {
        const message = err instanceof Error ? err.message : 'Could not confirm mileage';
        return { ok: false, error: message.slice(0, 200) };
    }
}
