"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.IntakeQError = void 0;
exports.intakeqIsConfigured = intakeqIsConfigured;
exports.getInvoices = getInvoices;
exports.getNoteSummaries = getNoteSummaries;
exports.isCancellationStatus = isCancellationStatus;
exports.getAppointments = getAppointments;
exports.getCancellations = getCancellations;
exports.complianceByStaff = complianceByStaff;
const config_1 = require("./config");
const config_2 = require("./config");
/**
 * Live reads from IntakeQ. Responses are not written to the Emme scheduling
 * database. Client phone, email, date of birth, and note text are dropped
 * before anything is returned to the dashboard.
 *
 * IntakeQ rate limit is about 10 requests per minute. Results are cached in
 * memory on the function instance for a few minutes and never persisted.
 */
const BASE = 'https://intakeq.com/api/v1';
const CACHE_MS = 5 * 60 * 1000;
const cache = new Map();
class IntakeQError extends Error {
    status;
    constructor(status, message) {
        super(message);
        this.name = 'IntakeQError';
        this.status = status;
    }
}
exports.IntakeQError = IntakeQError;
function intakeqIsConfigured() {
    return (0, config_1.intakeqConfigured)();
}
async function intakeGet(path) {
    const key = (0, config_2.getConfig)().intakeqApiKey;
    if (!key)
        throw new IntakeQError(503, 'IntakeQ is not configured. Set INTAKEQ_API_KEY.');
    const cached = cache.get(path);
    if (cached && cached.expires > Date.now())
        return cached.value;
    const res = await fetch(BASE + path, {
        headers: { 'X-Auth-Key': key, Accept: 'application/json' },
    });
    if (!res.ok) {
        throw new IntakeQError(res.status, 'IntakeQ request failed (' + res.status + ')');
    }
    const data = await res.json().catch(function () { return []; });
    cache.set(path, { expires: Date.now() + CACHE_MS, value: data });
    return data;
}
function asArray(data) {
    return Array.isArray(data) ? data : [];
}
function stamp(value) {
    if (value == null || value === '')
        return '';
    if (typeof value === 'number') {
        const ms = value > 1e12 ? value : value * 1000;
        return new Date(ms).toISOString();
    }
    return String(value);
}
function toInvoice(raw) {
    const amount = raw.TotalAmount ?? raw.Amount ?? raw.SubTotal ?? null;
    return {
        id: String(raw.Id || raw.id || ''),
        number: String(raw.Number || raw.InvoiceNumber || ''),
        status: String(raw.Status || ''),
        amount: amount == null ? null : Number(amount),
        amountDue: raw.AmountDue == null ? null : Number(raw.AmountDue),
        date: stamp(raw.Date || raw.DateCreated),
        dueDate: stamp(raw.DueDate),
        practitionerName: String(raw.PractitionerName || ''),
        practitionerEmail: String(raw.PractitionerEmail || '').toLowerCase(),
        clientName: String(raw.ClientName || ''),
    };
}
async function getInvoices(options) {
    const items = [];
    for (let page = 1; page <= 5; page++) {
        const qs = new URLSearchParams();
        if (options.status)
            qs.set('status', options.status);
        if (options.startDate)
            qs.set('startDate', options.startDate);
        if (options.endDate)
            qs.set('endDate', options.endDate);
        if (options.practitionerEmail)
            qs.set('practitionerEmail', options.practitionerEmail);
        qs.set('page', String(page));
        const batch = asArray(await intakeGet('/invoices?' + qs.toString()));
        batch.forEach(function (row) { items.push(toInvoice(row)); });
        if (batch.length < 100)
            break;
    }
    return items;
}
function noteKind(name) {
    const text = name.toLowerCase();
    if (/eval|evaluation|assessment/.test(text))
        return 'eval';
    if (/report/.test(text))
        return 'report';
    return 'note';
}
function noteIsOpen(status) {
    const text = String(status == null ? '' : status).trim().toLowerCase();
    if (!text)
        return false;
    if (text === 'locked' || text === '2' || text === 'complete' || text === 'completed')
        return false;
    return text === 'unlocked' || text === 'open' || text === 'draft' || text === '0' || text === '1';
}
function toNote(raw) {
    const noteName = String(raw.NoteName || raw.Name || '');
    const status = String(raw.Status == null ? '' : raw.Status);
    return {
        id: String(raw.Id || raw.id || ''),
        status,
        open: noteIsOpen(status),
        date: stamp(raw.Date || raw.DateCreated),
        noteName,
        kind: noteKind(noteName),
        practitionerEmail: String(raw.PractitionerEmail || '').toLowerCase(),
        practitionerName: String(raw.PractitionerName || ''),
        clientName: String(raw.ClientName || ''),
    };
}
let notesCache = null;
async function getNoteSummaries() {
    if (notesCache && notesCache.expires > Date.now())
        return notesCache.notes;
    const data = await intakeGet('/notes/summary');
    const notes = asArray(data).map(toNote);
    notesCache = { expires: Date.now() + CACHE_MS, notes };
    return notes;
}
function toAppointment(raw) {
    return {
        id: String(raw.Id || raw.id || ''),
        status: String(raw.Status || ''),
        when: String(raw.StartDateLocal || raw.StartDateIso || stamp(raw.StartDate) || ''),
        serviceName: String(raw.ServiceName || ''),
        practitionerEmail: String(raw.PractitionerEmail || '').toLowerCase(),
        practitionerName: String(raw.PractitionerName || ''),
        clientName: String(raw.ClientName || ''),
    };
}
function isCancellationStatus(status) {
    return /cancel|no-?show|missed|declined/i.test(status);
}
async function getAppointments(startDate, endDate) {
    const qs = new URLSearchParams({ startDate, endDate });
    return asArray(await intakeGet('/appointments?' + qs.toString())).map(toAppointment);
}
async function getCancellations(startDate, endDate, practitionerEmail) {
    const email = (practitionerEmail || '').trim().toLowerCase();
    const rows = await getAppointments(startDate, endDate);
    return rows.filter(function (row) {
        if (!isCancellationStatus(row.status))
            return false;
        if (email && row.practitionerEmail !== email)
            return false;
        return true;
    });
}
async function complianceByStaff() {
    const notes = await getNoteSummaries();
    const byEmail = new Map();
    notes.forEach(function (note) {
        if (!note.open || !note.practitionerEmail)
            return;
        let row = byEmail.get(note.practitionerEmail);
        if (!row) {
            row = {
                staffEmail: note.practitionerEmail,
                staffName: note.practitionerName || note.practitionerEmail,
                openNotes: 0,
                openEvals: 0,
                openReports: 0,
            };
            byEmail.set(note.practitionerEmail, row);
        }
        if (note.kind === 'eval')
            row.openEvals += 1;
        else if (note.kind === 'report')
            row.openReports += 1;
        else
            row.openNotes += 1;
    });
    return Array.from(byEmail.values()).sort(function (a, b) { return a.staffName.localeCompare(b.staffName); });
}
