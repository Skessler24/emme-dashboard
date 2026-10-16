"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.dispatch = dispatch;
const auth_1 = require("./auth");
const graphClient_1 = require("./graphClient");
const intakeQService_1 = require("./intakeQService");
const http_1 = require("./http");
const conversationStore_1 = require("./conversationStore");
const mileageWorkbook_1 = require("./mileageWorkbook");
const schedulingService_1 = require("./schedulingService");
const sharepointService_1 = require("./sharepointService");
function denverToday() {
    return new Intl.DateTimeFormat('en-CA', {
        timeZone: 'America/Denver',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
    }).format(new Date());
}
function addDays(iso, days) {
    const parts = iso.split('-').map(Number);
    const dt = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
    dt.setUTCDate(dt.getUTCDate() + days);
    return dt.toISOString().slice(0, 10);
}
function mondayOf(iso) {
    const parts = iso.split('-').map(Number);
    const dt = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
    const day = dt.getUTCDay();
    const shift = day === 0 ? -6 : 1 - day;
    return addDays(iso, shift);
}
function publicItem(item) {
    return Object.assign({ id: item.id, createdAt: item.createdAt || '' }, item.fields);
}
function matchesEmail(fields, email) {
    const wanted = email.trim().toLowerCase();
    const candidates = [fields.StaffEmail, fields.EmployeeEmail, fields.PractitionerEmail, fields.FromEmail];
    return candidates.some(function (value) { return String(value || '').trim().toLowerCase() === wanted; });
}
function safeError(err) {
    if (err instanceof graphClient_1.GraphError || err instanceof intakeQService_1.IntakeQError) {
        const status = err.status >= 400 && err.status < 600 ? err.status : 502;
        return { status: status === 401 || status === 403 ? 502 : status, message: err.message };
    }
    const message = err instanceof Error ? err.message : 'Request failed';
    return { status: 500, message: message.slice(0, 300) };
}
function windowDays(req) {
    const raw = Number((0, http_1.queryOf)(req, 'days') || '14');
    if (!Number.isFinite(raw))
        return 14;
    return Math.min(62, Math.max(1, Math.round(raw)));
}
async function me(context, req) {
    const principal = (0, auth_1.readPrincipal)(req);
    if (!principal || !principal.email) {
        (0, http_1.json)(context, 401, { isAdmin: false, error: 'Sign in required' });
        return;
    }
    const admin = await (0, auth_1.resolveIsAdmin)(principal.email);
    (0, http_1.json)(context, 200, {
        isAdmin: admin.isAdmin,
        email: principal.email,
        name: principal.name,
        adminCheck: admin.check,
    });
}
async function onboarding(context, req) {
    const admin = await (0, auth_1.requireAdmin)(context, req);
    if (!admin)
        return;
    const parts = (0, http_1.restParts)(req);
    const employeeId = parts[0] || '';
    if (!employeeId) {
        (0, http_1.json)(context, 400, { error: 'Employee id is required' });
        return;
    }
    const verb = (0, http_1.methodOf)(req);
    if (verb === 'GET' && parts.length === 1) {
        const items = await (0, sharepointService_1.listItems)(sharepointService_1.LISTS.onboarding);
        const checklist = items.filter(function (item) {
            return matchesEmail(item.fields, employeeId) || String(item.fields.EmployeeId || '') === employeeId;
        }).map(publicItem);
        const employeeName = (0, http_1.queryOf)(req, 'employeeName') || String((checklist[0] && checklist[0].EmployeeName) || '');
        let current = [];
        let archived = [];
        if (employeeName) {
            const folders = (0, sharepointService_1.onboardingFolders)(employeeName);
            const currentFiles = await (0, sharepointService_1.listFolder)(folders.current);
            const archivedFiles = await (0, sharepointService_1.listFolder)(folders.archived);
            current = currentFiles.filter(function (file) { return !file.isFolder; });
            archived = archivedFiles.filter(function (file) { return !file.isFolder; });
        }
        (0, http_1.json)(context, 200, { employeeId, checklist, currentDocuments: current, archivedDocuments: archived });
        return;
    }
    if (verb === 'POST' && parts[1] === 'documents') {
        const body = (0, http_1.readBody)(req);
        const employeeName = (0, http_1.asText)(body.employeeName, 80);
        const fileName = (0, http_1.asText)(body.fileName, 180);
        const contentBase64 = (0, http_1.asText)(body.contentBase64, 6_000_000);
        const folder = (0, http_1.asText)(body.folder, 20).toLowerCase() === 'archived' ? 'archived' : 'current';
        if (!employeeName || !fileName || !contentBase64) {
            (0, http_1.json)(context, 400, { error: 'employeeName, fileName, and contentBase64 are required' });
            return;
        }
        if (/[\\/]/.test(fileName) || fileName.indexOf('..') !== -1) {
            (0, http_1.json)(context, 400, { error: 'File name is not valid' });
            return;
        }
        const bytes = Buffer.from(contentBase64, 'base64');
        if (!bytes.length || bytes.length > 4 * 1024 * 1024) {
            (0, http_1.json)(context, 400, { error: 'File must be between 1 byte and 4 MB' });
            return;
        }
        const folders = (0, sharepointService_1.onboardingFolders)(employeeName);
        const saved = await (0, sharepointService_1.uploadEmployeeFile)(folder === 'archived' ? folders.archived : folders.current, fileName, new Uint8Array(bytes), 'application/octet-stream');
        (0, http_1.json)(context, 201, { ok: true, file: saved });
        return;
    }
    (0, http_1.json)(context, 404, { error: 'Not found' });
}
async function referrals(context, req) {
    const admin = await (0, auth_1.requireAdmin)(context, req);
    if (!admin)
        return;
    const verb = (0, http_1.methodOf)(req);
    if (verb === 'GET') {
        const items = await (0, sharepointService_1.listItems)(sharepointService_1.LISTS.referrals);
        (0, http_1.json)(context, 200, { referrals: items.map(publicItem) });
        return;
    }
    if (verb === 'POST') {
        const body = (0, http_1.readBody)(req);
        const referrerName = (0, http_1.asText)(body.referrerName, 200);
        if (!referrerName) {
            (0, http_1.json)(context, 400, { error: 'referrerName is required' });
            return;
        }
        const created = await (0, sharepointService_1.createItem)(sharepointService_1.LISTS.referrals, {
            Title: referrerName.slice(0, 250),
            ReferrerName: referrerName,
            ReceivedOn: (0, http_1.asText)(body.receivedOn, 40) || new Date().toISOString(),
            Status: (0, http_1.asText)(body.status, 40) || 'New',
            AssignedTo: (0, http_1.asText)(body.assignedTo, 200),
            Notes: (0, http_1.asText)(body.notes, 4000),
        });
        (0, http_1.json)(context, 201, { referral: publicItem(created) });
        return;
    }
    (0, http_1.json)(context, 405, { error: 'Method not allowed' });
}
async function trainings(context, req) {
    const admin = await (0, auth_1.requireAdmin)(context, req);
    if (!admin)
        return;
    if ((0, http_1.methodOf)(req) !== 'GET') {
        (0, http_1.json)(context, 405, { error: 'Method not allowed' });
        return;
    }
    const items = await (0, sharepointService_1.listItems)(sharepointService_1.LISTS.trainings);
    (0, http_1.json)(context, 200, { trainings: items.map(publicItem) });
}
async function staff(context, req) {
    const admin = await (0, auth_1.requireAdmin)(context, req);
    if (!admin)
        return;
    if ((0, http_1.methodOf)(req) !== 'GET') {
        (0, http_1.json)(context, 405, { error: 'Method not allowed' });
        return;
    }
    const parts = (0, http_1.restParts)(req);
    const staffId = (parts[0] || '').trim().toLowerCase();
    const action = parts[1] || '';
    if (!staffId || !action) {
        (0, http_1.json)(context, 400, { error: 'Staff id is required' });
        return;
    }
    if (action === 'trainings') {
        const items = await (0, sharepointService_1.listItems)(sharepointService_1.LISTS.staffTraining);
        (0, http_1.json)(context, 200, {
            staffId,
            trainings: items.filter(function (item) { return matchesEmail(item.fields, staffId); }).map(publicItem),
        });
        return;
    }
    if (action === 'open-notes' || action === 'open-evals') {
        const kind = action === 'open-evals' ? 'eval' : 'note';
        const notes = await (0, intakeQService_1.getNoteSummaries)();
        (0, http_1.json)(context, 200, {
            source: 'intakeq',
            persisted: false,
            staffId,
            notes: notes.filter(function (note) {
                return note.open && note.kind === kind && note.practitionerEmail === staffId;
            }),
        });
        return;
    }
    if (action === 'cancellations') {
        const end = (0, http_1.queryOf)(req, 'endDate') || denverToday();
        const start = (0, http_1.queryOf)(req, 'startDate') || addDays(end, -30);
        const rows = await (0, intakeQService_1.getCancellations)(start, end, staffId);
        (0, http_1.json)(context, 200, { source: 'intakeq', persisted: false, staffId, cancellations: rows });
        return;
    }
    (0, http_1.json)(context, 404, { error: 'Not found' });
}
function alertIsActive(fields, now) {
    const active = fields.Active === true || fields.Active === 'true' || fields.Active === 1;
    if (!active)
        return false;
    const start = fields.StartsOn ? Date.parse(String(fields.StartsOn)) : NaN;
    const end = fields.EndsOn ? Date.parse(String(fields.EndsOn)) : NaN;
    if (!Number.isNaN(start) && start > now)
        return false;
    if (!Number.isNaN(end) && end < now)
        return false;
    return true;
}
async function adminAlerts(context, req) {
    const verb = (0, http_1.methodOf)(req);
    const parts = (0, http_1.restParts)(req);
    if (verb === 'GET' && parts.length === 0) {
        const activeOnly = (0, http_1.queryOf)(req, 'active') === 'true';
        if (!activeOnly) {
            const admin = await (0, auth_1.requireAdmin)(context, req);
            if (!admin)
                return;
        }
        else {
            const user = await (0, auth_1.requireUser)(context, req);
            if (!user)
                return;
        }
        const now = Date.now();
        const alerts = (await (0, conversationStore_1.listAlerts)()).filter(function (item) {
            if (!activeOnly)
                return true;
            return alertIsActive(item, now);
        });
        (0, http_1.json)(context, 200, { alerts });
        return;
    }
    if (verb === 'POST' && parts.length === 0) {
        const admin = await (0, auth_1.requireAdmin)(context, req);
        if (!admin)
            return;
        const body = (0, http_1.readBody)(req);
        const message = (0, http_1.asText)(body.message, 240);
        if (!message) {
            (0, http_1.json)(context, 400, { error: 'A message is required' });
            return;
        }
        const created = await (0, conversationStore_1.addAlert)(message);
        (0, http_1.json)(context, 201, { ok: true, alert: created });
        return;
    }
    if (verb === 'POST' && parts[1] === 'end') {
        const admin = await (0, auth_1.requireAdmin)(context, req);
        if (!admin)
            return;
        await (0, conversationStore_1.endAlert)(parts[0]);
        (0, http_1.json)(context, 200, { ok: true });
        return;
    }
    (0, http_1.json)(context, 404, { error: 'Not found' });
}
async function staffMessages(context, req) {
    const verb = (0, http_1.methodOf)(req);
    const parts = (0, http_1.restParts)(req);
    if (verb === 'GET' && parts[0] === 'unread') {
        const user = await (0, auth_1.requireUser)(context, req);
        if (!user)
            return;
        const box = (0, http_1.queryOf)(req, 'box') === 'admin' ? 'admin' : 'staff';
        if (box === 'admin') {
            const admin = await (0, auth_1.requireAdmin)(context, req);
            if (!admin)
                return;
        }
        (0, http_1.json)(context, 200, { count: await (0, conversationStore_1.unreadCount)(box, user.email) });
        return;
    }
    if (verb === 'GET' && parts.length === 0) {
        const user = await (0, auth_1.requireUser)(context, req);
        if (!user)
            return;
        const box = (0, http_1.queryOf)(req, 'box');
        if (box !== 'staff') {
            const admin = await (0, auth_1.requireAdmin)(context, req);
            if (!admin)
                return;
        }
        const mine = user.email.trim().toLowerCase();
        const messages = (await (0, conversationStore_1.listMessages)()).filter(function (item) {
            return box === 'staff' ? item.FromEmail === mine : true;
        }).slice(0, 50);
        (0, http_1.json)(context, 200, { messages });
        return;
    }
    if (verb === 'POST' && parts[0] === 'seen') {
        const user = await (0, auth_1.requireUser)(context, req);
        if (!user)
            return;
        const body = (0, http_1.readBody)(req);
        const box = (0, http_1.asText)(body.box, 20) === 'admin' ? 'admin' : 'staff';
        if (box === 'admin') {
            const admin = await (0, auth_1.requireAdmin)(context, req);
            if (!admin)
                return;
        }
        await (0, conversationStore_1.markSeen)(box, user.email);
        (0, http_1.json)(context, 200, { ok: true });
        return;
    }
    if (verb === 'POST' && parts.length === 0) {
        const user = await (0, auth_1.requireUser)(context, req);
        if (!user)
            return;
        const body = (0, http_1.readBody)(req);
        const message = (0, http_1.asText)(body.message, 4000);
        if (!message) {
            (0, http_1.json)(context, 400, { error: 'A message is required' });
            return;
        }
        const created = await (0, conversationStore_1.addMessage)({ message, email: user.email, name: user.name || user.email });
        (0, http_1.json)(context, 201, { ok: true, id: created.id });
        return;
    }
    if (verb === 'POST' && parts[1] === 'reply') {
        const admin = await (0, auth_1.requireAdmin)(context, req);
        if (!admin)
            return;
        const body = (0, http_1.readBody)(req);
        const reply = (0, http_1.asText)(body.reply, 4000);
        if (!reply) {
            (0, http_1.json)(context, 400, { error: 'A reply is required' });
            return;
        }
        await (0, conversationStore_1.replyToMessage)(parts[0], reply, admin.email || admin.name);
        (0, http_1.json)(context, 200, { ok: true });
        return;
    }
    (0, http_1.json)(context, 404, { error: 'Not found' });
}
async function mileage(context, req) {
    const verb = (0, http_1.methodOf)(req);
    if (verb === 'GET') {
        const admin = await (0, auth_1.requireAdmin)(context, req);
        if (!admin)
            return;
        const result = await (0, mileageWorkbook_1.syncAllMileageLogs)();
        (0, http_1.json)(context, 200, { mileage: result.rows, syncError: result.error || '' });
        return;
    }
    if (verb === 'POST') {
        const body = (0, http_1.readBody)(req);
        if (body.action === 'confirm') {
            const admin = await (0, auth_1.requireAdmin)(context, req);
            if (!admin)
                return;
            const result = await (0, mileageWorkbook_1.confirmMileageRow)((0, http_1.asText)(body.staffEmail, 200), (0, http_1.asText)(body.appointmentId, 80));
            (0, http_1.json)(context, result.ok ? 200 : 400, result);
            return;
        }
        const automated = (0, auth_1.automationAuthorized)(req);
        const user = automated ? null : await (0, auth_1.requireUser)(context, req);
        if (!automated && !user)
            return;
        const miles = Number(body.miles);
        const purpose = (0, http_1.asText)(body.purpose, 500);
        const tripDate = (0, http_1.asText)(body.tripDate, 40);
        if (!tripDate || !purpose || !Number.isFinite(miles) || miles < 0) {
            (0, http_1.json)(context, 400, { error: 'tripDate, miles, and purpose are required' });
            return;
        }
        let staffEmail = (0, http_1.asText)(body.staffEmail, 200).toLowerCase();
        let staffName = (0, http_1.asText)(body.staffName, 200);
        if (!automated && user) {
            const admin = await (0, auth_1.resolveIsAdmin)(user.email);
            if (!admin.isAdmin) {
                staffEmail = user.email;
                staffName = user.name || user.email;
            }
        }
        if (!staffEmail)
            staffEmail = user ? user.email : '';
        if (!staffName)
            staffName = staffEmail;
        const created = await (0, sharepointService_1.createItem)(sharepointService_1.LISTS.mileage, {
            Title: staffName + ' ' + tripDate,
            StaffEmail: staffEmail,
            StaffName: staffName,
            TripDate: tripDate,
            Miles: miles,
            Purpose: purpose,
            Status: (0, http_1.asText)(body.status, 40) || 'Submitted',
        });
        (0, http_1.json)(context, 201, { mileage: publicItem(created) });
        return;
    }
    (0, http_1.json)(context, 405, { error: 'Method not allowed' });
}
async function cancellations(context, req) {
    const verb = (0, http_1.methodOf)(req);
    if (verb === 'GET') {
        const admin = await (0, auth_1.requireAdmin)(context, req);
        if (!admin)
            return;
        if ((0, http_1.queryOf)(req, 'source') === 'intakeq') {
            const end = (0, http_1.queryOf)(req, 'endDate') || denverToday();
            const start = (0, http_1.queryOf)(req, 'startDate') || addDays(end, -30);
            const rows = await (0, intakeQService_1.getCancellations)(start, end);
            (0, http_1.json)(context, 200, { source: 'intakeq', persisted: false, cancellations: rows });
            return;
        }
        const items = await (0, sharepointService_1.listItems)(sharepointService_1.LISTS.cancellations);
        (0, http_1.json)(context, 200, { source: 'sharepoint', cancellations: items.map(publicItem) });
        return;
    }
    if (verb === 'POST') {
        const admin = await (0, auth_1.requireAdmin)(context, req);
        if (!admin)
            return;
        const body = (0, http_1.readBody)(req);
        // POWER AUTOMATE nightly flow: POST { "sync": true } to copy recent
        // IntakeQ cancellations into the Cancellations list. Matched on ExternalId.
        if (body.sync === true) {
            const end = (0, http_1.asText)(body.endDate, 40) || denverToday();
            const start = (0, http_1.asText)(body.startDate, 40) || addDays(end, -30);
            const rows = await (0, intakeQService_1.getCancellations)(start, end);
            const existing = await (0, sharepointService_1.listItems)(sharepointService_1.LISTS.cancellations);
            const known = new Set(existing.map(function (item) { return String(item.fields.ExternalId || ''); }));
            let created = 0;
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                if (!row.id || known.has(row.id))
                    continue;
                await (0, sharepointService_1.createItem)(sharepointService_1.LISTS.cancellations, {
                    Title: (row.practitionerName || row.practitionerEmail || 'Cancellation') + ' ' + row.when.slice(0, 10),
                    StaffEmail: row.practitionerEmail,
                    StaffName: row.practitionerName,
                    AppointmentDate: row.when,
                    Status: row.status,
                    ServiceName: row.serviceName,
                    ClientName: row.clientName,
                    ExternalId: row.id,
                    Source: 'IntakeQ',
                });
                created += 1;
            }
            (0, http_1.json)(context, 200, { ok: true, created, scanned: rows.length });
            return;
        }
        (0, http_1.json)(context, 400, { error: 'Send { "sync": true } to update the Cancellations list from IntakeQ' });
        return;
    }
    (0, http_1.json)(context, 405, { error: 'Method not allowed' });
}
async function documentation(context, req) {
    const verb = (0, http_1.methodOf)(req);
    if (verb === 'GET') {
        const admin = await (0, auth_1.requireAdmin)(context, req);
        if (!admin)
            return;
        if ((0, http_1.queryOf)(req, 'source') === 'intakeq') {
            const rows = await (0, intakeQService_1.complianceByStaff)();
            (0, http_1.json)(context, 200, {
                source: 'intakeq',
                persisted: false,
                weekOf: mondayOf(denverToday()),
                rows,
            });
            return;
        }
        const items = await (0, sharepointService_1.listItems)(sharepointService_1.LISTS.compliance);
        (0, http_1.json)(context, 200, { source: 'sharepoint', rows: items.map(publicItem) });
        return;
    }
    if (verb === 'POST') {
        const admin = await (0, auth_1.requireAdmin)(context, req);
        if (!admin)
            return;
        // POWER AUTOMATE weekly flow: POST { "sync": true } to store this week's
        // open note / eval / report counts on DocumentationCompliance.
        const body = (0, http_1.readBody)(req);
        if (body.sync !== true) {
            (0, http_1.json)(context, 400, { error: 'Send { "sync": true } to store this week\'s documentation counts' });
            return;
        }
        const weekOf = (0, http_1.asText)(body.weekOf, 40) || mondayOf(denverToday());
        const rows = await (0, intakeQService_1.complianceByStaff)();
        const existing = await (0, sharepointService_1.listItems)(sharepointService_1.LISTS.compliance);
        let created = 0;
        let updated = 0;
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const match = existing.find(function (item) {
                return String(item.fields.StaffEmail || '').toLowerCase() === row.staffEmail && String(item.fields.WeekOf || '') === weekOf;
            });
            const fields = {
                Title: row.staffName + ' ' + weekOf,
                StaffEmail: row.staffEmail,
                StaffName: row.staffName,
                WeekOf: weekOf,
                OpenNotes: row.openNotes,
                OpenEvals: row.openEvals,
                OpenReports: row.openReports,
            };
            if (match) {
                await (0, sharepointService_1.updateItem)(sharepointService_1.LISTS.compliance, match.id, fields);
                updated += 1;
            }
            else {
                await (0, sharepointService_1.createItem)(sharepointService_1.LISTS.compliance, fields);
                created += 1;
            }
        }
        (0, http_1.json)(context, 200, { ok: true, weekOf, created, updated });
        return;
    }
    (0, http_1.json)(context, 405, { error: 'Method not allowed' });
}
async function billing(context, req) {
    const admin = await (0, auth_1.requireAdmin)(context, req);
    if (!admin)
        return;
    if ((0, http_1.methodOf)(req) !== 'GET' || (0, http_1.restParts)(req)[0] !== 'invoices') {
        (0, http_1.json)(context, 404, { error: 'Not found' });
        return;
    }
    const end = (0, http_1.queryOf)(req, 'endDate') || denverToday();
    const start = (0, http_1.queryOf)(req, 'startDate') || addDays(end, -90);
    const invoices = await (0, intakeQService_1.getInvoices)({
        status: (0, http_1.queryOf)(req, 'status'),
        startDate: start,
        endDate: end,
        practitionerEmail: (0, http_1.queryOf)(req, 'practitionerEmail'),
    });
    (0, http_1.json)(context, 200, { source: 'intakeq', persisted: false, invoices });
}
async function scheduling(context, req) {
    if ((0, http_1.methodOf)(req) !== 'GET') {
        (0, http_1.json)(context, 405, { error: 'Method not allowed' });
        return;
    }
    const action = (0, http_1.restParts)(req)[0] || '';
    if (action === 'availability') {
        // Anonymous callers must present the automation key. Signed-in admins may preview.
        // Power Automate uses this JSON and writes summary files to SharePoint itself.
        const admin = await (0, auth_1.requireAdmin)(context, req);
        if (!admin)
            return;
        const result = await (0, schedulingService_1.availabilityForWindow)(windowDays(req));
        if (!result.configured) {
            (0, http_1.json)(context, 503, {
                configured: false,
                error: 'Scheduling storage is not configured. Set AZURE_STORAGE_CONNECTION_STRING.',
            });
            return;
        }
        (0, http_1.json)(context, 200, {
            source: 'emme-scheduling-db',
            configured: true,
            asOf: result.asOf,
            windowDays: result.windowDays,
            generatedAt: result.generatedAt,
            timeZone: result.timeZone,
            staff: result.staff,
        });
        return;
    }
    if (action === 'openings') {
        const admin = await (0, auth_1.requireAdmin)(context, req);
        if (!admin)
            return;
        const files = await (0, sharepointService_1.readLatestScheduleSummaries)();
        (0, http_1.json)(context, 200, { source: 'sharepoint-schedule-summaries', staff: files });
        return;
    }
    (0, http_1.json)(context, 404, { error: 'Not found' });
}
const AREAS = {
    me,
    onboarding,
    referrals,
    trainings,
    staff,
    'admin-alerts': adminAlerts,
    'staff-messages': staffMessages,
    mileage,
    cancellations,
    'documentation-compliance': documentation,
    billing,
    scheduling,
};
async function dispatch(context, req, area) {
    if ((0, http_1.methodOf)(req) === 'OPTIONS') {
        context.res = { status: 204, headers: { 'Content-Type': 'application/json' } };
        return;
    }
    const handler = AREAS[area];
    if (!handler) {
        (0, http_1.json)(context, 404, { error: 'Not found' });
        return;
    }
    try {
        await handler(context, req);
    }
    catch (err) {
        const parsed = safeError(err);
        (0, http_1.json)(context, parsed.status, { error: parsed.message });
    }
}
