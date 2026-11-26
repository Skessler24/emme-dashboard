"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.handleReflectiveRequest = handleReflectiveRequest;
const auth_1 = require("./auth");
const graphClient_1 = require("./graphClient");
const http_1 = require("./http");
const reflectiveMap_1 = require("./reflectiveMap");
const reflectiveOffice_1 = require("./reflectiveOffice");
const reflectiveStore_1 = require("./reflectiveStore");
/**
 * Growth map, PDF, SharePoint copy, and appointment requests.
 * A provider can open only their own map. An administrator can open any provider.
 */
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
function niceName(email) {
    const local = email.split('@')[0] || email;
    return local.replace(/[._-]+/g, ' ').replace(/\b[a-z]/g, function (c) { return c.toUpperCase(); });
}
async function personName(email, fallback) {
    try {
        const real = await (0, graphClient_1.getUserDisplayName)(email);
        if (real && real.indexOf('@') === -1)
            return real;
    }
    catch (_) { /* the name made from the email is used */ }
    return fallback || niceName(email);
}
async function callerOf(context, req) {
    const user = await (0, auth_1.requireUser)(context, req);
    if (!user)
        return null;
    const admin = await (0, auth_1.resolveIsAdmin)(user.email);
    const email = user.email.trim().toLowerCase();
    return { email, name: user.name || email, admin: admin.isAdmin };
}
function targetEmail(caller, requested) {
    const want = requested.trim().toLowerCase();
    if (!caller.admin) {
        if (want && want !== caller.email)
            return null;
        return caller.email;
    }
    if (!want)
        return caller.email;
    if (!/^[a-z0-9._+-]+@emmeco\.org$/.test(want))
        return null;
    return want;
}
function fail(context, err) {
    if (err instanceof reflectiveStore_1.MapConflictError) {
        (0, http_1.json)(context, 409, { error: err.message, updatedAt: err.updatedAt, updatedBy: err.updatedBy });
        return;
    }
    if (err instanceof reflectiveStore_1.StoreNotReadyError) {
        (0, http_1.json)(context, 503, { error: err.message });
        return;
    }
    throw err;
}
function receivesSupervision(email, name, supervisorEmail) {
    if (!/^[a-z0-9._+-]+@emmeco\.org$/.test(email))
        return false;
    if (supervisorEmail && email === supervisorEmail)
        return false;
    if (name.trim().toLowerCase() === reflectiveOffice_1.SUPERVISOR_NAME.toLowerCase())
        return false;
    return true;
}
function daysSince(earlier, later) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(earlier) || !/^\d{4}-\d{2}-\d{2}$/.test(later))
        return 99;
    const start = Date.parse(earlier + 'T00:00:00Z');
    const end = Date.parse(later + 'T00:00:00Z');
    return Math.round((end - start) / 86400000);
}
async function placeOnCalendars(item) {
    const problems = [];
    if (!item.adminEmail)
        problems.push('Kelly Fitzpatrick’s calendar is not connected yet.');
    if (!item.adminEventId && item.adminEmail) {
        try {
            const id = await (0, reflectiveOffice_1.addCalendarEvent)(item.adminEmail, item.start, item.end);
            if (!id)
                problems.push('Outlook did not add the event to the administrator calendar.');
            else
                item.adminEventId = id;
        }
        catch (err) {
            problems.push(err instanceof Error ? err.message : 'The admin calendar was not updated.');
        }
    }
    if (!item.providerEventId && item.providerEmail) {
        try {
            const id = await (0, reflectiveOffice_1.addCalendarEvent)(item.providerEmail, item.start, item.end);
            if (!id)
                problems.push('Outlook did not add the event to the provider calendar.');
            else
                item.providerEventId = id;
        }
        catch (err) {
            problems.push(err instanceof Error ? err.message : 'The provider calendar was not updated.');
        }
    }
    item.status = item.adminEventId && item.providerEventId ? 'confirmed' : 'accepted';
    return problems.join(' ');
}
async function handleReflectiveRequest(context, req) {
    const part = (0, http_1.queryOf)(req, 'part');
    const verb = (0, http_1.methodOf)(req);
    if (part === 'remind' && verb === 'POST' && (0, auth_1.automationAuthorized)(req)) {
        try {
            await sendDueReminders(context);
        }
        catch (err) {
            fail(context, err);
        }
        return;
    }
    const caller = await callerOf(context, req);
    if (!caller)
        return;
    try {
        if (part === 'remind' && verb === 'POST') {
            if (!caller.admin) {
                (0, http_1.json)(context, 403, { error: 'Emme Admin access is required' });
                return;
            }
            await sendDueReminders(context);
            return;
        }
        if (part === 'providers' && verb === 'GET') {
            if (!caller.admin) {
                (0, http_1.json)(context, 403, { error: 'Emme Admin access is required' });
                return;
            }
            const supervisor = await (0, reflectiveOffice_1.reflectiveSupervisor)();
            const emails = await (0, reflectiveStore_1.listProviderEmails)();
            const today = denverToday();
            const all = await (0, reflectiveStore_1.listAllAppointments)();
            const providers = (await Promise.all(emails.map(async function (email) {
                const mine = all.filter(function (item) { return item.providerEmail === email; });
                const name = await personName(email, niceName(email));
                return { email, name, standing: (0, reflectiveMap_1.monthStanding)(mine, today) };
            }))).filter(function (person) { return receivesSupervision(person.email, person.name, supervisor.email); });
            providers.sort(function (a, b) { return a.name.localeCompare(b.name); });
            (0, http_1.json)(context, 200, {
                providers,
                supervisor: { name: supervisor.name, email: supervisor.email, error: supervisor.error },
            });
            return;
        }
        const requested = (0, http_1.queryOf)(req, 'provider') || String((0, http_1.readBody)(req).provider || '');
        const email = targetEmail(caller, requested);
        if (!email) {
            (0, http_1.json)(context, 403, { error: 'You can open only your own reflective supervision map.' });
            return;
        }
        const providerName = await personName(email, niceName(email));
        if (part === 'map' && verb === 'GET') {
            const map = await (0, reflectiveStore_1.readMap)(email, providerName);
            if (!map.name)
                map.name = providerName;
            (0, http_1.json)(context, 200, { role: caller.admin ? 'admin' : 'provider', me: caller.email, map });
            return;
        }
        if (part === 'map' && verb === 'PUT') {
            const body = (0, http_1.readBody)(req);
            const saved = await (0, reflectiveStore_1.saveMap)(email, providerName, body.map, caller.email, String(body.baseUpdatedAt || ''));
            const copy = await (0, reflectiveOffice_1.saveReadOnlyCopy)(saved.name || providerName, email, (0, reflectiveMap_1.buildGrowthMapPdf)(saved));
            (0, http_1.json)(context, 200, { role: caller.admin ? 'admin' : 'provider', me: caller.email, map: saved, copy });
            return;
        }
        if (part === 'pdf' && verb === 'GET') {
            const map = await (0, reflectiveStore_1.readMap)(email, providerName);
            const pdf = (0, reflectiveMap_1.buildGrowthMapPdf)(map);
            const fileName = (map.name || 'Provider').replace(/[\\/:*?"<>|]/g, ' ') + ' - Reflective Supervision.pdf';
            context.res = {
                status: 200,
                headers: {
                    'Content-Type': 'application/pdf',
                    'Content-Disposition': 'attachment; filename="' + fileName.replace(/"/g, '') + '"',
                    'Cache-Control': 'no-store',
                },
                body: pdf,
            };
            return;
        }
        if (part === 'copy' && verb === 'POST') {
            const map = await (0, reflectiveStore_1.readMap)(email, providerName);
            const copy = await (0, reflectiveOffice_1.saveReadOnlyCopy)(map.name || providerName, email, (0, reflectiveMap_1.buildGrowthMapPdf)(map));
            (0, http_1.json)(context, 200, { copy });
            return;
        }
        if (part === 'sessions' && verb === 'GET') {
            const workspace = await (0, reflectiveStore_1.readScheduleWorkspace)(email);
            const start = denverToday();
            const outlook = await (0, reflectiveOffice_1.outlookBlocks)(email, start, addDays(start, 14));
            (0, http_1.json)(context, 200, {
                timeZone: 'America/Denver',
                visits: (0, reflectiveMap_1.clientVisitTimes)(workspace, 14, start),
                outlook: outlook.events,
                outlookError: outlook.error,
            });
            return;
        }
        if (part === 'appointments' && verb === 'GET') {
            const appointments = await (0, reflectiveStore_1.listAppointments)(email);
            (0, http_1.json)(context, 200, {
                appointments,
                standing: (0, reflectiveMap_1.monthStanding)(appointments, denverToday()),
                role: caller.admin ? 'admin' : 'provider',
            });
            return;
        }
        if (part === 'appointments' && verb === 'POST') {
            await changeAppointment(context, req, caller, email, providerName);
            return;
        }
        (0, http_1.json)(context, 404, { error: 'Not found' });
    }
    catch (err) {
        fail(context, err);
    }
}
async function changeAppointment(context, req, caller, email, providerName) {
    const body = (0, http_1.readBody)(req);
    const action = String(body.action || '');
    if (action === 'propose') {
        if (!caller.admin) {
            (0, http_1.json)(context, 403, { error: 'An administrator sets the reflective supervision time.' });
            return;
        }
        const start = String(body.start || '');
        const end = String(body.end || '') || ((0, reflectiveMap_1.validWhen)(start) ? (0, reflectiveMap_1.plusHour)(start) : '');
        if (!(0, reflectiveMap_1.validWhen)(start) || !(0, reflectiveMap_1.validWhen)(end) || end <= start) {
            (0, http_1.json)(context, 400, { error: 'Enter a start and end in Mountain Time.' });
            return;
        }
        const supervisor = await (0, reflectiveOffice_1.reflectiveSupervisor)();
        if (!receivesSupervision(email, providerName, supervisor.email)) {
            (0, http_1.json)(context, 400, { error: 'Kelly Fitzpatrick gives reflective supervision. Choose one of the other providers.' });
            return;
        }
        const saved = await (0, reflectiveStore_1.addAppointment)({
            providerEmail: email,
            providerName,
            adminEmail: supervisor.email,
            adminName: supervisor.name,
            start,
            end,
            suggestedStart: '',
            suggestedEnd: '',
            status: 'proposed',
            providerEventId: '',
            adminEventId: '',
        });
        const mailed = supervisor.email
            ? await (0, reflectiveOffice_1.notifyPerson)(supervisor.email, email, 'Reflective supervision time to confirm', 'Kelly Fitzpatrick proposed a reflective supervision time for ' + start.replace('T', ' ') + ' Mountain Time. Open Staff Expectations on the Emme dashboard to accept, decline, or suggest another time.')
            : supervisor.error;
        const mailWarning = mailed || supervisor.error;
        (0, http_1.json)(context, 200, { appointment: saved, mailWarning });
        return;
    }
    const id = String(body.id || '').replace(/[^a-z0-9]/g, '');
    const current = (await (0, reflectiveStore_1.listAppointments)(email)).find(function (item) { return item.id === id; });
    if (!current) {
        (0, http_1.json)(context, 404, { error: 'That appointment was not found.' });
        return;
    }
    if (!caller.admin && current.providerEmail !== caller.email) {
        (0, http_1.json)(context, 403, { error: 'You can answer only your own reflective supervision time.' });
        return;
    }
    if (action === 'accept' || action === 'decline' || action === 'suggest') {
        if (caller.admin && caller.email !== current.providerEmail) {
            (0, http_1.json)(context, 403, { error: 'The provider accepts, declines, or suggests a different time.' });
            return;
        }
    }
    if ((action === 'confirm' || action === 'withdraw' || action === 'calendar') && !caller.admin) {
        (0, http_1.json)(context, 403, { error: 'An administrator confirms the time.' });
        return;
    }
    const suggestStart = String(body.start || '');
    const suggestEnd = String(body.end || '') || ((0, reflectiveMap_1.validWhen)(suggestStart) ? (0, reflectiveMap_1.plusHour)(suggestStart) : '');
    if (action === 'suggest' && (!(0, reflectiveMap_1.validWhen)(suggestStart) || !(0, reflectiveMap_1.validWhen)(suggestEnd) || suggestEnd <= suggestStart)) {
        (0, http_1.json)(context, 400, { error: 'Enter the day and time you want instead.' });
        return;
    }
    let mailWarning = '';
    let calendarWarning = '';
    let saved;
    try {
        saved = await (0, reflectiveStore_1.updateAppointment)(id, function (item) {
            if (action === 'decline' && (item.status === 'proposed' || item.status === 'suggested'))
                item.status = 'declined';
            else if (action === 'suggest' && (item.status === 'proposed' || item.status === 'suggested')) {
                item.suggestedStart = suggestStart;
                item.suggestedEnd = suggestEnd;
                item.status = 'suggested';
            }
            else if (action === 'withdraw' && item.status !== 'confirmed')
                item.status = 'withdrawn';
            else if (action === 'accept' && item.status === 'proposed')
                item.status = 'accepted';
            else if (action === 'confirm' && item.status === 'suggested' && (0, reflectiveMap_1.validWhen)(item.suggestedStart) && (0, reflectiveMap_1.validWhen)(item.suggestedEnd)) {
                item.start = item.suggestedStart;
                item.end = item.suggestedEnd;
                item.status = 'accepted';
            }
            else if (action === 'calendar' && (item.status === 'accepted' || item.status === 'confirmed')) {
                /* The calendars are written again below when an event is still missing. */
            }
            else {
                throw new Error('That appointment cannot be changed that way.');
            }
        });
    }
    catch (err) {
        if (err instanceof Error && err.message === 'That appointment cannot be changed that way.') {
            (0, http_1.json)(context, 400, { error: err.message });
            return;
        }
        throw err;
    }
    if (!saved) {
        (0, http_1.json)(context, 404, { error: 'That appointment was not found.' });
        return;
    }
    const supervisor = await (0, reflectiveOffice_1.reflectiveSupervisor)();
    if (supervisor.email) {
        saved.adminEmail = supervisor.email;
        saved.adminName = supervisor.name;
    }
    if (saved.status === 'declined') {
        mailWarning = supervisor.email
            ? await (0, reflectiveOffice_1.notifyPerson)(supervisor.email, supervisor.email, 'Reflective supervision time declined', saved.providerName + ' declined the proposed reflective supervision time. Open Reflective Supervision on the admin dashboard to choose another time.', true)
            : supervisor.error;
    }
    else if (saved.status === 'suggested') {
        mailWarning = supervisor.email
            ? await (0, reflectiveOffice_1.notifyPerson)(supervisor.email, supervisor.email, 'Another reflective supervision time was suggested', saved.providerName + ' suggested a different reflective supervision time: ' + saved.suggestedStart.replace('T', ' ') + ' Mountain Time. Open Reflective Supervision on the admin dashboard to confirm it.', true)
            : supervisor.error;
    }
    else if ((action === 'accept' || action === 'confirm' || action === 'calendar') && (saved.status === 'accepted' || saved.status === 'confirmed')) {
        calendarWarning = await placeOnCalendars(saved);
        if (saved.status === 'confirmed' && supervisor.email) {
            mailWarning = await (0, reflectiveOffice_1.notifyPerson)(supervisor.email, saved.providerEmail, 'Reflective supervision is on the calendar', 'The reflective supervision time with Kelly Fitzpatrick is confirmed and is on both Outlook calendars: ' + saved.start.replace('T', ' ') + ' Mountain Time.');
        }
        else if (!supervisor.email)
            mailWarning = supervisor.error;
    }
    await (0, reflectiveStore_1.updateAppointment)(saved.id, function (item) {
        item.status = saved.status;
        item.adminEmail = saved.adminEmail;
        item.adminName = saved.adminName;
        item.adminEventId = saved.adminEventId;
        item.providerEventId = saved.providerEventId;
        item.start = saved.start;
        item.end = saved.end;
    });
    (0, http_1.json)(context, 200, { appointment: saved, mailWarning, calendarWarning });
}
async function sendDueReminders(context) {
    const supervisor = await (0, reflectiveOffice_1.reflectiveSupervisor)();
    const today = denverToday();
    const log = await (0, reflectiveStore_1.readReminderLog)();
    const reminded = [];
    if (!supervisor.email) {
        (0, http_1.json)(context, 200, { reminded, mailWarning: supervisor.error, supervisor: supervisor.name });
        return;
    }
    const emails = await (0, reflectiveStore_1.listProviderEmails)();
    const all = await (0, reflectiveStore_1.listAllAppointments)();
    const waitingForKelly = [];
    let mailWarning = '';
    for (let i = 0; i < emails.length; i++) {
        const email = emails[i];
        const name = await personName(email, niceName(email));
        if (!receivesSupervision(email, name, supervisor.email))
            continue;
        const mine = all.filter(function (item) { return item.providerEmail === email; });
        const standing = (0, reflectiveMap_1.monthStanding)(mine, today);
        const open = mine.find(function (item) {
            return item.status === 'proposed' || item.status === 'suggested' || item.status === 'accepted';
        });
        if (open && open.status !== 'proposed') {
            const when = (open.suggestedStart || open.start).replace('T', ' ');
            waitingForKelly.push(name + ' — ' + open.status + ' ' + when + ' Mountain Time');
            continue;
        }
        let about = '';
        if (open && open.status === 'proposed') {
            about = 'A reflective supervision time with Kelly Fitzpatrick is waiting for your answer: ' + open.start.replace('T', ' ') + ' Mountain Time. Open Staff Expectations on the Emme dashboard.';
        }
        else if (standing.short === 'Still needed' || standing.short === 'Behind') {
            about = 'This month’s reflective supervision session with Kelly Fitzpatrick is ' + standing.short.toLowerCase() + '. Open Staff Expectations on the Emme dashboard.';
        }
        if (!about || daysSince(log[email] || '', today) < 7)
            continue;
        mailWarning = await (0, reflectiveOffice_1.notifyPerson)(supervisor.email, email, 'Reflective supervision reminder', about);
        if (mailWarning)
            break;
        await (0, reflectiveStore_1.markReminded)(email, today);
        reminded.push({ email, name, about });
    }
    const summaryKey = 'summary:' + supervisor.email;
    if (!mailWarning && (reminded.length || waitingForKelly.length) && daysSince(log[summaryKey] || '', today) >= 7) {
        const lines = reminded.map(function (item) { return item.name + ' — ' + item.about; });
        const text = ['Reflective supervision reminders for this week.'].concat(lines, waitingForKelly).join('\n');
        mailWarning = await (0, reflectiveOffice_1.notifyPerson)(supervisor.email, supervisor.email, 'Reflective supervision reminders', text, true);
        if (!mailWarning)
            await (0, reflectiveStore_1.markReminded)(summaryKey, today);
    }
    (0, http_1.json)(context, 200, { reminded, waitingForKelly, mailWarning, supervisor: supervisor.name });
}
