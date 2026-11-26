"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.expectationsView = expectationsView;
exports.expectationsAction = expectationsAction;
exports.expectationsError = expectationsError;
exports.pendingRequest = pendingRequest;
const intakeQService_1 = require("./intakeQService");
const mileageWorkbook_1 = require("./mileageWorkbook");
const sharepointService_1 = require("./sharepointService");
const expectationsRules_1 = require("./expectationsRules");
const expectationsStore_1 = require("./expectationsStore");
const expectationsZone_1 = require("./expectationsZone");
const schedulingService_1 = require("./schedulingService");
function clip(value, max) {
    return String(value == null ? '' : value).replace(/\u0000/g, '').trim().slice(0, max);
}
function emailOf(value) {
    return clip(value, 200).toLowerCase();
}
function roundMiles(value) {
    const miles = Number(value);
    if (!Number.isFinite(miles) || miles < 0 || miles > 400)
        return NaN;
    return Math.round(miles * 10) / 10;
}
function frequencyOf(value) {
    const text = clip(value, 40);
    if (text === 'monthly' || text === 'twice-monthly')
        return text;
    return 'weekly';
}
function own(rows, email, isAdmin) {
    if (isAdmin)
        return rows;
    return rows.filter(function (row) { return row.staffEmail === email; });
}
function monthTrips(trips, email, month) {
    return trips.filter(function (trip) { return trip.staffEmail === email && trip.date.slice(0, 7) === month; });
}
function publicHome(state, email) {
    const home = state.homes[email];
    if (!home)
        return null;
    return { address: home.address, mapLink: home.mapLink, saved: true };
}
function readableName(email, saved) {
    const name = String(saved || '').trim();
    if (name && name.indexOf('@') === -1)
        return name;
    const local = email.split('@')[0] || email;
    return local.split(/[._-]+/).filter(Boolean).map(function (part) {
        return part.charAt(0).toUpperCase() + part.slice(1);
    }).join(' ');
}
async function providerChoices(state, email, isAdmin) {
    const byEmail = new Map();
    function add(addressEmail, savedName) {
        const key = emailOf(addressEmail);
        if (!key)
            return;
        const home = state.homes[key];
        const current = byEmail.get(key);
        const name = readableName(key, (home && home.name) || savedName || (current && current.name) || '');
        byEmail.set(key, {
            email: key,
            name,
            address: home ? home.address : '',
            mapLink: home ? home.mapLink : '',
        });
    }
    if (isAdmin) {
        try {
            const loaded = await (0, schedulingService_1.loadStaffWorkspaces)();
            loaded.workspaces.forEach(function (row) { add(row.email, ''); });
        }
        catch (_) { /* schedules are optional; saved maps still show */ }
        Object.keys(state.homes).forEach(function (key) { add(key, state.homes[key].name); });
    }
    add(email, '');
    return Array.from(byEmail.values()).sort(function (a, b) { return a.name.localeCompare(b.name); });
}
async function notesFor(email, isAdmin, now) {
    const notes = await (0, intakeQService_1.getNoteSummaries)();
    const late = [];
    const evals = [];
    const reports = [];
    notes.forEach(function (note) {
        if (!note.open)
            return;
        if (!isAdmin && note.practitionerEmail !== email)
            return;
        const row = {
            id: note.id,
            staffEmail: note.practitionerEmail,
            staffName: note.practitionerName || note.practitionerEmail,
            noteName: note.noteName,
            date: note.date,
            clientName: note.clientName,
            hoursOpen: (0, expectationsRules_1.hoursSince)(note.date, now),
        };
        if (note.kind === 'note' && (0, expectationsRules_1.noteIsLate)(note.date, now))
            late.push(row);
        if (note.kind === 'eval')
            evals.push(Object.assign({}, row, { pastEndOfWeek: (0, expectationsRules_1.pastEndOfWeek)(note.date, now) }));
        if (note.kind === 'report')
            reports.push(Object.assign({}, row, { pastEndOfWeek: (0, expectationsRules_1.pastEndOfWeek)(note.date, now) }));
    });
    return { late, evals, reports };
}
async function cancellationsFor(email, isAdmin, state, today) {
    const start = (0, expectationsRules_1.addDays)(today, -183);
    const items = await (0, sharepointService_1.listItems)(sharepointService_1.LISTS.cancellations);
    const grouped = new Map();
    items.forEach(function (item) {
        const fields = item.fields;
        const staffEmail = emailOf(String(fields.StaffEmail || ''));
        if (!isAdmin && staffEmail !== email)
            return;
        const when = clip(fields.AppointmentDate, 40);
        const day = when.slice(0, 10);
        if (day && start && day < start)
            return;
        const status = clip(fields.Status, 80);
        const classified = (0, expectationsRules_1.cancellationClass)(status);
        if (!classified.productivity && !classified.incident)
            return;
        const client = clip(fields.ClientName, 120);
        const key = (0, expectationsRules_1.clientKey)(client) || 'unknown';
        let group = grouped.get(staffEmail + '|' + key);
        if (!group) {
            group = { client, staffName: clip(fields.StaffName, 120), staffEmail, productivity: 0, incidents: 0, rows: [] };
            grouped.set(staffEmail + '|' + key, group);
        }
        if (classified.productivity)
            group.productivity += 1;
        if (classified.incident)
            group.incidents += 1;
        group.rows.push({ date: day, status, kind: classified.kind, service: clip(fields.ServiceName, 120) });
    });
    return Array.from(grouped.entries()).map(function (entry) {
        const key = entry[0].split('|').slice(1).join('|');
        const group = entry[1];
        const frequency = state.frequency[entry[0]] || state.frequency[key] || 'weekly';
        const limit = (0, expectationsRules_1.incidentLimit)(frequency);
        const outreach = state.outreach[entry[0]] || null;
        return {
            key: entry[0],
            client: group.client,
            staffName: group.staffName,
            staffEmail: group.staffEmail,
            frequency,
            limit,
            productivity: group.productivity,
            incidents: group.incidents,
            needsOutreach: group.incidents >= limit,
            outreach,
            message: group.incidents >= limit ? expectationsRules_1.OUTREACH_MESSAGE : '',
            rows: group.rows,
        };
    }).sort(function (a, b) { return b.incidents - a.incidents || a.client.localeCompare(b.client); });
}
async function expectationsView(user, isAdmin, part) {
    const state = await (0, expectationsStore_1.loadBoard)();
    const email = emailOf(user.email);
    const today = (0, expectationsRules_1.denverDate)(new Date());
    if (part === 'email') {
        const checks = state.emailChecks[today] || {};
        return { checked: !!checks[email], day: today };
    }
    const month = (0, expectationsRules_1.denverMonth)(new Date());
    const trips = own(state.trips, email, isAdmin);
    const mine = monthTrips(state.trips, email, month);
    const miles = mine.reduce(function (sum, trip) { return sum + trip.miles; }, 0);
    const hours = Number(state.hours[email + '|' + month] || 0);
    const ratio = (0, expectationsRules_1.mileageRatio)(miles, hours);
    let notes = { late: [], evals: [], reports: [] };
    let notesError = '';
    try {
        notes = await notesFor(email, isAdmin, new Date());
    }
    catch (err) {
        notesError = err instanceof Error ? err.message.slice(0, 180) : 'IntakeQ could not be read';
    }
    let clients = [];
    let cancellationsError = '';
    try {
        clients = await cancellationsFor(email, isAdmin, state, today);
    }
    catch (err) {
        cancellationsError = err instanceof Error ? err.message.slice(0, 180) : 'Cancellations could not be read';
    }
    const followUps = state.followUps.filter(function (item) { return isAdmin || item.staffEmail === email; });
    return {
        role: isAdmin ? 'admin' : 'provider',
        me: email,
        day: today,
        month,
        zoneMiles: 5,
        measure: 'straight-line',
        home: publicHome(state, email),
        providers: await providerChoices(state, email, isAdmin),
        homes: isAdmin ? Object.keys(state.homes).map(function (key) {
            return { email: key, address: state.homes[key].address, mapLink: state.homes[key].mapLink, name: state.homes[key].name };
        }) : [],
        trips,
        monthMiles: miles,
        monthHours: hours,
        ratio,
        ratioHigh: (0, expectationsRules_1.ratioIsHigh)(ratio),
        requests: own(state.requests, email, isAdmin),
        emailChecked: !!(state.emailChecks[today] && state.emailChecks[today][email]),
        notes,
        notesError,
        clients,
        cancellationsError,
        followUps,
        outreachMessage: expectationsRules_1.OUTREACH_MESSAGE,
    };
}
function staffTarget(user, isAdmin, requested, requestedName = '') {
    if (isAdmin && requested)
        return { email: emailOf(requested), name: clip(requestedName, 120) };
    return { email: emailOf(user.email), name: user.name || user.email };
}
async function expectationsAction(user, isAdmin, body) {
    const action = clip(body.action, 40);
    const email = emailOf(user.email);
    if (action === 'expectations-home') {
        const address = clip(body.homeAddress, 180);
        const mapLink = clip(body.mapLink, 300);
        if (!address && !mapLink)
            throw new expectationsStore_1.ExpectationsStoreError('Add a map link or a home address.', 400);
        if (mapLink && !/^https:\/\//i.test(mapLink)) {
            throw new expectationsStore_1.ExpectationsStoreError('The map link must start with https://', 400);
        }
        let lat = null;
        let lon = null;
        if (address) {
            const point = await (0, expectationsZone_1.geocodeAddress)(address);
            if (!point)
                throw new expectationsStore_1.ExpectationsStoreError('That home address could not be found. Include the street, city, and ZIP.', 400);
            lat = point.lat;
            lon = point.lon;
        }
        const target = staffTarget(user, isAdmin, clip(body.staffEmail, 200), clip(body.staffName, 120));
        await (0, expectationsStore_1.updateBoard)(function (state) {
            const previous = state.homes[target.email];
            state.homes[target.email] = {
                address,
                lat,
                lon,
                mapLink,
                name: target.name || (previous && previous.name) || target.email,
            };
        });
        return { ok: true };
    }
    if (action === 'expectations-zone') {
        const clientAddress = clip(body.clientAddress, 180);
        const state = await (0, expectationsStore_1.loadBoard)();
        const target = staffTarget(user, isAdmin, clip(body.staffEmail, 200), clip(body.staffName, 120));
        const home = state.homes[target.email];
        if (!home || (!home.mapLink && !home.address)) {
            throw new expectationsStore_1.ExpectationsStoreError('Save a map link or a home address first.', 400);
        }
        const point = await (0, expectationsZone_1.geocodeAddress)(clientAddress);
        if (!point)
            throw new expectationsStore_1.ExpectationsStoreError('That client address could not be found. Include the street, city, and ZIP.', 400);
        const map = home.mapLink ? await (0, expectationsZone_1.readMapGeometry)(home.mapLink) : null;
        const pin = (map && map.points[0]) || null;
        const saved = home.address && Number.isFinite(home.lat) && Number.isFinite(home.lon)
            ? { lat: home.lat, lon: home.lon }
            : null;
        if (map && map.polygons.length) {
            const inside = map.polygons.some(function (ring) { return (0, expectationsRules_1.pointInRing)(point.lat, point.lon, ring); });
            const from = pin || saved;
            const miles = from ? (0, expectationsRules_1.haversineMiles)(from.lat, from.lon, point.lat, point.lon) : null;
            return {
                miles,
                outside: !inside,
                measure: 'map',
                detail: inside
                    ? 'Inside the area drawn on this provider’s map.'
                    : 'Outside the area drawn on this provider’s map.',
            };
        }
        const from = pin || saved;
        if (!from) {
            throw new expectationsStore_1.ExpectationsStoreError('This provider’s map could not be read. Add a home address, or share the Google My Map so anyone with the link can view it.', 400);
        }
        const miles = (0, expectationsRules_1.haversineMiles)(from.lat, from.lon, point.lat, point.lon);
        const fromMap = !!pin;
        return {
            miles,
            outside: (0, expectationsRules_1.outsideZone)(miles),
            measure: fromMap ? 'map-pin' : 'straight-line',
            detail: miles + ' straight-line miles from ' + (fromMap ? 'the pin on this provider’s map' : 'the saved home address') + '. '
                + ((0, expectationsRules_1.outsideZone)(miles) ? 'This address is outside the 5-mile range.' : 'This address is inside the 5-mile range.'),
        };
    }
    if (action === 'expectations-approval') {
        const clientAddress = clip(body.clientAddress, 180);
        const miles = roundMiles(body.miles);
        const onMap = clip(body.measure, 20) === 'map';
        if (!clientAddress)
            throw new expectationsStore_1.ExpectationsStoreError('The address is required.', 400);
        if (!onMap && !(0, expectationsRules_1.outsideZone)(miles))
            return { ok: true, outside: false };
        if (onMap && body.outside !== true)
            return { ok: true, outside: false };
        const target = staffTarget(user, isAdmin, clip(body.staffEmail, 200), clip(body.staffName, 120));
        const created = await (0, expectationsStore_1.updateBoard)(function (state) {
            const existing = state.requests.find(function (item) {
                return item.staffEmail === target.email && item.status === 'pending' && item.clientAddress.toLowerCase() === clientAddress.toLowerCase();
            });
            if (existing)
                return;
            state.requests.push({
                id: (0, expectationsStore_1.newId)('trip'),
                staffEmail: target.email,
                staffName: target.name || (state.homes[target.email] && state.homes[target.email].name) || target.email,
                clientAddress,
                miles: Number.isFinite(miles) ? miles : 0,
                status: 'pending',
                decidedBy: '',
                createdAt: new Date().toISOString(),
            });
        });
        const pending = created.requests.filter(function (item) { return item.staffEmail === target.email && item.status === 'pending'; });
        return { ok: true, outside: true, requests: pending };
    }
    if (action === 'expectations-decision') {
        if (!isAdmin)
            throw new expectationsStore_1.ExpectationsStoreError('An administrator approves or denies travel.', 403);
        const id = clip(body.id, 40);
        const decision = clip(body.decision, 20);
        if (decision !== 'approved' && decision !== 'denied')
            throw new expectationsStore_1.ExpectationsStoreError('Choose approve or deny.', 400);
        let found = false;
        await (0, expectationsStore_1.updateBoard)(function (state) {
            state.requests.forEach(function (item) {
                if (item.id !== id)
                    return;
                item.status = decision;
                item.decidedBy = email;
                found = true;
            });
        });
        if (!found)
            throw new expectationsStore_1.ExpectationsStoreError('That request is no longer on the list.', 404);
        return { ok: true };
    }
    if (action === 'expectations-trip') {
        const target = staffTarget(user, isAdmin, clip(body.staffEmail, 200));
        const id = clip(body.id, 40);
        if (body.remove === true) {
            await (0, expectationsStore_1.updateBoard)(function (state) {
                state.trips = state.trips.filter(function (trip) {
                    if (trip.id !== id)
                        return true;
                    return !isAdmin && trip.staffEmail !== email;
                });
            });
            return { ok: true };
        }
        const miles = roundMiles(body.miles);
        const date = clip(body.date, 10);
        const startPlace = clip(body.startPlace, 160);
        const endPlace = clip(body.endPlace, 160);
        if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !startPlace || !endPlace || !Number.isFinite(miles)) {
            throw new expectationsStore_1.ExpectationsStoreError('Date, starting location, ending location, and miles are required.', 400);
        }
        await (0, expectationsStore_1.updateBoard)(function (state) {
            const next = {
                id: id || (0, expectationsStore_1.newId)('mile'),
                staffEmail: target.email,
                staffName: target.name || (state.homes[target.email] && state.homes[target.email].name) || target.email,
                date,
                startPlace,
                endPlace,
                miles,
            };
            const index = state.trips.findIndex(function (trip) { return trip.id === next.id; });
            if (index >= 0) {
                if (!isAdmin && state.trips[index].staffEmail !== email)
                    return;
                next.staffEmail = state.trips[index].staffEmail;
                next.staffName = state.trips[index].staffName;
                state.trips[index] = next;
            }
            else
                state.trips.push(next);
        });
        return { ok: true };
    }
    if (action === 'expectations-hours') {
        const target = staffTarget(user, isAdmin, clip(body.staffEmail, 200));
        const hours = Number(body.hours);
        const month = clip(body.month, 7) || (0, expectationsRules_1.denverMonth)(new Date());
        if (!/^\d{4}-\d{2}$/.test(month) || !Number.isFinite(hours) || hours < 0 || hours > 400) {
            throw new expectationsStore_1.ExpectationsStoreError('Enter the hours worked this month.', 400);
        }
        await (0, expectationsStore_1.updateBoard)(function (state) {
            state.hours[target.email + '|' + month] = Math.round(hours * 10) / 10;
        });
        return { ok: true };
    }
    if (action === 'expectations-copy') {
        const target = staffTarget(user, isAdmin, clip(body.staffEmail, 200));
        const state = await (0, expectationsStore_1.loadBoard)();
        const month = (0, expectationsRules_1.denverMonth)(new Date());
        const trips = monthTrips(state.trips, target.email, month);
        const hours = Number(state.hours[target.email + '|' + month] || 0);
        const home = state.homes[target.email];
        const saved = await (0, mileageWorkbook_1.writeMileageDashboardCopy)(home ? home.name : target.name, target.email, trips, hours);
        if (!saved.ok)
            throw new expectationsStore_1.ExpectationsStoreError(saved.error || 'The read-only copy could not be saved.', 400);
        return { ok: true, path: saved.path || '' };
    }
    if (action === 'expectations-email') {
        const today = (0, expectationsRules_1.denverDate)(new Date());
        const checked = body.checked !== false;
        await (0, expectationsStore_1.updateBoard)(function (state) {
            const day = state.emailChecks[today] || {};
            day[email] = checked;
            state.emailChecks[today] = day;
        });
        return { ok: true, checked, day: today };
    }
    if (action === 'expectations-frequency') {
        if (!isAdmin)
            throw new expectationsStore_1.ExpectationsStoreError('An administrator sets how often the visits are scheduled.', 403);
        const key = clip(body.key, 160);
        if (!key)
            throw new expectationsStore_1.ExpectationsStoreError('Choose a client first.', 400);
        const frequency = frequencyOf(body.frequency);
        await (0, expectationsStore_1.updateBoard)(function (state) {
            state.frequency[key] = frequency;
        });
        return { ok: true };
    }
    if (action === 'expectations-followup') {
        if (!isAdmin)
            throw new expectationsStore_1.ExpectationsStoreError('An administrator records the UNICORN or PAR follow-up.', 403);
        const noteId = clip(body.noteId, 80);
        const followUp = clip(body.followUp, 20);
        if (!noteId || (followUp !== 'unicorn' && followUp !== 'par' && followUp !== 'done')) {
            throw new expectationsStore_1.ExpectationsStoreError('Choose UNICORN, PAR fax, or done.', 400);
        }
        const kind = clip(body.kind, 20) === 'report' ? 'report' : 'eval';
        await (0, expectationsStore_1.updateBoard)(function (state) {
            const next = {
                noteId,
                staffEmail: emailOf(clip(body.staffEmail, 200)),
                noteName: clip(body.noteName, 120),
                date: clip(body.date, 40),
                kind,
                followUp,
            };
            const index = state.followUps.findIndex(function (item) { return item.noteId === noteId; });
            if (index >= 0)
                state.followUps[index] = next;
            else
                state.followUps.push(next);
        });
        return { ok: true };
    }
    if (action === 'expectations-outreach') {
        if (!isAdmin)
            throw new expectationsStore_1.ExpectationsStoreError('An administrator records the email or the call.', 403);
        const key = clip(body.key, 160);
        const choice = clip(body.choice, 20);
        if (!key || (choice !== 'email' && choice !== 'call'))
            throw new expectationsStore_1.ExpectationsStoreError('Choose email or a call.', 400);
        await (0, expectationsStore_1.updateBoard)(function (state) {
            state.outreach[key] = { choice, at: new Date().toISOString(), by: email };
        });
        return { ok: true, message: choice === 'email' ? expectationsRules_1.OUTREACH_MESSAGE : '' };
    }
    if (action === 'expectations-sync') {
        const view = await expectationsView(user, true, 'expectations');
        const due = Array.isArray(view.clients) ? view.clients.filter(function (item) {
            return item.needsOutreach && !item.outreach;
        }) : [];
        const late = view.notes && typeof view.notes === 'object' ? view.notes.late || [] : [];
        return {
            ok: true,
            day: view.day,
            lateNotes: Array.isArray(late) ? late.length : 0,
            outreachDue: due.length,
            messages: due.map(function (item) { return { staffEmail: item.staffEmail || '', client: item.client || '', message: item.message || expectationsRules_1.OUTREACH_MESSAGE }; }),
        };
    }
    throw new expectationsStore_1.ExpectationsStoreError('That mileage action is not recognized.', 400);
}
function expectationsError(err) {
    if (err instanceof expectationsStore_1.ExpectationsStoreError)
        return { status: err.status, message: err.message };
    const message = err instanceof Error ? err.message : 'Request failed';
    return { status: 500, message: message.slice(0, 180) };
}
function pendingRequest(requests, email, address) {
    return requests.find(function (item) {
        return item.staffEmail === email && item.status === 'pending' && item.clientAddress.toLowerCase() === address.toLowerCase();
    });
}
