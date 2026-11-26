"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.REFERRAL_SOURCES = exports.PRIVATE_FIRST_MESSAGE = exports.SC_NO_RESPONSE_MESSAGE = exports.EI_LAST_MESSAGE = exports.EI_FIRST_EMAIL = void 0;
exports.parseIsoDate = parseIsoDate;
exports.addDays = addDays;
exports.mondayOf = mondayOf;
exports.workingDaysBetween = workingDaysBetween;
exports.calendarDaysAfter = calendarDaysAfter;
exports.sourceByValue = sourceByValue;
exports.shouldSitAtBottom = shouldSitAtBottom;
exports.buildNewReferral = buildNewReferral;
exports.applyReferralAction = applyReferralAction;
exports.projectReferral = projectReferral;
exports.summarizeOpenings = summarizeOpenings;
exports.memoryList = memoryList;
exports.handleReferrals = handleReferrals;
const storage_blob_1 = require("@azure/storage-blob");
const crypto_1 = require("crypto");
const config_1 = require("./config");
const http_1 = require("./http");
const schedulingService_1 = require("./schedulingService");
const sharepointService_1 = require("./sharepointService");
/**
 * Referral checklist, message drafts, and opening counts.
 * The SharePoint list Referrals still holds the family name, dates, status,
 * assigned staff, and notes. The checklist lives beside it in blob storage
 * (referrals/records.json) because that list has no columns for steps.
 * Schedule openings are read live. They are not copied into this file.
 * Nothing here sends a message.
 */
const BLOB_NAME = 'referrals/records.json';
const TIME_ZONE = 'America/Denver';
const OPENING_GOAL_MAX = 20;
const DAY_GOAL = 28;
const RESPONSE_DAYS = 15;
const OUTREACHES_BEFORE_BOTTOM = 2;
const OWNER = 'Administrator';
const SARA_BETH = 'Sara Beth';
const PAPERWORK_LINE = 'Please contact our referrals team at referrals@emmeco.org with any paperwork or questions.';
exports.EI_FIRST_EMAIL = 'Hi Sam! SB & Jenny here from EMME. I received your information from your service coordinator with RMHS';
exports.EI_LAST_MESSAGE = 'Hi hi Jenny, out of respect for you and your family today will be the last day I reach out to you. I will need to hear from you by end of today or I can let your service coordinator know there has been no response. For the record, we would love to support you and your family!';
exports.SC_NO_RESPONSE_MESSAGE = 'Hey, haven\'t heard a response, here is the phone number, can you verify and check in if they want services?';
exports.PRIVATE_FIRST_MESSAGE = 'Hi Jenny! Sara Beth & Jenny from EMME Therapy Services here. It looks like you reached out to us on our website for Speech and OT services. Do you have time to connect with us this week to talk about needs and next steps?';
const THERAPY_IDS = ['slp', 'ot', 'pt', 'psych', 'other'];
const THERAPY_LABELS = {
    slp: 'Speech',
    ot: 'Occupational therapy',
    pt: 'Physical therapy',
    psych: 'Psychology',
    other: 'Other',
};
const EI_STEPS = [
    { id: 'call', label: 'Call' },
    { id: 'mid', label: 'MID (Most Important Document), and put Medicaid and insurance on the PAR/HIA form' },
    { id: 'provider', label: 'Provider notification' },
    { id: 'welcome', label: 'Welcome email' },
    { id: 'sc', label: 'Service coordinator notified' },
    { id: 'days', label: 'Record the days it takes to place' },
    { id: 'remove', label: 'Remove from the referral list' },
];
exports.REFERRAL_SOURCES = [
    { id: 'ai', label: 'A&I Avenues', how: 'Early Intervention Colorado website', note: 'Log in and enter provider availability. Check often so a referral sent straight to a provider is not missed.', paperwork: true },
    { id: 'blue-peaks', label: 'Blue Peaks', how: 'Emailed to the inbox', note: 'Try to accept these when you can. They serve rural areas with few providers. Tracie Lopez is the coordinator who works with Blue Peaks.', paperwork: true },
    { id: 'cci', label: 'CCI', how: 'Emailed to the inbox', note: 'These are uncommon. Scott oversees CCI. Watch the inbox.', paperwork: true },
    { id: 'blue-sky', label: 'Colorado Blue Sky', how: 'Online call-out sheet', note: 'Check the call-out sheet daily. Enter the provider name, referrals@emmeco.org, and availability.', paperwork: true },
    { id: 'cornerstone', label: 'Cornerstone', how: 'Emailed to the inbox', note: 'Try to accept these when you can. They serve rural areas with few providers.', paperwork: true },
    { id: 'ddrc', label: 'DDRC', how: 'Referral lists by email', note: 'We rarely accept these. Only a few providers are credentialed. Set the email aside unless Kelly, Sara Beth, or Jenny are credentialed to pick up children through DDRC.', paperwork: true },
    { id: 'dp', label: 'Developmental Pathways', how: 'Emailed one at a time', note: 'These expire about 48 hours after they arrive. If we cannot accept one, reply and say we are unavailable. Everyone is credentialed with Developmental Pathways except Kellie, Kyla, Linda, and Dominique.', paperwork: true },
    { id: 'ecsdd', label: 'ECSDD', how: 'Emailed to the inbox', note: 'Watch the inbox and reply with provider availability.', paperwork: true },
    { id: 'envision', label: 'Envision', how: 'Online call-out sheet', note: 'Use the call-out sheet and list referrals@emmeco.org as the contact email.', paperwork: true },
    { id: 'foothills', label: 'Foothills Gateway', how: 'Online call-out sheet', note: 'Watch the call-out sheet and list the referrals email as the contact.', paperwork: true },
    { id: 'horizons', label: 'Horizons', how: 'Emailed to the inbox', note: 'Watch the inbox and reply with provider availability.', paperwork: true },
    { id: 'inspiration', label: 'Inspiration Field', how: 'Emailed to the inbox', note: 'Watch the inbox and reply with provider availability.', paperwork: true },
    { id: 'mvds', label: 'MVDS', how: 'Emailed to the inbox', note: 'Watch the inbox and reply with provider availability.', paperwork: true },
    { id: 'north-metro', label: 'North Metro', how: 'All-call emails', note: 'Review the all-call emails and reply with the providers who are available.', paperwork: true },
    { id: 'rmhs', label: 'RMHS', how: 'Online referral spreadsheet', note: 'Track every referral on the RMHS spreadsheet. Follow up every two days until it is placed and the paperwork is received, and keep following up after placement until the paperwork arrives.', paperwork: true },
    { id: 'website', label: 'Website or a family called', how: 'Phone, voicemail, or the website', note: 'This is a direct intake, not a broker placement. Collect the child\'s information, the pediatrician, and insurance, then start intake. Phone calls go on the call sheet for Sara Beth.', paperwork: false },
    { id: 'other', label: 'Other', how: 'Not listed', note: 'The written list does not name this source. Keep the name in the notes.', paperwork: false },
];
const PRIVATE_STEPS = [
    { id: 'first', label: 'First contact' },
    { id: 'needs', label: 'Needs discussion' },
    { id: 'placement', label: 'Placement' },
    { id: 'follow', label: 'Follow up' },
    { id: 'removal', label: 'Removal' },
];
const EI_DRAFTS = [
    { id: 'ei-first', label: 'First message to the Early Intervention family', body: exports.EI_FIRST_EMAIL },
    { id: 'ei-last', label: 'Last message to a family who has not answered', body: exports.EI_LAST_MESSAGE },
    { id: 'sc-no-response', label: 'Message to the service coordinator when the family has not answered', body: exports.SC_NO_RESPONSE_MESSAGE },
];
const PRIVATE_DRAFTS = [
    { id: 'private-first', label: 'First message to the private family', body: exports.PRIVATE_FIRST_MESSAGE },
];
const ASSIGN_DRAFTS = [
    { id: 'provider-notice', label: 'Provider notification', body: '' },
    { id: 'welcome', label: 'Welcome email', body: '' },
    { id: 'sc-assigned', label: 'Tell the service coordinator this family has a staff member', body: '' },
];
const sharePointList = {
    list: function () { return (0, sharepointService_1.listItems)(sharepointService_1.LISTS.referrals); },
    create: function (fields) { return (0, sharepointService_1.createItem)(sharepointService_1.LISTS.referrals, fields); },
    update: function (id, fields) { return (0, sharepointService_1.updateItem)(sharepointService_1.LISTS.referrals, id, fields); },
};
let writeChain = Promise.resolve();
function locked(work) {
    const run = writeChain.then(work, work);
    writeChain = run.then(function () { return undefined; }, function () { return undefined; });
    return run;
}
function denverToday(now = new Date()) {
    return new Intl.DateTimeFormat('en-CA', {
        timeZone: TIME_ZONE,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
    }).format(now);
}
function parseIsoDate(iso) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(iso))
        return '';
    const parts = iso.split('-').map(Number);
    const dt = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
    if (dt.getUTCFullYear() !== parts[0] || dt.getUTCMonth() !== parts[1] - 1 || dt.getUTCDate() !== parts[2])
        return '';
    return iso;
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
/** Working days after the start date, through the end date. Weekends are skipped. Holidays are not. */
function workingDaysBetween(start, end) {
    const from = parseIsoDate(start);
    const to = parseIsoDate(end);
    if (!from || !to || to < from)
        return null;
    let count = 0;
    let cursor = addDays(from, 1);
    while (cursor <= to) {
        const parts = cursor.split('-').map(Number);
        const day = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2])).getUTCDay();
        if (day !== 0 && day !== 6)
            count += 1;
        cursor = addDays(cursor, 1);
    }
    return count;
}
/** Calendar days after the start date, through the end date. */
function calendarDaysAfter(start, end) {
    const from = parseIsoDate(start);
    const to = parseIsoDate(end);
    if (!from || !to || to < from)
        return null;
    let count = 0;
    let cursor = addDays(from, 1);
    while (cursor <= to) {
        count += 1;
        cursor = addDays(cursor, 1);
    }
    return count;
}
function sourceByValue(value) {
    const text = value.trim().toLowerCase();
    if (!text)
        return undefined;
    return exports.REFERRAL_SOURCES.find(function (item) {
        return item.id === text || item.label.toLowerCase() === text;
    });
}
function normalizeSource(value) {
    const found = sourceByValue(value);
    return found ? found.id : value.trim().slice(0, 200);
}
function actorName(actor) {
    return (0, http_1.asText)(actor.name, 120) || (0, http_1.asText)(actor.email, 120) || OWNER;
}
function cleanTherapies(value) {
    const list = Array.isArray(value) ? value : [];
    const picked = [];
    list.forEach(function (item) {
        const id = (0, http_1.asText)(item, 20);
        if (THERAPY_IDS.indexOf(id) !== -1 && picked.indexOf(id) === -1)
            picked.push(id);
    });
    return picked;
}
function cleanKind(value) {
    const kind = (0, http_1.asText)(value, 20);
    if (!kind)
        return '';
    if (kind === 'ei' || kind === 'private')
        return kind;
    return null;
}
function templateBody(id) {
    const all = EI_DRAFTS.concat(PRIVATE_DRAFTS, ASSIGN_DRAFTS);
    const found = all.find(function (item) { return item.id === id; });
    return found ? found.body : '';
}
function refreshDraft(draft) {
    const body = draft.body.trim();
    if (!body)
        draft.incomplete = true;
    else if (draft.id === 'ei-first')
        draft.incomplete = body === exports.EI_FIRST_EMAIL;
    else
        draft.incomplete = false;
    if (draft.status !== 'approved') {
        draft.status = 'draft';
        draft.approvedBy = '';
        draft.approvedOn = '';
    }
}
function makeDraft(seed) {
    const draft = {
        id: seed.id,
        label: seed.label,
        body: seed.body,
        incomplete: false,
        status: 'draft',
        approvedBy: '',
        approvedOn: '',
    };
    refreshDraft(draft);
    return draft;
}
function makeSteps(kind) {
    const source = kind === 'ei' ? EI_STEPS : PRIVATE_STEPS;
    return source.map(function (step) {
        return { id: step.id, label: step.label, owner: OWNER, doneOn: '' };
    });
}
function blankDetail(id, familyName, today) {
    return {
        id: id,
        familyName: familyName,
        kind: '',
        status: 'New',
        serviceCoordinator: '',
        assignedTo: '',
        assignedEmail: '',
        notes: '',
        receivedOn: today,
        lastContactOn: '',
        serviceGridOn: '',
        placedOn: '',
        removedOn: '',
        source: '',
        outreachDates: [],
        therapies: [],
        steps: [],
        drafts: [],
        todos: [],
        updatedAt: '',
        updatedBy: '',
    };
}
function sharePointFields(detail) {
    return {
        Title: (detail.familyName || 'Referral').slice(0, 250),
        ReferrerName: detail.familyName.slice(0, 255),
        ReceivedOn: detail.receivedOn,
        Status: detail.status || 'Active',
        AssignedTo: detail.assignedTo.slice(0, 255),
        Notes: detail.notes.slice(0, 4000),
    };
}
function textField(fields, key) {
    return (0, http_1.asText)(fields[key], 4000);
}
function fromSharePoint(item, today) {
    const fields = item.fields || {};
    const detail = blankDetail(item.id, textField(fields, 'ReferrerName') || textField(fields, 'Title'), today);
    detail.notes = textField(fields, 'Notes');
    detail.assignedTo = textField(fields, 'AssignedTo').slice(0, 200);
    detail.status = textField(fields, 'Status') || 'New';
    const received = parseIsoDate(textField(fields, 'ReceivedOn').slice(0, 10));
    detail.receivedOn = received || textField(fields, 'ReceivedOn').slice(0, 40) || today;
    return detail;
}
function coerceStored(raw, id, today) {
    if (!raw || typeof raw !== 'object')
        return null;
    const row = raw;
    const familyName = (0, http_1.asText)(row.familyName, 200);
    if (!familyName)
        return null;
    const kind = cleanKind(row.kind);
    const detail = blankDetail(id, familyName, today);
    detail.kind = kind || '';
    detail.status = (0, http_1.asText)(row.status, 40) || 'Active';
    detail.serviceCoordinator = (0, http_1.asText)(row.serviceCoordinator, 200);
    detail.assignedTo = (0, http_1.asText)(row.assignedTo, 200);
    detail.assignedEmail = (0, http_1.asText)(row.assignedEmail, 200);
    detail.notes = (0, http_1.asText)(row.notes, 4000);
    detail.receivedOn = parseIsoDate((0, http_1.asText)(row.receivedOn, 40)) || (0, http_1.asText)(row.receivedOn, 40) || today;
    detail.lastContactOn = parseIsoDate((0, http_1.asText)(row.lastContactOn, 40));
    detail.serviceGridOn = parseIsoDate((0, http_1.asText)(row.serviceGridOn, 40));
    detail.placedOn = parseIsoDate((0, http_1.asText)(row.placedOn, 40));
    detail.removedOn = parseIsoDate((0, http_1.asText)(row.removedOn, 40));
    detail.source = normalizeSource((0, http_1.asText)(row.source, 200));
    detail.outreachDates = Array.isArray(row.outreachDates) ? row.outreachDates.map(function (item) {
        return parseIsoDate((0, http_1.asText)(item, 40));
    }).filter(function (item) { return !!item; }).sort() : [];
    detail.therapies = cleanTherapies(row.therapies);
    detail.steps = Array.isArray(row.steps) ? row.steps.map(function (step) {
        return {
            id: (0, http_1.asText)(step && step.id, 40),
            label: (0, http_1.asText)(step && step.label, 200),
            owner: (0, http_1.asText)(step && step.owner, 120) || OWNER,
            doneOn: parseIsoDate((0, http_1.asText)(step && step.doneOn, 40)),
        };
    }).filter(function (step) { return !!step.id; }) : [];
    detail.drafts = Array.isArray(row.drafts) ? row.drafts.map(function (draft) {
        const status = draft && draft.status === 'approved' ? 'approved' : 'draft';
        const next = {
            id: (0, http_1.asText)(draft && draft.id, 40),
            label: (0, http_1.asText)(draft && draft.label, 200),
            body: (0, http_1.asText)(draft && draft.body, 4000),
            incomplete: false,
            status: status,
            approvedBy: status === 'approved' ? (0, http_1.asText)(draft && draft.approvedBy, 120) : '',
            approvedOn: status === 'approved' ? parseIsoDate((0, http_1.asText)(draft && draft.approvedOn, 40)) : '',
        };
        refreshDraft(next);
        return next;
    }).filter(function (draft) { return !!draft.id; }) : [];
    detail.todos = Array.isArray(row.todos) ? row.todos.map(function (todo) {
        return {
            id: (0, http_1.asText)(todo && todo.id, 80),
            title: (0, http_1.asText)(todo && todo.title, 300),
            owner: (0, http_1.asText)(todo && todo.owner, 120) || OWNER,
            dueOn: parseIsoDate((0, http_1.asText)(todo && todo.dueOn, 40)),
            reminderOn: parseIsoDate((0, http_1.asText)(todo && todo.reminderOn, 40)),
            doneOn: parseIsoDate((0, http_1.asText)(todo && todo.doneOn, 40)),
        };
    }).filter(function (todo) { return !!todo.id && !!todo.title; }) : [];
    detail.updatedAt = (0, http_1.asText)(row.updatedAt, 40);
    detail.updatedBy = (0, http_1.asText)(row.updatedBy, 120);
    return detail;
}
function untouched(detail) {
    if (detail.steps.some(function (step) { return !!step.doneOn; }))
        return false;
    if (detail.todos.length)
        return false;
    return detail.drafts.every(function (draft) {
        return draft.status !== 'approved' && draft.body === templateBody(draft.id);
    });
}
function startProcess(detail, kind) {
    detail.kind = kind;
    detail.steps = makeSteps(kind);
    detail.drafts = (kind === 'ei' ? EI_DRAFTS : PRIVATE_DRAFTS).map(makeDraft);
    if (detail.status === 'New' || !detail.status)
        detail.status = 'Active';
}
function addTodo(detail, todo) {
    if (detail.todos.some(function (item) { return item.id === todo.id && !item.doneOn; }))
        return;
    const id = detail.todos.some(function (item) { return item.id === todo.id; })
        ? todo.id + '-' + String(detail.todos.length + 1)
        : todo.id;
    detail.todos.push({
        id: id,
        title: todo.title,
        owner: todo.owner || OWNER,
        dueOn: todo.dueOn,
        reminderOn: todo.dueOn,
        doneOn: '',
    });
}
function dropOpenSendTodo(detail, draftId) {
    detail.todos = detail.todos.filter(function (todo) {
        if (todo.doneOn)
            return true;
        return todo.id !== 'send-' + draftId && todo.id.indexOf('send-' + draftId + '-') !== 0;
    });
}
function ensureAssignWork(detail, actor, today) {
    ASSIGN_DRAFTS.forEach(function (seed) {
        if (!detail.drafts.some(function (draft) { return draft.id === seed.id; }))
            detail.drafts.push(makeDraft(seed));
    });
    const owner = actorName(actor);
    addTodo(detail, {
        id: 'assign-provider',
        title: 'Provider notification: write the message, approve it, and send it yourself.',
        owner: owner,
        dueOn: today,
    });
    addTodo(detail, {
        id: 'assign-welcome',
        title: 'Welcome email: write the message, approve it, and send it yourself.',
        owner: owner,
        dueOn: today,
    });
    addTodo(detail, {
        id: 'assign-sc',
        title: 'Tell the service coordinator this family has a staff member. This does not send by itself.',
        owner: owner,
        dueOn: today,
    });
}
function trackDaysTodo(detail, today) {
    if (detail.kind !== 'ei' || !detail.serviceGridOn || !detail.placedOn)
        return;
    addTodo(detail, {
        id: 'tracking-doc',
        title: 'Add the days to place on the referrals spreadsheet. This screen does not open that spreadsheet.',
        owner: OWNER,
        dueOn: today,
    });
}
function applyStatusFromDates(detail) {
    if (detail.removedOn || detail.status === 'Removed')
        return;
    if (detail.placedOn)
        detail.status = 'Placed';
    else if (detail.status === 'Placed')
        detail.status = 'Active';
}
function familyConfirmed(detail) {
    return detail.status === 'Answered' || detail.status === 'Placed' || detail.status === 'Removed' || !!detail.placedOn || !!detail.removedOn;
}
function shouldSitAtBottom(detail, today) {
    if (detail.kind !== 'ei' || familyConfirmed(detail))
        return false;
    if (detail.outreachDates.length < OUTREACHES_BEFORE_BOTTOM)
        return false;
    const first = detail.outreachDates.slice().sort()[0];
    if (!first)
        return false;
    return today > addDays(first, RESPONSE_DAYS);
}
function applyListPosition(detail, today) {
    if (familyConfirmed(detail))
        return;
    if (shouldSitAtBottom(detail, today))
        detail.status = 'Bottom';
    else if (detail.status === 'Bottom')
        detail.status = 'Active';
}
function ensureOnce(detail, todo) {
    const exists = detail.todos.some(function (item) {
        return item.id === todo.id || item.id.indexOf(todo.id + '-') === 0;
    });
    if (!exists)
        addTodo(detail, todo);
}
function openTodo(detail, prefix) {
    return detail.todos.find(function (item) {
        return !item.doneOn && (item.id === prefix || item.id.indexOf(prefix + '-') === 0);
    });
}
function refreshFollowUps(detail, today) {
    const source = sourceByValue(detail.source);
    if (source && source.id === 'rmhs' && detail.status !== 'Removed' && !detail.removedOn) {
        const title = 'RMHS: follow up on this referral. Do this every two days until it is placed and the paperwork is received, then keep following up until the paperwork arrives. This screen does not send the follow-up.';
        const open = openTodo(detail, 'rmhs-followup');
        const doneBefore = detail.todos.some(function (item) {
            return !!item.doneOn && (item.id === 'rmhs-followup' || item.id.indexOf('rmhs-followup-') === 0);
        });
        const anchor = detail.lastContactOn || detail.receivedOn || today;
        const due = doneBefore && !open ? addDays(today, 2) : addDays(anchor, 2);
        if (open) {
            if (!doneBefore) {
                open.dueOn = addDays(anchor, 2);
                open.reminderOn = open.dueOn;
            }
            open.title = title;
        }
        else {
            addTodo(detail, { id: 'rmhs-followup', title: title, owner: OWNER, dueOn: due });
        }
    }
    if (source && source.id === 'dp' && !familyConfirmed(detail)) {
        const due = detail.receivedOn ? addDays(detail.receivedOn, 2) : today;
        const title = 'Developmental Pathways referrals expire about 48 hours after they arrive. If we cannot accept this one, reply and say we are unavailable. This screen does not send that reply.';
        const open = openTodo(detail, 'dp-expire');
        if (open) {
            open.dueOn = due;
            open.reminderOn = due;
            open.title = title;
        }
        else {
            ensureOnce(detail, { id: 'dp-expire', title: title, owner: OWNER, dueOn: due });
        }
    }
    if (detail.kind === 'ei' && !familyConfirmed(detail) && detail.serviceGridOn) {
        const elapsed = workingDaysBetween(detail.serviceGridOn, today);
        if (elapsed != null && elapsed > DAY_GOAL) {
            ensureOnce(detail, {
                id: 'day-goal-up',
                title: 'The 28 working days from the service grid are up and the family has not confirmed. They can come off the list when you remove them.',
                owner: OWNER,
                dueOn: today,
            });
        }
    }
}
function settle(detail, today) {
    refreshFollowUps(detail, today);
    applyListPosition(detail, today);
}
function buildNewReferral(body, today) {
    const familyName = (0, http_1.asText)(body.familyName, 200) || (0, http_1.asText)(body.referrerName, 200);
    if (!familyName)
        return { ok: false, status: 400, error: 'Family name is required.' };
    const kind = cleanKind(body.kind);
    if (kind === null)
        return { ok: false, status: 400, error: 'Choose Early Intervention or private.' };
    const detail = blankDetail('', familyName, today);
    detail.serviceCoordinator = (0, http_1.asText)(body.serviceCoordinator, 200);
    detail.assignedTo = (0, http_1.asText)(body.assignedTo, 200);
    detail.notes = (0, http_1.asText)(body.notes, 4000);
    detail.source = normalizeSource((0, http_1.asText)(body.source, 200));
    detail.therapies = cleanTherapies(body.therapies);
    detail.receivedOn = parseIsoDate((0, http_1.asText)(body.receivedOn, 40)) || today;
    detail.serviceGridOn = parseIsoDate((0, http_1.asText)(body.serviceGridOn, 40));
    detail.lastContactOn = parseIsoDate((0, http_1.asText)(body.lastContactOn, 40));
    if (kind)
        startProcess(detail, kind);
    else
        detail.status = (0, http_1.asText)(body.status, 40) || 'New';
    trackDaysTodo(detail, today);
    settle(detail, today);
    return { ok: true, detail: detail };
}
function fail(status, error) {
    return { ok: false, status: status, error: error };
}
function applyReferralAction(detailIn, body, actor, today) {
    const detail = JSON.parse(JSON.stringify(detailIn));
    const action = (0, http_1.asText)(body.action, 40) || 'save';
    if (action === 'save') {
        const familyName = (0, http_1.asText)(body.familyName, 200);
        if (!familyName)
            return fail(400, 'Family name is required.');
        const kind = cleanKind(body.kind);
        if (kind === null)
            return fail(400, 'Choose Early Intervention or private.');
        if (kind !== detail.kind) {
            if (kind && !detail.kind && detail.steps.length === 0)
                startProcess(detail, kind);
            else if (kind && untouched(detail))
                startProcess(detail, kind);
            else if (kind !== detail.kind)
                return fail(400, 'Uncheck the steps and leave the messages as they started before switching Early Intervention and private.');
        }
        detail.familyName = familyName;
        detail.serviceCoordinator = (0, http_1.asText)(body.serviceCoordinator, 200);
        detail.notes = (0, http_1.asText)(body.notes, 4000);
        detail.source = normalizeSource((0, http_1.asText)(body.source, 200));
        detail.therapies = cleanTherapies(body.therapies);
        const received = parseIsoDate((0, http_1.asText)(body.receivedOn, 40));
        if ((0, http_1.asText)(body.receivedOn, 40) && !received)
            return fail(400, 'Received date must be a real date.');
        if (received)
            detail.receivedOn = received;
        const last = (0, http_1.asText)(body.lastContactOn, 40);
        if (last && !parseIsoDate(last))
            return fail(400, 'Last contact must be a real date.');
        detail.lastContactOn = parseIsoDate(last);
        const grid = (0, http_1.asText)(body.serviceGridOn, 40);
        if (grid && !parseIsoDate(grid))
            return fail(400, 'Service grid date must be a real date.');
        detail.serviceGridOn = parseIsoDate(grid);
        const placed = (0, http_1.asText)(body.placedOn, 40);
        if (placed && !parseIsoDate(placed))
            return fail(400, 'Placed date must be a real date.');
        detail.placedOn = parseIsoDate(placed);
        applyStatusFromDates(detail);
        trackDaysTodo(detail, today);
        settle(detail, today);
        return { ok: true, detail: detail };
    }
    if (action === 'outreach') {
        const typed = (0, http_1.asText)(body.on, 40);
        if (typed && !parseIsoDate(typed))
            return fail(400, 'Outreach date must be a real date.');
        const on = parseIsoDate(typed) || today;
        detail.outreachDates.push(on);
        detail.outreachDates.sort();
        if (!detail.lastContactOn || on > detail.lastContactOn)
            detail.lastContactOn = on;
        settle(detail, today);
        return { ok: true, detail: detail };
    }
    if (action === 'step') {
        const stepId = (0, http_1.asText)(body.stepId, 40);
        const step = detail.steps.find(function (item) { return item.id === stepId; });
        if (!step)
            return fail(400, 'That step is not on this referral.');
        if (Object.prototype.hasOwnProperty.call(body, 'owner'))
            step.owner = (0, http_1.asText)(body.owner, 120) || OWNER;
        if (Object.prototype.hasOwnProperty.call(body, 'doneOn')) {
            const doneOn = (0, http_1.asText)(body.doneOn, 40);
            if (doneOn && !parseIsoDate(doneOn))
                return fail(400, 'Step date must be a real date.');
            if (doneOn && stepId === 'days' && (!detail.serviceGridOn || !detail.placedOn)) {
                return fail(400, 'Enter the service grid date and the placed date, save them, then record the days.');
            }
            step.doneOn = parseIsoDate(doneOn);
            if (stepId === 'remove' || stepId === 'removal') {
                if (step.doneOn) {
                    detail.status = 'Removed';
                    detail.removedOn = step.doneOn;
                }
                else if (detail.status === 'Removed') {
                    detail.status = detail.placedOn ? 'Placed' : 'Active';
                    detail.removedOn = '';
                }
            }
            if (stepId === 'mid' && step.doneOn) {
                addTodo(detail, {
                    id: 'par-hia',
                    title: 'Put Medicaid and insurance on the PAR/HIA form. Then follow the prior authorization steps for that funding source. This screen does not fill out the form.',
                    owner: step.owner || OWNER,
                    dueOn: step.doneOn,
                });
            }
        }
        settle(detail, today);
        return { ok: true, detail: detail };
    }
    if (action === 'assign') {
        const assignedTo = (0, http_1.asText)(body.assignedTo, 200);
        const assignedEmail = (0, http_1.asText)(body.assignedEmail, 200).toLowerCase();
        if (!assignedTo)
            return fail(400, 'Choose a staff member.');
        if (assignedEmail && assignedEmail.indexOf('@') === -1)
            return fail(400, 'Staff email is not valid.');
        detail.assignedTo = assignedTo;
        detail.assignedEmail = assignedEmail;
        ensureAssignWork(detail, actor, today);
        settle(detail, today);
        return { ok: true, detail: detail };
    }
    if (action === 'answered') {
        const wasRemoved = detail.status === 'Removed' || !!detail.removedOn;
        if (wasRemoved) {
            detail.removedOn = '';
            detail.steps.forEach(function (step) {
                if (step.id === 'remove' || step.id === 'removal')
                    step.doneOn = '';
            });
        }
        detail.status = 'Answered';
        addTodo(detail, {
            id: 'to-sara-beth',
            title: wasRemoved
                ? 'This family answered after being removed, so they are back on the active list. Send them to Sara Beth for onboarding.'
                : 'Send this family to Sara Beth for onboarding.',
            owner: SARA_BETH,
            dueOn: today,
        });
        settle(detail, today);
        return { ok: true, detail: detail };
    }
    if (action === 'save-draft' || action === 'approve') {
        const draft = detail.drafts.find(function (item) { return item.id === (0, http_1.asText)(body.draftId, 40); });
        if (!draft)
            return fail(400, 'That message is not on this referral.');
        if (Object.prototype.hasOwnProperty.call(body, 'body')) {
            const nextBody = (0, http_1.asText)(body.body, 4000);
            if (nextBody !== draft.body && draft.status === 'approved') {
                draft.status = 'draft';
                draft.approvedBy = '';
                draft.approvedOn = '';
                dropOpenSendTodo(detail, draft.id);
            }
            draft.body = nextBody;
            refreshDraft(draft);
        }
        if (action === 'save-draft') {
            settle(detail, today);
            return { ok: true, detail: detail };
        }
        refreshDraft(draft);
        if (draft.incomplete)
            return fail(400, 'Finish this message before you approve it. Nothing is sent from this screen.');
        draft.status = 'approved';
        draft.approvedBy = actorName(actor);
        draft.approvedOn = today;
        addTodo(detail, {
            id: 'send-' + draft.id,
            title: 'Send this message yourself: ' + draft.label + '. Approving it does not send it.',
            owner: actorName(actor),
            dueOn: today,
        });
        settle(detail, today);
        return { ok: true, detail: detail };
    }
    if (action === 'todo-done') {
        const todo = detail.todos.find(function (item) { return item.id === (0, http_1.asText)(body.todoId, 80); });
        if (!todo)
            return fail(400, 'That to-do is not on this referral.');
        if (!todo.doneOn)
            todo.doneOn = today;
        settle(detail, today);
        return { ok: true, detail: detail };
    }
    if (action === 'add-todo') {
        const title = (0, http_1.asText)(body.title, 300);
        const owner = (0, http_1.asText)(body.owner, 120);
        const dueOn = parseIsoDate((0, http_1.asText)(body.dueOn, 40));
        if (!title || !owner || !dueOn)
            return fail(400, 'A to-do needs a task, an owner, and a due date.');
        detail.todos.push({
            id: 'custom-' + (0, crypto_1.randomBytes)(4).toString('hex'),
            title: title,
            owner: owner,
            dueOn: dueOn,
            reminderOn: dueOn,
            doneOn: '',
        });
        settle(detail, today);
        return { ok: true, detail: detail };
    }
    return fail(400, 'That action is not available.');
}
function projectReferral(detail, today, createdAt = '') {
    const placedDays = detail.serviceGridOn && detail.placedOn ? workingDaysBetween(detail.serviceGridOn, detail.placedOn) : null;
    const elapsed = detail.kind === 'ei' && detail.serviceGridOn ? workingDaysBetween(detail.serviceGridOn, today) : null;
    const daysToPlace = placedDays;
    const daysOverGoal = detail.kind === 'ei' && ((daysToPlace != null && daysToPlace > DAY_GOAL) ||
        (!detail.placedOn && elapsed != null && elapsed > DAY_GOAL));
    const source = sourceByValue(detail.source);
    const firstOutreach = detail.outreachDates.length ? detail.outreachDates.slice().sort()[0] : '';
    const respondBy = firstOutreach ? addDays(firstOutreach, RESPONSE_DAYS) : '';
    const responseDaysLeft = respondBy && today <= respondBy ? calendarDaysAfter(today, respondBy) : null;
    return Object.assign({}, detail, {
        createdAt: createdAt,
        daysToPlace: daysToPlace,
        daysOverGoal: daysOverGoal,
        elapsedWorkingDays: elapsed,
        remainingWorkingDays: elapsed == null ? null : DAY_GOAL - elapsed,
        stepsUnconfirmed: detail.kind === 'private',
        dayGoal: detail.kind === 'ei' ? DAY_GOAL : null,
        outreachCount: detail.outreachDates.length,
        outreachesBeforeBottom: OUTREACHES_BEFORE_BOTTOM,
        responseDays: RESPONSE_DAYS,
        firstOutreachOn: firstOutreach,
        respondBy: respondBy,
        responseDaysLeft: responseDaysLeft,
        atBottom: detail.status === 'Bottom',
        sourceLabel: source ? source.label : detail.source,
        sourceHow: source ? source.how : '',
        sourceNote: source ? source.note : '',
        paperworkLine: source && source.paperwork ? PAPERWORK_LINE : '',
        dpExpiresOn: source && source.id === 'dp' && detail.receivedOn ? addDays(detail.receivedOn, 2) : '',
        todos: detail.todos.map(function (todo) {
            return Object.assign({}, todo, { overdue: !todo.doneOn && !!todo.dueOn && todo.dueOn < today });
        }),
        ReferrerName: detail.familyName,
        Status: detail.status,
        AssignedTo: detail.assignedTo,
        ReceivedOn: detail.receivedOn,
        Notes: detail.notes,
        Title: detail.familyName,
    });
}
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
function summarizeOpenings(options) {
    const today = parseIsoDate(options.today) || options.today;
    const weekStart = mondayOf(today);
    const weekEnd = addDays(weekStart, 6);
    const through = addDays(today, 29);
    const kindByEmail = new Map();
    (options.kinds || []).forEach(function (kind) {
        kindByEmail.set(kind.email.trim().toLowerCase(), cleanTherapies(kind.therapies));
    });
    let total = 0;
    let weekTotal = 0;
    const monthCounts = new Map();
    const therapyCounts = new Map();
    const providers = [];
    const slots = [];
    (options.staff || []).forEach(function (person) {
        const email = (0, http_1.asText)(person.staffEmail, 200).toLowerCase();
        const therapyIds = kindByEmail.get(email) || [];
        const therapyLabels = therapyIds.length ? therapyIds.map(function (id) { return THERAPY_LABELS[id] || id; }) : ['Not set'];
        let next30 = 0;
        let week = 0;
        const months = new Map();
        (person.openings || []).forEach(function (opening) {
            const date = parseIsoDate(opening.date);
            if (!date || date < today || date > through)
                return;
            total += 1;
            next30 += 1;
            const month = date.slice(0, 7);
            monthCounts.set(month, (monthCounts.get(month) || 0) + 1);
            months.set(month, (months.get(month) || 0) + 1);
            if (date >= weekStart && date <= weekEnd) {
                weekTotal += 1;
                week += 1;
            }
            therapyLabels.forEach(function (label) {
                therapyCounts.set(label, (therapyCounts.get(label) || 0) + 1);
            });
            slots.push({
                staffName: person.staffName,
                staffEmail: email,
                therapies: therapyLabels,
                therapyIds: therapyIds,
                date: date,
                day: (0, http_1.asText)(opening.day, 20),
                time: (0, http_1.asText)(opening.time, 40),
                status: (0, http_1.asText)(opening.status, 80),
            });
        });
        providers.push({
            staffName: person.staffName || email,
            staffEmail: email,
            therapyIds: therapyIds,
            therapies: therapyLabels,
            week: week,
            next30: next30,
        });
    });
    providers.sort(function (a, b) { return String(a.staffName).localeCompare(String(b.staffName)); });
    slots.sort(function (a, b) {
        return String(a.date).localeCompare(String(b.date)) || String(a.time).localeCompare(String(b.time)) || String(a.staffName).localeCompare(String(b.staffName));
    });
    const byMonth = Array.from(monthCounts.keys()).sort().map(function (month) {
        const index = Number(month.slice(5, 7)) - 1;
        return { month: month, label: (MONTHS[index] || month) + ' ' + month.slice(0, 4), count: monthCounts.get(month) || 0 };
    });
    const byTherapy = Array.from(therapyCounts.keys()).sort().map(function (label) {
        return { label: label, count: therapyCounts.get(label) || 0 };
    });
    const disciplineCount = byTherapy.filter(function (row) { return row.label !== 'Not set' && row.count > 0; }).length;
    return {
        source: 'emme-scheduling-db',
        timeZone: TIME_ZONE,
        today: today,
        weekStart: weekStart,
        weekEnd: weekEnd,
        through: through,
        windowDays: 30,
        total: total,
        weekTotal: weekTotal,
        disciplineCount: disciplineCount,
        unlabeled: therapyCounts.get('Not set') || 0,
        goalMax: OPENING_GOAL_MAX,
        overGoal: total > OPENING_GOAL_MAX,
        byMonth: byMonth,
        byTherapy: byTherapy,
        providers: providers,
        slots: slots,
    };
}
function emptyStore() {
    return { version: 1, records: {}, kinds: [] };
}
function storageReady() {
    return !!(0, config_1.getConfig)().storageConnection;
}
function container() {
    const conn = (0, config_1.getConfig)().storageConnection;
    if (!conn)
        throw new Error('Referral storage is not configured.');
    return storage_blob_1.BlobServiceClient.fromConnectionString(conn).getContainerClient((0, config_1.getConfig)().storageContainer);
}
async function readStore() {
    if (!storageReady())
        return emptyStore();
    const blob = container().getBlockBlobClient(BLOB_NAME);
    if (!(await blob.exists()))
        return emptyStore();
    let parsed;
    try {
        parsed = JSON.parse((await blob.downloadToBuffer()).toString('utf8'));
    }
    catch (_) {
        throw new Error('Referral storage could not be read.');
    }
    if (!parsed || typeof parsed !== 'object')
        throw new Error('Referral storage could not be read.');
    const file = parsed;
    const records = file.records && typeof file.records === 'object' ? file.records : null;
    if (!records)
        throw new Error('Referral storage could not be read.');
    const kinds = Array.isArray(file.kinds) ? file.kinds.map(function (kind) {
        return {
            email: (0, http_1.asText)(kind && kind.email, 200).toLowerCase(),
            name: (0, http_1.asText)(kind && kind.name, 200),
            therapies: cleanTherapies(kind && kind.therapies),
        };
    }).filter(function (kind) { return kind.email.indexOf('@') !== -1; }) : [];
    return { version: 1, records: records, kinds: kinds };
}
async function writeStore(file) {
    const target = container();
    await target.createIfNotExists();
    const payload = JSON.stringify(file);
    await target.getBlockBlobClient(BLOB_NAME).upload(payload, Buffer.byteLength(payload), {
        blobHTTPHeaders: { blobContentType: 'application/json' },
    });
}
function mergeDetail(item, stored, today) {
    const saved = stored ? coerceStored(stored, item.id, today) : null;
    if (saved) {
        saved.id = item.id;
        return saved;
    }
    return fromSharePoint(item, today);
}
function validId(id) {
    return /^[A-Za-z0-9_-]{1,80}$/.test(id);
}
async function defaultAvailability() {
    const result = await (0, schedulingService_1.availabilityForWindow)(30);
    return {
        configured: result.configured,
        asOf: result.asOf,
        staff: result.staff.map(function (person) {
            return {
                staffEmail: person.staffEmail,
                staffName: person.staffName,
                openings: (person.summary && person.summary.openings) || [],
            };
        }),
    };
}
function stamp(detail, actor) {
    detail.updatedAt = new Date().toISOString();
    detail.updatedBy = actorName(actor);
}
function memoryList() {
    const rows = [];
    let next = 1;
    return {
        rows: rows,
        async list() {
            return rows.map(function (row) { return { id: row.id, createdAt: row.createdAt, fields: Object.assign({}, row.fields) }; });
        },
        async create(fields) {
            const row = { id: String(next++), createdAt: new Date().toISOString(), fields: Object.assign({}, fields) };
            rows.push(row);
            return { id: row.id, createdAt: row.createdAt, fields: Object.assign({}, row.fields) };
        },
        async update(id, fields) {
            const row = rows.find(function (item) { return item.id === id; });
            if (!row)
                throw new Error('Referral was not found.');
            row.fields = Object.assign({}, row.fields, fields);
        },
    };
}
async function handleReferrals(context, req, actor, deps) {
    const list = (deps && deps.list) || sharePointList;
    const today = (deps && deps.today) || denverToday();
    const verb = (0, http_1.methodOf)(req);
    const body = (0, http_1.readBody)(req);
    if (verb === 'GET' && (0, http_1.queryOf)(req, 'view') === 'openings') {
        const load = (deps && deps.loadAvailability) || defaultAvailability;
        const loaded = await load();
        if (!loaded.configured) {
            (0, http_1.json)(context, 503, { error: 'Scheduling storage is not configured. Set AZURE_STORAGE_CONNECTION_STRING.' });
            return;
        }
        const store = await readStore();
        const summaryToday = parseIsoDate(loaded.asOf || '') || today;
        (0, http_1.json)(context, 200, summarizeOpenings({ today: summaryToday, staff: loaded.staff || [], kinds: store.kinds }));
        return;
    }
    if (verb === 'GET') {
        const items = await list.list();
        const referrals = await locked(async function () {
            const store = await readStore();
            let dirty = false;
            const rows = items.map(function (item) {
                const stored = store.records[item.id];
                const detail = mergeDetail(item, stored, today);
                const statusBefore = detail.status;
                const todosBefore = JSON.stringify(detail.todos);
                settle(detail, today);
                if (stored && (detail.status !== statusBefore || JSON.stringify(detail.todos) !== todosBefore)) {
                    detail.updatedAt = stored.updatedAt;
                    store.records[item.id] = detail;
                    dirty = true;
                }
                return { detail: detail, createdAt: item.createdAt || '', statusBefore: statusBefore, hadRecord: !!stored };
            });
            if (dirty) {
                await writeStore(store);
                for (let i = 0; i < rows.length; i++) {
                    if (rows[i].hadRecord && rows[i].detail.status !== rows[i].statusBefore) {
                        await list.update(rows[i].detail.id, sharePointFields(rows[i].detail));
                    }
                }
            }
            return rows.map(function (row) { return projectReferral(row.detail, today, row.createdAt); });
        });
        referrals.sort(function (a, b) {
            const rank = function (status) {
                if (status === 'Removed')
                    return 3;
                if (status === 'Bottom')
                    return 2;
                return 1;
            };
            const byRank = rank(a.status) - rank(b.status);
            if (byRank)
                return byRank;
            const received = a.status === 'Bottom'
                ? String(a.receivedOn || '').localeCompare(String(b.receivedOn || ''))
                : String(b.receivedOn || '').localeCompare(String(a.receivedOn || ''));
            return received || String(a.familyName || '').localeCompare(String(b.familyName || ''));
        });
        (0, http_1.json)(context, 200, { referrals: referrals, today: today, timeZone: TIME_ZONE, sources: exports.REFERRAL_SOURCES });
        return;
    }
    if (verb !== 'POST') {
        (0, http_1.json)(context, 405, { error: 'Method not allowed' });
        return;
    }
    if (!storageReady()) {
        (0, http_1.json)(context, 503, { error: 'Referral storage is not configured. Set AZURE_STORAGE_CONNECTION_STRING.' });
        return;
    }
    const action = (0, http_1.asText)(body.action, 40);
    if (action === 'provider-kinds') {
        const incoming = Array.isArray(body.providers) ? body.providers : [];
        if (!incoming.length) {
            (0, http_1.json)(context, 400, { error: 'Choose a staff member and a kind of therapy.' });
            return;
        }
        const cleaned = [];
        for (let i = 0; i < incoming.length; i++) {
            const row = incoming[i];
            const email = (0, http_1.asText)(row && row.email, 200).toLowerCase();
            if (!email || email.indexOf('@') === -1) {
                (0, http_1.json)(context, 400, { error: 'Staff email is not valid.' });
                return;
            }
            cleaned.push({ email: email, name: (0, http_1.asText)(row && row.name, 200) || email, therapies: cleanTherapies(row && row.therapies) });
        }
        await locked(async function () {
            const store = await readStore();
            cleaned.forEach(function (kind) {
                const index = store.kinds.findIndex(function (item) { return item.email === kind.email; });
                if (index === -1)
                    store.kinds.push(kind);
                else
                    store.kinds[index] = kind;
            });
            await writeStore(store);
        });
        (0, http_1.json)(context, 200, { ok: true });
        return;
    }
    const id = (0, http_1.restParts)(req)[0] || (0, http_1.asText)(body.id, 80);
    if (!id) {
        if (action && action !== 'save') {
            (0, http_1.json)(context, 400, { error: 'Choose a referral.' });
            return;
        }
        const built = buildNewReferral(body, today);
        if (!built.ok) {
            (0, http_1.json)(context, built.status, { error: built.error });
            return;
        }
        const created = await list.create(sharePointFields(built.detail));
        built.detail.id = created.id;
        stamp(built.detail, actor);
        await locked(async function () {
            const store = await readStore();
            store.records[created.id] = built.detail;
            await writeStore(store);
        });
        (0, http_1.json)(context, 201, { referral: projectReferral(built.detail, today, created.createdAt || '') });
        return;
    }
    if (!validId(id)) {
        (0, http_1.json)(context, 400, { error: 'That referral was not found.' });
        return;
    }
    const outcome = await locked(async function () {
        const items = await list.list();
        const item = items.find(function (row) { return row.id === id; });
        if (!item)
            return { status: 404, error: 'That referral was not found.' };
        const store = await readStore();
        const current = mergeDetail(item, store.records[id], today);
        const base = (0, http_1.asText)(body.baseUpdatedAt, 40);
        if ((current.updatedAt || '') !== base) {
            return { status: 409, error: 'Someone else saved this referral. Reload the page, then make your changes again.' };
        }
        const applied = applyReferralAction(current, Object.assign({}, body, { action: action || 'save' }), actor, today);
        if (!applied.ok)
            return { status: applied.status, error: applied.error };
        stamp(applied.detail, actor);
        applied.detail.id = id;
        store.records[id] = applied.detail;
        await writeStore(store);
        await list.update(id, sharePointFields(applied.detail));
        return { detail: applied.detail, createdAt: item.createdAt || '' };
    });
    if ('error' in outcome) {
        (0, http_1.json)(context, outcome.status, { error: outcome.error });
        return;
    }
    (0, http_1.json)(context, 200, { referral: projectReferral(outcome.detail, today, outcome.createdAt) });
}
