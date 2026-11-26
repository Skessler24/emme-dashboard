"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.TIME_ZONE = void 0;
exports.blankMeeting = blankMeeting;
exports.blankMap = blankMap;
exports.cleanMap = cleanMap;
exports.copyFileName = copyFileName;
exports.matchingFolder = matchingFolder;
exports.validWhen = validWhen;
exports.monthStanding = monthStanding;
exports.plusHour = plusHour;
exports.clientVisitTimes = clientVisitTimes;
exports.buildGrowthMapPdf = buildGrowthMapPdf;
const schedulingService_1 = require("./schedulingService");
/**
 * The blank Our Reflective Practice Growth Map.
 * Filled-in examples from a scanned copy are never used as starter text.
 */
exports.TIME_ZONE = 'America/Denver';
const TOPICS = [
    { id: 'observe', title: 'Observation before interpretation', foundation: true, prompt: 'What did I actually see and hear?' },
    { id: 'connection', title: 'Connection and engagement', foundation: true, prompt: 'How did I join the child\'s interest?' },
    { id: 'regulation', title: 'Regulation and co-regulation', foundation: true, prompt: 'What changed in the interaction or setting?' },
    { id: 'family', title: 'Family partnership', foundation: true, prompt: 'Whose priorities shaped our plan? How did we adapt a plan that didn\'t work?' },
    { id: 'coaching', title: 'Caregiver coaching', foundation: true, prompt: 'Where did the caregiver reflect or choose?' },
    { id: 'reactions', title: 'Our reactions and assumptions', foundation: true, prompt: 'What did I feel or assume in this moment?' },
    { id: 'motivation', title: 'Motivation and learning', foundation: false, prompt: 'What made participation meaningful for the child?' },
    { id: 'communication', title: 'Communication and behavior', foundation: false, prompt: 'What might this communicate? What else?' },
    { id: 'sensory', title: 'Sensory and individual differences', foundation: false, prompt: 'Which specific demand might matter?' },
    { id: 'play', title: 'Play and development', foundation: false, prompt: 'What small challenge fits this child?' },
    { id: 'feeding', title: 'Feeding and mealtime', foundation: false, prompt: 'What can I explore within my scope?' },
    { id: 'reasoning', title: 'Clinical reasoning', foundation: false, prompt: 'What evidence could change my plan? How did I recognize when something needs to change and decide what to try next?' },
    { id: 'conversations', title: 'Difficult conversations', foundation: false, prompt: 'What needs clarity and curiosity?' },
    { id: 'collaboration', title: 'Collaboration and consultation', foundation: false, prompt: 'When and how do you involve others? Knowing when to implement a team?' },
    { id: 'culture', title: 'Cultural responsiveness and family perspectives', foundation: false, prompt: 'What have I learned about this family\'s values, culture, language, and experiences? How have I adapted my approach rather than assumed what works for them?' },
    { id: 'selected', title: 'Clinician-selected growth topic', foundation: false, prompt: 'What else would I like to explore or develop? When to ask for help?' },
];
const TEXT_MAX = 4000;
function clip(value) {
    return String(value == null ? '' : value).replace(/\u0000/g, '').trim().slice(0, TEXT_MAX);
}
const FOUNDATION_IDS = new Set(TOPICS.filter(function (topic) { return topic.foundation; }).map(function (topic) { return topic.id; }));
function blankMeeting() {
    return {
        id: 'm-' + Math.random().toString(16).slice(2, 10),
        date: '',
        foundationId: '',
        update: '',
        background: '',
        clip: '',
        questions: '',
        noticed: '',
        action: '',
        observe: '',
        revisit: '',
        growthDates: '',
    };
}
function blankMap(email, name) {
    return {
        email: email.trim().toLowerCase(),
        name: name || email,
        updatedAt: '',
        updatedBy: '',
        topics: TOPICS.map(function (topic) {
            return { id: topic.id, title: topic.title, foundation: topic.foundation, prompt: topic.prompt, e: '', p: '', r: '' };
        }),
        meetings: [blankMeeting()],
    };
}
function cleanMeeting(value) {
    const row = value && typeof value === 'object' ? value : {};
    const id = clip(row.id).replace(/[^a-zA-Z0-9-]/g, '').slice(0, 40);
    const meeting = blankMeeting();
    meeting.id = id || meeting.id;
    meeting.date = /^\d{4}-\d{2}-\d{2}$/.test(clip(row.date)) ? clip(row.date) : '';
    const foundationId = clip(row.foundationId);
    meeting.foundationId = FOUNDATION_IDS.has(foundationId) ? foundationId : '';
    meeting.update = clip(row.update);
    meeting.background = clip(row.background);
    meeting.clip = clip(row.clip);
    meeting.questions = clip(row.questions);
    meeting.noticed = clip(row.noticed);
    meeting.action = clip(row.action);
    meeting.observe = clip(row.observe);
    meeting.revisit = clip(row.revisit);
    meeting.growthDates = clip(row.growthDates);
    return meeting;
}
/** Keeps only the real map fields. Topic titles always come from the blank form. */
function cleanMap(email, name, incoming, updatedBy, updatedAt) {
    const src = incoming && typeof incoming === 'object' ? incoming : {};
    const byId = new Map();
    if (Array.isArray(src.topics)) {
        src.topics.forEach(function (item) {
            if (!item || typeof item !== 'object')
                return;
            const id = String(item.id || '');
            if (id)
                byId.set(id, item);
        });
    }
    const topics = TOPICS.map(function (topic) {
        const got = byId.get(topic.id) || {};
        return {
            id: topic.id,
            title: topic.title,
            foundation: topic.foundation,
            prompt: topic.prompt,
            e: clip(got.e),
            p: clip(got.p),
            r: clip(got.r),
        };
    });
    const incomingMeetings = Array.isArray(src.meetings) ? src.meetings.slice(0, 24) : [];
    const meetings = (incomingMeetings.length ? incomingMeetings : [blankMeeting()]).map(cleanMeeting);
    return {
        email: email.trim().toLowerCase(),
        name: clip(name).slice(0, 120) || email,
        updatedAt,
        updatedBy,
        topics,
        meetings,
    };
}
function copyFileName(name) {
    const clean = String(name || 'Provider').replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim() || 'Provider';
    return clean.slice(0, 80) + ' - Reflective Supervision.pdf';
}
function compact(value) {
    return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}
/** Picks the provider's folder inside RS Providers when the folder name matches the person. */
function matchingFolder(folders, providerName, email) {
    const want = compact(providerName);
    const local = compact(String(email || '').split('@')[0] || '');
    let best = null;
    let bestScore = 0;
    folders.forEach(function (folder) {
        const got = compact(folder.name);
        if (!got)
            return;
        let score = 0;
        if (want && got === want)
            score = 100;
        else if (want.length > 3 && got.length > 3 && (got.indexOf(want) !== -1 || want.indexOf(got) !== -1))
            score = 80;
        else if (local.length > 3 && (got.indexOf(local) !== -1 || local.indexOf(got) !== -1))
            score = 60;
        if (score > bestScore) {
            best = folder;
            bestScore = score;
        }
    });
    return best;
}
function validWhen(value) {
    if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(value))
        return false;
    const hh = Number(value.slice(11, 13));
    const mm = Number(value.slice(14, 16));
    return hh >= 0 && hh <= 23 && mm >= 0 && mm <= 59;
}
function previousMonth(yearMonth) {
    const parts = yearMonth.split('-').map(Number);
    const dt = new Date(Date.UTC(parts[0], parts[1] - 2, 1));
    return dt.toISOString().slice(0, 7);
}
function monthName(isoDate) {
    const parts = isoDate.slice(0, 10).split('-').map(Number);
    return new Intl.DateTimeFormat('en-US', { month: 'long', year: 'numeric', timeZone: 'UTC' }).format(new Date(Date.UTC(parts[0], parts[1] - 1, parts[2] || 1)));
}
function hoursBetween(start, end) {
    if (!validWhen(start) || !validWhen(end) || end <= start)
        return 0;
    const startDay = start.slice(0, 10).split('-').map(Number);
    const endDay = end.slice(0, 10).split('-').map(Number);
    const startClock = start.slice(11).split(':').map(Number);
    const endClock = end.slice(11).split(':').map(Number);
    const from = Date.UTC(startDay[0], startDay[1] - 1, startDay[2], startClock[0], startClock[1]);
    const to = Date.UTC(endDay[0], endDay[1] - 1, endDay[2], endClock[0], endClock[1]);
    return Math.round(((to - from) / 3600000) * 10) / 10;
}
/**
 * One confirmed reflective supervision session each calendar month, in Mountain Time.
 * Hours are the length of those confirmed sessions. There is no separate hour goal.
 */
function monthStanding(appointments, today) {
    const month = today.slice(0, 7);
    const prior = previousMonth(month);
    const name = monthName(today);
    const confirmed = appointments.filter(function (item) { return item.status === 'confirmed' && validWhen(item.start); });
    const thisMonth = confirmed.filter(function (item) { return item.start.slice(0, 7) === month; });
    const lastMonth = confirmed.filter(function (item) { return item.start.slice(0, 7) === prior; });
    const waiting = appointments.some(function (item) {
        return item.start.slice(0, 7) === month && (item.status === 'proposed' || item.status === 'suggested' || item.status === 'accepted');
    });
    const hours = Math.round(thisMonth.reduce(function (sum, item) { return sum + hoursBetween(item.start, item.end); }, 0) * 10) / 10;
    const done = thisMonth.length > 0;
    const behind = !done && !waiting && confirmed.length > 0 && lastMonth.length === 0;
    const hourText = hours === 1 ? '1 hour' : hours + ' hours';
    let short = 'Still needed';
    let label = 'Still needed in ' + name + '. One session is due each month.';
    if (done) {
        short = 'Done';
        label = 'Done for ' + name + ' (' + thisMonth.length + (thisMonth.length === 1 ? ' session, ' : ' sessions, ') + hourText + '). One session is due each month.';
    }
    else if (waiting) {
        short = 'Scheduled';
        label = 'A time is already on the list for ' + name + '. It counts once it is confirmed. One session is due each month.';
    }
    else if (behind) {
        short = 'Behind';
        label = 'Behind. No confirmed session last month or this month. One session is due each month.';
    }
    return { month, short, done, behind, sessions: thisMonth.length, hours, label };
}
/** Adds one hour on the clock, without changing the time zone. */
function plusHour(start) {
    const parts = start.split('T');
    const day = parts[0].split('-').map(Number);
    const clock = parts[1].split(':').map(Number);
    const dt = new Date(Date.UTC(day[0], day[1] - 1, day[2], clock[0], clock[1]));
    dt.setUTCMinutes(dt.getUTCMinutes() + 60);
    const p = function (n) { return String(n).padStart(2, '0'); };
    return dt.getUTCFullYear() + '-' + p(dt.getUTCMonth() + 1) + '-' + p(dt.getUTCDate()) + 'T' + p(dt.getUTCHours()) + ':' + p(dt.getUTCMinutes());
}
function addDays(iso, days) {
    const parts = iso.split('-').map(Number);
    const dt = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
    dt.setUTCDate(dt.getUTCDate() + days);
    return dt.toISOString().slice(0, 10);
}
function weekdayName(iso) {
    const parts = iso.split('-').map(Number);
    const dt = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
    return ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][dt.getUTCDay()];
}
function sameDay(slotDay, weekday) {
    const a = String(slotDay || '').trim().toLowerCase();
    const b = weekday.toLowerCase();
    if (!a)
        return false;
    return a === b || b.indexOf(a.slice(0, 3)) === 0;
}
/** Booked visit times from a saved schedule. Notes are left out. */
function clientVisitTimes(workspace, windowDays, asOf) {
    const board = (0, schedulingService_1.currentBoard)((workspace && workspace.boards) || undefined);
    const out = [];
    const days = Math.min(31, Math.max(1, windowDays));
    for (let offset = 0; offset < days; offset++) {
        const date = addDays(asOf, offset);
        const weekday = weekdayName(date);
        const day = ((board && board.days) || []).find(function (item) { return sameDay(String(item.name || ''), weekday); });
        ((day && day.slots) || []).forEach(function (slot) {
            const key = String(slot.status || '').trim().toLowerCase();
            if (key !== 'full' && key !== 'makeup')
                return;
            out.push({
                date,
                day: weekday,
                time: String(slot.time || ''),
                label: key === 'makeup' ? 'Makeup visit' : 'With a client',
            });
        });
    }
    return out;
}
function pdfEscape(value) {
    return value.replace(/[^\x20-\x7E]/g, ' ').replace(/\\/g, '\\\\').replace(/\(/g, '\\(').replace(/\)/g, '\\)');
}
function wrap(value) {
    const text = String(value || '').replace(/\s+/g, ' ').trim();
    if (!text)
        return [''];
    const words = text.split(' ');
    const lines = [];
    let line = '';
    words.forEach(function (word) {
        const next = line ? line + ' ' + word : word;
        if (next.length > 90 && line) {
            lines.push(line);
            line = word.slice(0, 90);
        }
        else {
            line = next.slice(0, 90);
        }
    });
    if (line)
        lines.push(line);
    return lines;
}
/** A plain PDF of the growth map. The bytes are the document, and they are not written to the log. */
function buildGrowthMapPdf(map) {
    const lines = [];
    const push = function (text) { wrap(text).forEach(function (line) { lines.push(line); }); };
    push('Our Reflective Practice Growth Map');
    push(map.name || map.email);
    push('E = explored: what I noticed or better understood. P = practiced: what I tried. R = revisited: what I learned.');
    push('The six F spaces are shared foundations. Other spaces follow the provider\'s needs, interests, and scope.');
    push('');
    map.topics.forEach(function (topic) {
        push((topic.foundation ? '(F) ' : '') + topic.title);
        push(topic.prompt);
        push('E: ' + (topic.e || '______') + '   P: ' + (topic.p || '______') + '   R: ' + (topic.r || '______'));
        push('');
    });
    map.meetings.forEach(function (meeting, index) {
        push('Reflective Supervision Preparation ' + (index + 1));
        push('Meeting date: ' + (meeting.date || ''));
        const chosen = map.topics.find(function (topic) { return topic.id === meeting.foundationId; });
        if (chosen)
            push('Foundation for this session: (F) ' + chosen.title + ' — ' + chosen.prompt);
        push('Brief update: ' + meeting.update);
        push('Background: ' + meeting.background);
        push('Clip timestamp: ' + meeting.clip);
        push('Questions and reflections: ' + meeting.questions);
        push('Something I noticed or better understood: ' + meeting.noticed);
        push('One action I will try and why: ' + meeting.action);
        push('What I will observe or ask the family: ' + meeting.observe);
        push('When will we revisit this: ' + meeting.revisit);
        push('Growth map dates to add: ' + meeting.growthDates);
        push('');
    });
    return pdfFromLines(lines);
}
function pdfFromLines(lines) {
    const pageSize = 46;
    const pages = [];
    for (let i = 0; i < lines.length; i += pageSize)
        pages.push(lines.slice(i, i + pageSize));
    if (!pages.length)
        pages.push(['Our Reflective Practice Growth Map']);
    const parts = [];
    parts[3] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>';
    const kids = [];
    pages.forEach(function (pageLines, index) {
        const contentId = 4 + index * 2;
        const pageId = contentId + 1;
        let stream = 'BT\n/F1 11 Tf\n36 750 Td\n';
        pageLines.forEach(function (line, lineIndex) {
            if (lineIndex)
                stream += '0 -15 Td\n';
            stream += '(' + pdfEscape(line) + ') Tj\n';
        });
        stream += 'ET';
        parts[contentId] = '<< /Length ' + stream.length + ' >>\nstream\n' + stream + '\nendstream';
        parts[pageId] = '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents ' + contentId + ' 0 R /Resources << /Font << /F1 3 0 R >> >> >>';
        kids.push(pageId + ' 0 R');
    });
    parts[1] = '<< /Type /Catalog /Pages 2 0 R >>';
    parts[2] = '<< /Type /Pages /Kids [' + kids.join(' ') + '] /Count ' + pages.length + ' >>';
    const header = '%PDF-1.4\n';
    let body = '';
    const offsets = [];
    for (let n = 1; n < parts.length; n++) {
        if (!parts[n])
            continue;
        offsets[n] = header.length + body.length;
        body += n + ' 0 obj\n' + parts[n] + '\nendobj\n';
    }
    let xref = 'xref\n0 ' + parts.length + '\n0000000000 65535 f \n';
    for (let p = 1; p < parts.length; p++) {
        xref += String(offsets[p] || 0).padStart(10, '0') + ' 00000 n \n';
    }
    const pdf = header + body + xref + 'trailer\n<< /Size ' + parts.length + ' /Root 1 0 R >>\nstartxref\n' + (header.length + body.length) + '\n%%EOF';
    return Buffer.from(pdf, 'binary');
}
