"use strict";
/**
 * Staff expectation rules taken from Emme's answers and the 2026 cancellation policy.
 * Pure functions only. No note text and no addresses are handled here.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.OUTREACH_MESSAGE = exports.TIME_ZONE = exports.NOTE_HOURS = exports.RATIO_LIMIT = exports.ZONE_MILES = void 0;
exports.denverDate = denverDate;
exports.denverMonth = denverMonth;
exports.haversineMiles = haversineMiles;
exports.outsideZone = outsideZone;
exports.mileageRatio = mileageRatio;
exports.ratioIsHigh = ratioIsHigh;
exports.hoursSince = hoursSince;
exports.noteIsLate = noteIsLate;
exports.fridayOfWeek = fridayOfWeek;
exports.pastEndOfWeek = pastEndOfWeek;
exports.addDays = addDays;
exports.incidentLimit = incidentLimit;
exports.cancellationClass = cancellationClass;
exports.clientKey = clientKey;
exports.isDashboardCopyName = isDashboardCopyName;
exports.ZONE_MILES = 5;
exports.RATIO_LIMIT = 1.8;
exports.NOTE_HOURS = 48;
exports.TIME_ZONE = 'America/Denver';
function denverDate(now) {
    return new Intl.DateTimeFormat('en-CA', {
        timeZone: exports.TIME_ZONE,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
    }).format(now);
}
function denverMonth(now) {
    return denverDate(now).slice(0, 7);
}
function haversineMiles(lat1, lon1, lat2, lon2) {
    const radius = 3958.8;
    const toRad = function (value) { return value * Math.PI / 180; };
    const dLat = toRad(lat2 - lat1);
    const dLon = toRad(lon2 - lon1);
    const a = Math.sin(dLat / 2) * Math.sin(dLat / 2)
        + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
    const miles = radius * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    return Math.round(miles * 10) / 10;
}
function outsideZone(miles) {
    return Number.isFinite(miles) && miles > exports.ZONE_MILES;
}
function mileageRatio(miles, hours) {
    if (!Number.isFinite(miles) || !Number.isFinite(hours) || hours <= 0)
        return null;
    return Math.round((miles / hours) * 100) / 100;
}
function ratioIsHigh(ratio) {
    return ratio != null && ratio > exports.RATIO_LIMIT;
}
function hoursSince(iso, now) {
    const ms = Date.parse(String(iso || ''));
    if (Number.isNaN(ms))
        return null;
    return Math.round(((now.getTime() - ms) / 36e5) * 10) / 10;
}
function noteIsLate(iso, now) {
    const hours = hoursSince(iso, now);
    return hours != null && hours > exports.NOTE_HOURS;
}
function dateParts(iso) {
    const match = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!match)
        return null;
    return { y: Number(match[1]), m: Number(match[2]), d: Number(match[3]) };
}
function isoFromUtc(y, m, d) {
    return new Date(Date.UTC(y, m - 1, d)).toISOString().slice(0, 10);
}
function fridayOfWeek(isoDate) {
    const parts = dateParts(isoDate);
    if (!parts)
        return '';
    const utc = new Date(Date.UTC(parts.y, parts.m - 1, parts.d));
    const day = utc.getUTCDay();
    const shift = day === 6 ? -1 : 5 - day;
    utc.setUTCDate(utc.getUTCDate() + shift);
    return utc.toISOString().slice(0, 10);
}
function pastEndOfWeek(iso, now) {
    const parts = dateParts(iso);
    if (!parts)
        return false;
    const friday = fridayOfWeek(isoFromUtc(parts.y, parts.m, parts.d));
    if (!friday)
        return false;
    return denverDate(now) > friday;
}
function addDays(iso, days) {
    const parts = dateParts(iso);
    if (!parts)
        return '';
    const utc = new Date(Date.UTC(parts.y, parts.m - 1, parts.d));
    utc.setUTCDate(utc.getUTCDate() + days);
    return utc.toISOString().slice(0, 10);
}
function incidentLimit(frequency) {
    if (frequency === 'monthly')
        return 1;
    if (frequency === 'twice-monthly')
        return 2;
    return 4;
}
function cancellationClass(status) {
    const text = String(status || '').toLowerCase();
    if (/reschedul/.test(text))
        return { kind: 'rescheduled', productivity: false, incident: false };
    const noShow = /no[-\s]?show|missed/.test(text);
    const late = /late/.test(text);
    const cancel = /cancel|no[-\s]?show|missed|declined/.test(text);
    if (noShow)
        return { kind: 'no-show', productivity: true, incident: true };
    if (late)
        return { kind: 'late', productivity: true, incident: true };
    if (cancel)
        return { kind: 'advance', productivity: true, incident: false };
    return { kind: 'other', productivity: false, incident: false };
}
function clientKey(name) {
    return String(name || '').toLowerCase().replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80);
}
exports.OUTREACH_MESSAGE = 'We have missed several visits. Do you still want services? Has your availability changed, or do you need a new time or a different provider? If we do not hear back, this time will be offered to another family.';
function isDashboardCopyName(name) {
    return /dashboard mileage copy/i.test(String(name || ''));
}
