"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.resetReflectiveOfficeCache = resetReflectiveOfficeCache;
exports.saveReadOnlyCopy = saveReadOnlyCopy;
exports.addCalendarEvent = addCalendarEvent;
exports.notifyPerson = notifyPerson;
exports.outlookBlocks = outlookBlocks;
const graphClient_1 = require("./graphClient");
const reflectiveMap_1 = require("./reflectiveMap");
/**
 * SharePoint snapshot and Outlook for reflective supervision.
 * Calendar events and emails say that a meeting exists. They do not include what was said.
 */
const SHARE_URL = 'https://emmeco031.sharepoint.com/:f:/s/Emmeteam/IgC76zwHwa13Rp6tgPf4ysldAf8v9GKA9LPp1ML4qvUVUl8?e=4J5wWy';
let shareCache = null;
function resetReflectiveOfficeCache() {
    shareCache = null;
}
function shareId(url) {
    return 'u!' + Buffer.from(url, 'utf8').toString('base64').replace(/=+$/g, '').replace(/\//g, '_').replace(/\+/g, '-');
}
function graphMessage(err) {
    if (err instanceof graphClient_1.GraphError)
        return err.message;
    return 'Microsoft 365 did not accept that request.';
}
async function rsProvidersFolder() {
    if (shareCache && shareCache.expires > Date.now())
        return shareCache.folder;
    const item = await (0, graphClient_1.graphGet)('/shares/' + shareId(SHARE_URL) + '/driveItem?$select=id,name,parentReference');
    const driveId = String(item && item.parentReference && item.parentReference.driveId || '');
    const folderId = String(item && item.id || '');
    if (!driveId || !folderId)
        throw new graphClient_1.GraphError(502, 'The RS Providers folder could not be opened.');
    const folder = { driveId, folderId, name: String(item.name || 'RS Providers') };
    shareCache = { folder, expires: Date.now() + 5 * 60 * 1000 };
    return folder;
}
/** Replaces the provider's read-only PDF in the RS Providers folder. */
async function saveReadOnlyCopy(providerName, email, pdf) {
    try {
        const root = await rsProvidersFolder();
        const children = await (0, graphClient_1.graphGet)('/drives/' + root.driveId + '/items/' + root.folderId + '/children?$select=id,name,folder&$top=200');
        const folders = ((children && children.value) || [])
            .filter(function (item) { return item && item.folder && item.id; })
            .map(function (item) { return { id: String(item.id), name: String(item.name || '') }; });
        const matched = (0, reflectiveMap_1.matchingFolder)(folders, providerName, email);
        const parentId = matched ? matched.id : root.folderId;
        const parentName = matched ? matched.name : root.name;
        const fileName = (0, reflectiveMap_1.copyFileName)(providerName);
        const saved = await (0, graphClient_1.graphSend)('/drives/' + root.driveId + '/items/' + parentId + ':/' + encodeURIComponent(fileName) + ':/content', {
            method: 'PUT',
            rawBody: pdf,
            contentType: 'application/pdf',
        });
        return {
            ok: true,
            path: parentName + '/' + fileName,
            webUrl: String(saved && saved.webUrl || ''),
            error: '',
        };
    }
    catch (err) {
        return { ok: false, path: '', webUrl: '', error: graphMessage(err) };
    }
}
async function addCalendarEvent(email, start, end) {
    const data = await (0, graphClient_1.graphSend)('/users/' + encodeURIComponent(email) + '/events', {
        method: 'POST',
        jsonBody: {
            subject: 'Reflective supervision',
            body: {
                contentType: 'Text',
                content: 'Confirmed in the Emme dashboard. What was said in the session is not written on this calendar event.',
            },
            start: { dateTime: start + ':00', timeZone: 'America/Denver' },
            end: { dateTime: end + ':00', timeZone: 'America/Denver' },
            isReminderOn: true,
            reminderMinutesBeforeStart: 60,
        },
    });
    return String(data && data.id || '');
}
async function notifyPerson(fromEmail, toEmail, subject, text) {
    if (!fromEmail || !toEmail || fromEmail === toEmail)
        return '';
    try {
        await (0, graphClient_1.graphSend)('/users/' + encodeURIComponent(fromEmail) + '/sendMail', {
            method: 'POST',
            jsonBody: {
                message: {
                    subject,
                    body: { contentType: 'Text', content: text },
                    toRecipients: [{ emailAddress: { address: toEmail } }],
                },
                saveToSentItems: false,
            },
        });
        return '';
    }
    catch (err) {
        return graphMessage(err);
    }
}
/** The provider's Outlook events, so a reflective supervision time is not booked on top of a client visit. */
async function outlookBlocks(email, startDate, endDate) {
    try {
        const path = '/users/' + encodeURIComponent(email) + '/calendarView?startDateTime=' + encodeURIComponent(startDate + 'T00:00:00')
            + '&endDateTime=' + encodeURIComponent(endDate + 'T23:59:59')
            + '&$select=subject,start,end,showAs&$top=40';
        const data = await (0, graphClient_1.graphGet)(path, { Prefer: 'outlook.timezone="America/Denver"' });
        const events = [];
        ((data && data.value) || []).forEach(function (item) {
            const show = String(item.showAs || '').toLowerCase();
            if (show === 'free')
                return;
            events.push({
                start: String(item.start && item.start.dateTime || ''),
                end: String(item.end && item.end.dateTime || ''),
                subject: String(item.subject || 'Busy').slice(0, 120),
            });
        });
        return { events, error: '' };
    }
    catch (err) {
        return { events: [], error: graphMessage(err) };
    }
}
