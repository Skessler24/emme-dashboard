"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.listAlerts = listAlerts;
exports.addAlert = addAlert;
exports.endAlert = endAlert;
exports.listMessages = listMessages;
exports.addMessage = addMessage;
exports.replyToMessage = replyToMessage;
exports.unreadCount = unreadCount;
exports.markSeen = markSeen;
const storage_blob_1 = require("@azure/storage-blob");
const crypto_1 = require("crypto");
/**
 * Announcements and staff messages live in the same storage account as
 * schedules. They do not depend on SharePoint lists being created first.
 */
const CONTAINER = process.env.AZURE_STORAGE_CONTAINER || 'emme-reports';
const ALERTS = 'announcements/alerts.json';
const MESSAGES = 'messages/staff.json';
let writeChain = Promise.resolve();
function container() {
    const conn = process.env.AZURE_STORAGE_CONNECTION_STRING || '';
    if (!conn)
        throw new Error('Message storage is not configured.');
    return storage_blob_1.BlobServiceClient.fromConnectionString(conn).getContainerClient(CONTAINER);
}
function locked(work) {
    const run = writeChain.then(work, work);
    writeChain = run.then(function () { return undefined; }, function () { return undefined; });
    return run;
}
async function readList(name) {
    const blob = container().getBlockBlobClient(name);
    if (!(await blob.exists()))
        return [];
    try {
        const parsed = JSON.parse((await blob.downloadToBuffer()).toString('utf8'));
        return Array.isArray(parsed) ? parsed : [];
    }
    catch (_) {
        return [];
    }
}
async function writeList(name, rows) {
    const target = container();
    await target.createIfNotExists();
    const payload = JSON.stringify(rows);
    await target.getBlockBlobClient(name).upload(payload, Buffer.byteLength(payload), {
        blobHTTPHeaders: { blobContentType: 'application/json' },
    });
}
function id() {
    return (0, crypto_1.randomBytes)(8).toString('hex');
}
function listAlerts() {
    return readList(ALERTS);
}
function addAlert(message) {
    return locked(async function () {
        const rows = await readList(ALERTS);
        const alert = {
            id: id(),
            Title: message.slice(0, 80),
            Message: message,
            Active: true,
            StartsOn: new Date().toISOString(),
            EndsOn: '',
            createdAt: new Date().toISOString(),
        };
        rows.unshift(alert);
        await writeList(ALERTS, rows.slice(0, 40));
        return alert;
    });
}
function endAlert(alertId) {
    return locked(async function () {
        const rows = await readList(ALERTS);
        rows.forEach(function (row) {
            if (row.id === alertId)
                row.Active = false;
        });
        await writeList(ALERTS, rows);
    });
}
function listMessages() {
    return readList(MESSAGES);
}
function addMessage(options) {
    return locked(async function () {
        const rows = await readList(MESSAGES);
        const created = {
            id: id(),
            Title: 'Staff message ' + new Date().toISOString().slice(0, 16),
            Message: options.message,
            FromEmail: options.email,
            FromName: options.name || options.email,
            Status: 'Open',
            Reply: '',
            RepliedBy: '',
            RepliedAt: '',
            createdAt: new Date().toISOString(),
            seenByAdmin: false,
            seenByStaff: true,
        };
        rows.unshift(created);
        await writeList(MESSAGES, rows.slice(0, 200));
        return created;
    });
}
function replyToMessage(messageId, reply, adminName) {
    return locked(async function () {
        const rows = await readList(MESSAGES);
        const row = rows.find(function (item) { return item.id === messageId; });
        if (!row)
            throw new Error('Message was not found');
        row.Reply = reply;
        row.RepliedBy = adminName;
        row.RepliedAt = new Date().toISOString();
        row.Status = 'Replied';
        row.seenByAdmin = true;
        row.seenByStaff = false;
        await writeList(MESSAGES, rows);
    });
}
function unreadCount(box, email) {
    return listMessages().then(function (rows) {
        if (box === 'admin')
            return rows.filter(function (row) { return !row.seenByAdmin; }).length;
        const mine = email.trim().toLowerCase();
        return rows.filter(function (row) {
            return row.FromEmail === mine && row.Reply && !row.seenByStaff;
        }).length;
    });
}
function markSeen(box, email) {
    return locked(async function () {
        const rows = await readList(MESSAGES);
        const mine = email.trim().toLowerCase();
        rows.forEach(function (row) {
            if (box === 'admin')
                row.seenByAdmin = true;
            else if (row.FromEmail === mine)
                row.seenByStaff = true;
        });
        await writeList(MESSAGES, rows);
    });
}
