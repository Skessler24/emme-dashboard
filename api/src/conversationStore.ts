import { BlobServiceClient } from '@azure/storage-blob';
import { randomBytes } from 'crypto';

/**
 * Announcements and staff messages live in the same storage account as
 * schedules. They do not depend on SharePoint lists being created first.
 */

const CONTAINER = process.env.AZURE_STORAGE_CONTAINER || 'emme-reports';
const ALERTS = 'announcements/alerts.json';
const MESSAGES = 'messages/staff.json';

export type StoredAlert = {
  id: string;
  Title: string;
  Message: string;
  Active: boolean;
  StartsOn: string;
  EndsOn: string;
  createdAt: string;
};

export type StoredMessage = {
  id: string;
  Title: string;
  Message: string;
  FromEmail: string;
  FromName: string;
  Status: string;
  Reply: string;
  RepliedBy: string;
  RepliedAt: string;
  createdAt: string;
  seenByAdmin: boolean;
  seenByStaff: boolean;
};

let writeChain: Promise<void> = Promise.resolve();

function container() {
  const conn = process.env.AZURE_STORAGE_CONNECTION_STRING || '';
  if (!conn) throw new Error('Message storage is not configured.');
  return BlobServiceClient.fromConnectionString(conn).getContainerClient(CONTAINER);
}

function locked<T>(work: () => Promise<T>): Promise<T> {
  const run = writeChain.then(work, work);
  writeChain = run.then(function () { return undefined; }, function () { return undefined; });
  return run;
}

async function readList<T>(name: string): Promise<T[]> {
  const blob = container().getBlockBlobClient(name);
  if (!(await blob.exists())) return [];
  try {
    const parsed = JSON.parse((await blob.downloadToBuffer()).toString('utf8'));
    return Array.isArray(parsed) ? parsed as T[] : [];
  } catch (_) {
    return [];
  }
}

async function writeList(name: string, rows: unknown[]): Promise<void> {
  const target = container();
  await target.createIfNotExists();
  const payload = JSON.stringify(rows);
  await target.getBlockBlobClient(name).upload(payload, Buffer.byteLength(payload), {
    blobHTTPHeaders: { blobContentType: 'application/json' },
  });
}

function id(): string {
  return randomBytes(8).toString('hex');
}

export function listAlerts(): Promise<StoredAlert[]> {
  return readList<StoredAlert>(ALERTS);
}

export function addAlert(message: string): Promise<StoredAlert> {
  return locked(async function () {
    const rows = await readList<StoredAlert>(ALERTS);
    const alert: StoredAlert = {
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

export function endAlert(alertId: string): Promise<void> {
  return locked(async function () {
    const rows = await readList<StoredAlert>(ALERTS);
    rows.forEach(function (row) {
      if (row.id === alertId) row.Active = false;
    });
    await writeList(ALERTS, rows);
  });
}

export function listMessages(): Promise<StoredMessage[]> {
  return readList<StoredMessage>(MESSAGES);
}

export function addMessage(options: { message: string; email: string; name: string }): Promise<StoredMessage> {
  return locked(async function () {
    const rows = await readList<StoredMessage>(MESSAGES);
    const created: StoredMessage = {
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

export function replyToMessage(messageId: string, reply: string, adminName: string): Promise<void> {
  return locked(async function () {
    const rows = await readList<StoredMessage>(MESSAGES);
    const row = rows.find(function (item) { return item.id === messageId; });
    if (!row) throw new Error('Message was not found');
    row.Reply = reply;
    row.RepliedBy = adminName;
    row.RepliedAt = new Date().toISOString();
    row.Status = 'Replied';
    row.seenByAdmin = true;
    row.seenByStaff = false;
    await writeList(MESSAGES, rows);
  });
}

export function unreadCount(box: 'admin' | 'staff', email: string): Promise<number> {
  return listMessages().then(function (rows) {
    if (box === 'admin') return rows.filter(function (row) { return !row.seenByAdmin; }).length;
    const mine = email.trim().toLowerCase();
    return rows.filter(function (row) {
      return row.FromEmail === mine && row.Reply && !row.seenByStaff;
    }).length;
  });
}

export function markSeen(box: 'admin' | 'staff', email: string): Promise<void> {
  return locked(async function () {
    const rows = await readList<StoredMessage>(MESSAGES);
    const mine = email.trim().toLowerCase();
    rows.forEach(function (row) {
      if (box === 'admin') row.seenByAdmin = true;
      else if (row.FromEmail === mine) row.seenByStaff = true;
    });
    await writeList(MESSAGES, rows);
  });
}
