import { BlobServiceClient } from '@azure/storage-blob';
import { randomBytes } from 'crypto';
import { GrowthMap, blankMap, cleanMap } from './reflectiveMap';

/**
 * Live growth maps and reflective supervision appointments.
 * Stored beside schedules, under reflective-supervision/, so other tiles never read them.
 * The words typed into a map are not written to the application log.
 */

const CONTAINER = process.env.AZURE_STORAGE_CONTAINER || 'emme-reports';

export type AppointmentStatus = 'proposed' | 'declined' | 'suggested' | 'accepted' | 'confirmed' | 'withdrawn';

export type Appointment = {
  id: string;
  providerEmail: string;
  providerName: string;
  adminEmail: string;
  adminName: string;
  start: string;
  end: string;
  suggestedStart: string;
  suggestedEnd: string;
  status: AppointmentStatus;
  providerEventId: string;
  adminEventId: string;
  createdAt: string;
  updatedAt: string;
};

export class MapConflictError extends Error {
  updatedBy: string;
  updatedAt: string;
  constructor(updatedBy: string, updatedAt: string) {
    super('Someone else saved this growth map after you opened it. Reload the page, then make your edits again.');
    this.updatedBy = updatedBy;
    this.updatedAt = updatedAt;
  }
}

export class StoreNotReadyError extends Error {
  status: number;
  constructor() {
    super('Reflective supervision storage is not configured.');
    this.status = 503;
  }
}

let writeChain: Promise<void> = Promise.resolve();

function locked<T>(work: () => Promise<T>): Promise<T> {
  const run = writeChain.then(work, work);
  writeChain = run.then(function () { return undefined; }, function () { return undefined; });
  return run;
}

function container() {
  const conn = process.env.AZURE_STORAGE_CONNECTION_STRING || '';
  if (!conn) throw new StoreNotReadyError();
  return BlobServiceClient.fromConnectionString(conn).getContainerClient(CONTAINER);
}

export function mapBlobName(email: string): string {
  const key = email.trim().toLowerCase().replace(/[^a-z0-9._-]/g, '_').slice(0, 120) || 'user';
  return 'reflective-supervision/maps/' + key + '.json';
}

function scheduleBlobName(email: string): string {
  const key = email.trim().toLowerCase().replace(/[^a-z0-9._-]/g, '_').slice(0, 120) || 'user';
  return 'schedules/by-email/' + key + '.json';
}

const APPOINTMENTS = 'reflective-supervision/appointments.json';

async function readJson<T>(name: string, fallback: T): Promise<T> {
  const blob = container().getBlockBlobClient(name);
  if (!(await blob.exists())) return fallback;
  try {
    return JSON.parse((await blob.downloadToBuffer()).toString('utf8')) as T;
  } catch (_) {
    return fallback;
  }
}

async function writeJson(name: string, value: unknown, owner: string): Promise<void> {
  const target = container();
  await target.createIfNotExists();
  const payload = JSON.stringify(value);
  await target.getBlockBlobClient(name).upload(payload, Buffer.byteLength(payload), {
    blobHTTPHeaders: { blobContentType: 'application/json' },
    metadata: owner ? { owner: owner.slice(0, 200) } : undefined,
  });
}

export async function readMap(email: string, name: string): Promise<GrowthMap> {
  const stored = await readJson<GrowthMap | null>(mapBlobName(email), null);
  if (!stored || !stored.topics) return blankMap(email, name);
  return cleanMap(email, stored.name || name, stored, stored.updatedBy || '', stored.updatedAt || '');
}

export async function saveMap(email: string, name: string, incoming: unknown, updatedBy: string, baseUpdatedAt: string): Promise<GrowthMap> {
  return locked(async function () {
    const current = await readJson<GrowthMap | null>(mapBlobName(email), null);
    const currentAt = current && current.updatedAt ? current.updatedAt : '';
    if (baseUpdatedAt !== currentAt) throw new MapConflictError((current && current.updatedBy) || '', currentAt);
    const saved = cleanMap(email, name, incoming, updatedBy, new Date().toISOString());
    await writeJson(mapBlobName(email), saved, email);
    return saved;
  });
}

export async function readScheduleWorkspace(email: string): Promise<{ boards?: Array<{ name?: string; days?: Array<{ name?: string; slots?: Array<{ time?: string; status?: string }> }> }> } | null> {
  return readJson(scheduleBlobName(email), null);
}

export async function listProviderEmails(): Promise<string[]> {
  const found: string[] = [];
  const box = container();
  if (!(await box.exists().catch(function () { return false; }))) return found;
  for await (const item of box.listBlobsFlat({ prefix: 'schedules/by-email/', includeMetadata: true })) {
    const owner = String((item.metadata && (item.metadata.owner || item.metadata.Owner)) || '').trim().toLowerCase();
    if (/^[a-z0-9._+-]+@emmeco\.org$/.test(owner) && found.indexOf(owner) === -1) found.push(owner);
  }
  for await (const item of box.listBlobsFlat({ prefix: 'reflective-supervision/maps/', includeMetadata: true })) {
    const owner = String((item.metadata && (item.metadata.owner || item.metadata.Owner)) || '').trim().toLowerCase();
    if (/^[a-z0-9._+-]+@emmeco\.org$/.test(owner) && found.indexOf(owner) === -1) found.push(owner);
  }
  found.sort();
  return found;
}

function blankAppointment(): Appointment {
  return {
    id: '',
    providerEmail: '',
    providerName: '',
    adminEmail: '',
    adminName: '',
    start: '',
    end: '',
    suggestedStart: '',
    suggestedEnd: '',
    status: 'proposed',
    providerEventId: '',
    adminEventId: '',
    createdAt: '',
    updatedAt: '',
  };
}

function cleanAppointment(value: unknown): Appointment | null {
  if (!value || typeof value !== 'object') return null;
  const row = value as Record<string, unknown>;
  const item = blankAppointment();
  item.id = String(row.id || '').replace(/[^a-z0-9]/g, '').slice(0, 40);
  item.providerEmail = String(row.providerEmail || '').trim().toLowerCase();
  item.providerName = String(row.providerName || '').slice(0, 120);
  item.adminEmail = String(row.adminEmail || '').trim().toLowerCase();
  item.adminName = String(row.adminName || '').slice(0, 120);
  item.start = String(row.start || '');
  item.end = String(row.end || '');
  item.suggestedStart = String(row.suggestedStart || '');
  item.suggestedEnd = String(row.suggestedEnd || '');
  const status = String(row.status || '');
  const allowed: AppointmentStatus[] = ['proposed', 'declined', 'suggested', 'accepted', 'confirmed', 'withdrawn'];
  item.status = allowed.indexOf(status as AppointmentStatus) === -1 ? 'proposed' : status as AppointmentStatus;
  item.providerEventId = String(row.providerEventId || '').slice(0, 200);
  item.adminEventId = String(row.adminEventId || '').slice(0, 200);
  item.createdAt = String(row.createdAt || '');
  item.updatedAt = String(row.updatedAt || '');
  if (!item.id || !item.providerEmail) return null;
  return item;
}

async function readAppointments(): Promise<Appointment[]> {
  const stored = await readJson<unknown[]>(APPOINTMENTS, []);
  if (!Array.isArray(stored)) return [];
  return stored.map(cleanAppointment).filter(function (item): item is Appointment { return !!item; });
}

async function writeAppointments(items: Appointment[]): Promise<void> {
  await writeJson(APPOINTMENTS, items.slice(-400), '');
}

export async function listAppointments(email: string): Promise<Appointment[]> {
  const all = await readAppointments();
  return all.filter(function (item) { return item.providerEmail === email; });
}

export async function listAllAppointments(): Promise<Appointment[]> {
  return readAppointments();
}

export async function addAppointment(item: Omit<Appointment, 'id' | 'createdAt' | 'updatedAt'>): Promise<Appointment> {
  return locked(async function () {
    const all = await readAppointments();
    const now = new Date().toISOString();
    const saved: Appointment = Object.assign({}, item, {
      id: randomBytes(8).toString('hex'),
      createdAt: now,
      updatedAt: now,
    });
    all.push(saved);
    await writeAppointments(all);
    return saved;
  });
}

export async function updateAppointment(id: string, change: (item: Appointment) => void): Promise<Appointment | null> {
  return locked(async function () {
    const all = await readAppointments();
    const item = all.find(function (row) { return row.id === id; });
    if (!item) return null;
    change(item);
    item.updatedAt = new Date().toISOString();
    await writeAppointments(all);
    return item;
  });
}
