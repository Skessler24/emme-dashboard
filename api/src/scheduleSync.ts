import crypto from 'crypto';
import { GraphError, getUserDisplayName, graphIsConfigured } from './graphClient';
import {
  LocatedSchedule,
  SheetDays,
  alignRows,
  applyDaysToWorkspace,
  applyScheduleWorkbook,
  exportedDays,
  fingerprintDays,
  locateSchedule,
  readScheduleBytes,
  readScheduleDays,
  scheduleItemInfo,
  writeScheduleBytes,
} from './scheduleWorkbook';

/**
 * Two-way sync between a provider's dashboard schedule and their SharePoint workbook.
 *
 * `meta` is saved next to the provider's schedule and remembers what the workbook looked like
 * the last time both sides agreed (eTag plus the rows with their dashboard ids). That lets us tell
 * "someone edited the workbook" from "nothing changed", and line workbook rows up with dashboard slots.
 */

// eslint-disable-next-line @typescript-eslint/no-var-requires
const scheduleMerge = require('../schedule/merge') as { merge: (base: unknown, local: unknown, remote: unknown) => any };

type Slot = { id?: string; time?: string; status?: string; notes?: string; lastDay?: string };
type Workspace = { boards?: Array<{ name?: string; days?: Array<{ name?: string; slots?: Slot[] }> }>; clients?: any[] } & Record<string, any>;

export type SyncMeta = {
  location?: { segments: string[]; fileName: string; path: string };
  eTag?: string;
  baseDays?: SheetDays;
  syncedAt?: string;
};

export type PullResult = {
  status: 'unchanged' | 'imported' | 'no-file' | 'skipped' | 'error';
  meta: SyncMeta;
  workspace?: Workspace;
  error?: string;
};

export type PushResult = {
  ok: boolean;
  meta: SyncMeta;
  conflict?: boolean;
  skipped?: boolean;
  unchanged?: boolean;
  path?: string;
  error?: string;
};

const LOCATION_TTL_MS = 10 * 60 * 1000;
const locationCache = new Map<string, { value: LocatedSchedule; expires: number }>();

function newId(): string {
  return 'slot_' + crypto.randomBytes(5).toString('hex');
}

function errorText(err: unknown, fallback: string): string {
  const message = err instanceof Error ? err.message : fallback;
  if (/403|401|accessdenied|authorization/i.test(message)) return 'Microsoft denied access to the Provider Schedules folder.';
  return message.slice(0, 200);
}

async function findLocation(email: string, meta: SyncMeta): Promise<LocatedSchedule | null> {
  const key = email.toLowerCase();
  const cached = locationCache.get(key);
  if (cached && cached.expires > Date.now()) return cached.value;
  if (meta.location) {
    const info = await scheduleItemInfo(meta.location.segments.concat([meta.location.fileName]));
    if (info) {
      locationCache.set(key, { value: meta.location, expires: Date.now() + LOCATION_TTL_MS });
      return meta.location;
    }
  }
  const staffName = await getUserDisplayName(email);
  const found = await locateSchedule(staffName, email);
  if (found) locationCache.set(key, { value: found, expires: Date.now() + LOCATION_TTL_MS });
  return found;
}

function dropLocation(email: string): void {
  locationCache.delete(email.toLowerCase());
}

function fileSegments(location: LocatedSchedule): string[] {
  return location.segments.concat([location.fileName]);
}

/**
 * Checks the provider's workbook. If someone edited it since the last sync, their changes are merged
 * into a copy of `workspace` (dashboard edits that are not in the workbook yet are kept) and returned.
 * The caller saves the returned workspace and meta.
 */
export async function pullSchedule(workspace: Workspace, email: string, meta: SyncMeta): Promise<PullResult> {
  if (!graphIsConfigured() || !email) return { status: 'skipped', meta };
  try {
    const location = await findLocation(email, meta);
    if (!location) return { status: 'no-file', meta };
    const segments = fileSegments(location);
    const info = await scheduleItemInfo(segments);
    if (!info) {
      dropLocation(email);
      return { status: 'no-file', meta };
    }
    const nextMeta: SyncMeta = Object.assign({}, meta, { location });
    if (meta.baseDays && meta.eTag && info.eTag && info.eTag === meta.eTag) return { status: 'unchanged', meta: nextMeta };

    const bytes = await readScheduleBytes(segments);
    if (!bytes) return { status: 'no-file', meta };
    const sheetDays = await readScheduleDays(bytes);

    if (!meta.baseDays) {
      // First contact. The dashboard is the source of truth, so the workbook is overwritten on the next push.
      const mine = exportedDays(workspace);
      const base: SheetDays = {};
      Object.keys(sheetDays).forEach(function (key) {
        base[key] = alignRows(mine[key] || [], sheetDays[key], newId);
      });
      nextMeta.baseDays = base;
      nextMeta.eTag = info.eTag;
      return { status: 'unchanged', meta: nextMeta };
    }

    const known = meta.baseDays;
    const theirs: SheetDays = {};
    Object.keys(sheetDays).forEach(function (key) {
      theirs[key] = alignRows(known[key] || [], sheetDays[key], newId);
    });
    nextMeta.eTag = info.eTag;
    if (fingerprintDays(theirs) === fingerprintDays(known)) return { status: 'unchanged', meta: nextMeta };

    const base = applyDaysToWorkspace(workspace, known, newId);
    const remote = applyDaysToWorkspace(workspace, theirs, newId);
    const merged = scheduleMerge.merge(base, workspace, remote) as Workspace;
    nextMeta.baseDays = theirs;
    nextMeta.syncedAt = new Date().toISOString();
    return { status: 'imported', meta: nextMeta, workspace: merged };
  } catch (err) {
    return { status: 'error', meta, error: errorText(err, 'Could not read the SharePoint schedule') };
  }
}

/**
 * Writes the dashboard schedule into the provider's workbook. Returns conflict when the workbook changed
 * after the last pull, so the caller can pull again and retry.
 */
export async function pushSchedule(workspace: Workspace, email: string, meta: SyncMeta): Promise<PushResult> {
  if (!graphIsConfigured() || !email) return { ok: false, skipped: true, meta };
  try {
    const location = await findLocation(email, meta);
    if (!location) {
      return { ok: false, meta, error: 'No existing schedule folder for ' + email + '. Nothing new was created.' };
    }
    const segments = fileSegments(location);
    const info = await scheduleItemInfo(segments);
    if (!info) {
      dropLocation(email);
      return { ok: false, meta, error: 'The schedule file for ' + email + ' could not be found.' };
    }
    const nextMeta: SyncMeta = Object.assign({}, meta, { location });
    if (meta.eTag && info.eTag && info.eTag !== meta.eTag) return { ok: false, conflict: true, meta };

    const desired = exportedDays(workspace);
    if (meta.baseDays && meta.eTag && fingerprintDays(desired) === fingerprintDays(meta.baseDays)) {
      return { ok: true, unchanged: true, path: location.path, meta: nextMeta };
    }

    const current = await readScheduleBytes(segments);
    if (!current) return { ok: false, meta, error: 'The schedule file could not be read.' };
    const next = await applyScheduleWorkbook(current, workspace);
    let saved;
    try {
      saved = await writeScheduleBytes(segments, new Uint8Array(next), info.eTag || undefined);
    } catch (err) {
      if (err instanceof GraphError && err.status === 412) return { ok: false, conflict: true, meta };
      throw err;
    }
    nextMeta.eTag = saved.eTag || ((await scheduleItemInfo(segments)) || { eTag: '' }).eTag;
    nextMeta.baseDays = desired;
    nextMeta.syncedAt = new Date().toISOString();
    return { ok: true, path: location.path, meta: nextMeta };
  } catch (err) {
    return { ok: false, meta, error: errorText(err, 'Could not update the SharePoint schedule') };
  }
}

/**
 * For a provider who has no dashboard schedule yet: builds one from their existing SharePoint workbook,
 * so opening the dashboard for the first time does not replace a schedule that is already in use.
 */
export async function importInitialSchedule(email: string, meta: SyncMeta): Promise<{ status: 'imported' | 'empty' | 'no-file' | 'skipped' | 'error'; workspace?: Workspace; meta: SyncMeta; error?: string }> {
  if (!graphIsConfigured() || !email) return { status: 'skipped', meta };
  try {
    const location = await findLocation(email, meta);
    if (!location) return { status: 'no-file', meta };
    const segments = fileSegments(location);
    const info = await scheduleItemInfo(segments);
    if (!info) {
      dropLocation(email);
      return { status: 'no-file', meta };
    }
    const bytes = await readScheduleBytes(segments);
    if (!bytes) return { status: 'no-file', meta };
    const sheetDays = await readScheduleDays(bytes);
    const keys = Object.keys(sheetDays);
    if (!keys.some(function (key) { return sheetDays[key].length > 0; })) return { status: 'empty', meta };
    const withIds: SheetDays = {};
    keys.forEach(function (key) {
      withIds[key] = sheetDays[key].map(function (row) { return Object.assign({}, row, { id: newId() }); });
    });
    const skeleton: Workspace = {
      version: 1,
      boards: [{ id: 'board_' + crypto.randomBytes(4).toString('hex'), name: 'Current Schedule', days: [] } as any],
      clients: [],
      dxCodes: [],
    };
    const workspace = applyDaysToWorkspace(skeleton, withIds, newId);
    const nextMeta: SyncMeta = { location, eTag: info.eTag, baseDays: withIds, syncedAt: new Date().toISOString() };
    return { status: 'imported', workspace, meta: nextMeta };
  } catch (err) {
    return { status: 'error', meta, error: errorText(err, 'Could not read the SharePoint schedule') };
  }
}

/** True when the dashboard has changes the workbook does not have yet (for example after a failed save). */
export function needsPush(workspace: Workspace, meta: SyncMeta): boolean {
  if (!meta || !meta.baseDays) return false;
  return fingerprintDays(exportedDays(workspace)) !== fingerprintDays(meta.baseDays);
}
