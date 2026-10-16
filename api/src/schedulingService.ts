import { BlobServiceClient } from '@azure/storage-blob';
import { getConfig } from './config';
import { getUserDisplayName, graphIsConfigured } from './graphClient';
import { employeeFolderName } from './sharepointService';

/**
 * Scheduling source of truth: the Emme Dashboard staff scheduling blobs
 * (schedules/by-email/*.json), the same store Schedule & Caseload already uses.
 *
 * This module projects weekly boards onto calendar dates. Slot notes and the
 * client list are PHI and are never copied into the availability payload.
 *
 * The SharePoint copy is one Excel workbook per provider, written when that
 * provider saves in Schedule & Caseload:
 *   /Employees/<Name>/Current Availability/Schedules/<Name> Schedule.xlsx
 * The admin dashboard reads those workbooks. It does not embed Excel.
 */

const TIME_ZONE = 'America/Denver';
const OPEN_WEEKS: Record<string, number[]> = {
  'open 4x': [1, 2, 3, 4],
  'open 3x (1,2,3)': [1, 2, 3],
  'open 3x (2,3,4)': [2, 3, 4],
  'open 3x (1,3,4)': [1, 3, 4],
  'open 2x (1&3)': [1, 3],
  'open 2x (2&4)': [2, 4],
  'open 1x (1)': [1],
  'open 1x (2)': [2],
  'open 1x (3)': [3],
  'open 1x (4)': [4],
};

export type Opening = {
  date: string;
  day: string;
  time: string;
  status: string;
};

export type ScheduleSummary = {
  staffName: string;
  staffEmail: string;
  asOf: string;
  windowDays: number;
  generatedAt: string;
  timeZone: string;
  source: 'emme-scheduling-db';
  boardName: string;
  openings: Opening[];
  inProgress: Opening[];
  counts: { open: number; full: number; inProgress: number; notOpen: number };
};

type Slot = { time?: string; status?: string; notes?: string; lastDay?: string };
type Day = { name?: string; slots?: Slot[] };
type Board = { name?: string; days?: Day[] };
type Workspace = { ownerEmail?: string; updatedAt?: string; boards?: Board[] };

function denverToday(): string {
  return new Intl.DateTimeFormat('en-CA', {
    timeZone: TIME_ZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).format(new Date());
}

function addDays(iso: string, days: number): string {
  const parts = iso.split('-').map(Number);
  const dt = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  dt.setUTCDate(dt.getUTCDate() + days);
  return dt.toISOString().slice(0, 10);
}

function weekdayName(iso: string): string {
  const parts = iso.split('-').map(Number);
  const dt = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  return ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][dt.getUTCDay()];
}

function weekOfMonth(iso: string): number {
  return Math.ceil(Number(iso.slice(8, 10)) / 7);
}

function sameDay(slotDay: string, weekday: string): boolean {
  const a = String(slotDay || '').trim().toLowerCase();
  const b = weekday.toLowerCase();
  if (!a) return false;
  return a === b || b.indexOf(a.slice(0, 3)) === 0;
}

export function currentBoard(boards: Board[] | undefined): Board | null {
  const list = boards || [];
  const named = list.find(function (board) { return /current/i.test(String(board.name || '')); });
  if (named) return named;
  const usable = list.find(function (board) { return !/not in regular/i.test(String(board.name || '')); });
  return usable || list[0] || null;
}

function classify(status: string): 'open' | 'full' | 'progress' | 'closed' | 'ignore' {
  const key = status.trim().toLowerCase();
  if (!key) return 'ignore';
  if (OPEN_WEEKS[key]) return 'open';
  if (key === 'full') return 'full';
  if (key === 'scheduling in progress') return 'progress';
  if (key === 'lunch' || key === 'lunch/paperwork' || key === 'makeup') return 'ignore';
  if (key === 'not open') return 'closed';
  return 'ignore';
}

/**
 * Builds the schedule summary object Power Automate should save to SharePoint.
 * `workspace` is read only long enough to project statuses. Notes are ignored.
 */
export function buildScheduleSummary(options: {
  staffName: string;
  staffEmail: string;
  workspace: Workspace;
  windowDays: number;
  asOf?: string;
  generatedAt?: string;
}): ScheduleSummary {
  const asOf = options.asOf || denverToday();
  const board = currentBoard(options.workspace.boards);
  const openings: Opening[] = [];
  const inProgress: Opening[] = [];
  const counts = { open: 0, full: 0, inProgress: 0, notOpen: 0 };
  for (let offset = 0; offset < options.windowDays; offset++) {
    const date = addDays(asOf, offset);
    const weekday = weekdayName(date);
    const week = weekOfMonth(date);
    const day = ((board && board.days) || []).find(function (item) { return sameDay(String(item.name || ''), weekday); });
    ((day && day.slots) || []).forEach(function (slot) {
      const status = String(slot.status || '');
      const kind = classify(status);
      if (kind === 'ignore') return;
      const row: Opening = { date, day: weekday, time: String(slot.time || ''), status };
      if (kind === 'open') {
        const weeks = OPEN_WEEKS[status.trim().toLowerCase()] || [];
        if (weeks.indexOf(week) === -1) return;
        openings.push(row);
        counts.open += 1;
      } else if (kind === 'full') {
        counts.full += 1;
      } else if (kind === 'progress') {
        inProgress.push(row);
        counts.inProgress += 1;
      } else if (kind === 'closed') {
        counts.notOpen += 1;
      }
    });
  }
  return {
    staffName: options.staffName,
    staffEmail: options.staffEmail,
    asOf,
    windowDays: options.windowDays,
    generatedAt: options.generatedAt || new Date().toISOString(),
    timeZone: TIME_ZONE,
    source: 'emme-scheduling-db',
    boardName: String((board && board.name) || ''),
    openings,
    inProgress,
    counts,
  };
}

export function sharePointSummaryPath(staffName: string, _asOf?: string): string {
  const folder = employeeFolderName(staffName);
  return '/Employees/' + folder + '/Current Availability/Schedules/' + folder + ' Schedule.xlsx';
}

function streamToString(readable: NodeJS.ReadableStream | undefined): Promise<string> {
  return new Promise(function (resolve, reject) {
    if (!readable) { resolve(''); return; }
    const chunks: Buffer[] = [];
    readable.on('data', function (data) { chunks.push(Buffer.isBuffer(data) ? data : Buffer.from(data)); });
    readable.on('end', function () { resolve(Buffer.concat(chunks).toString('utf8')); });
    readable.on('error', reject);
  });
}

async function readBlobJson(blob: { download: () => Promise<{ readableStreamBody?: NodeJS.ReadableStream }> }): Promise<Workspace | null> {
  const download = await blob.download();
  const text = await streamToString(download.readableStreamBody);
  try {
    const parsed = JSON.parse(text);
    if (!parsed || typeof parsed !== 'object') return null;
    return parsed as Workspace;
  } catch (_) {
    return null;
  }
}

export async function loadStaffWorkspaces(): Promise<{ configured: boolean; workspaces: Array<{ email: string; updatedAt: string; workspace: Workspace }> }> {
  const cfg = getConfig();
  if (!cfg.storageConnection) return { configured: false, workspaces: [] };
  const service = BlobServiceClient.fromConnectionString(cfg.storageConnection);
  const container = service.getContainerClient(cfg.storageContainer);
  if (!(await container.exists())) return { configured: true, workspaces: [] };
  const byEmail = new Map<string, { email: string; updatedAt: string; workspace: Workspace; preferred: boolean }>();
  for await (const item of container.listBlobsFlat({ prefix: 'schedules/', includeMetadata: true })) {
    if (item.name.indexOf('/shares/') !== -1) continue;
    if ((item.properties.contentLength || 0) > 8 * 1024 * 1024) continue;
    const blob = container.getBlockBlobClient(item.name);
    const workspace = await readBlobJson(blob);
    if (!workspace) continue;
    const metaOwner = String((item.metadata && (item.metadata.owner || item.metadata.Owner)) || '').toLowerCase();
    const email = String(workspace.ownerEmail || metaOwner || '').trim().toLowerCase();
    if (!email || email.indexOf('@') === -1) continue;
    const preferred = item.name.indexOf('schedules/by-email/') === 0;
    const existing = byEmail.get(email);
    if (existing && existing.preferred && !preferred) continue;
    byEmail.set(email, {
      email,
      updatedAt: String(workspace.updatedAt || ''),
      workspace,
      preferred,
    });
  }
  return {
    configured: true,
    workspaces: Array.from(byEmail.values()).map(function (row) {
      return { email: row.email, updatedAt: row.updatedAt, workspace: row.workspace };
    }),
  };
}

export async function availabilityForWindow(windowDays: number): Promise<{
  configured: boolean;
  asOf: string;
  windowDays: number;
  generatedAt: string;
  timeZone: string;
  staff: Array<{
    staffEmail: string;
    staffName: string;
    updatedAt: string;
    boardName: string;
    sharePointPath: string;
    summary: ScheduleSummary;
  }>;
}> {
  const loaded = await loadStaffWorkspaces();
  const asOf = denverToday();
  const generatedAt = new Date().toISOString();
  if (!loaded.configured) {
    return { configured: false, asOf, windowDays, generatedAt, timeZone: TIME_ZONE, staff: [] };
  }
  const staff = [];
  for (let i = 0; i < loaded.workspaces.length; i++) {
    const row = loaded.workspaces[i];
    const staffName = graphIsConfigured() ? await getUserDisplayName(row.email) : row.email;
    const summary = buildScheduleSummary({
      staffName,
      staffEmail: row.email,
      workspace: row.workspace,
      windowDays,
      asOf,
      generatedAt,
    });
    staff.push({
      staffEmail: row.email,
      staffName,
      updatedAt: row.updatedAt,
      boardName: summary.boardName,
      sharePointPath: sharePointSummaryPath(staffName, asOf),
      summary,
    });
  }
  staff.sort(function (a, b) { return a.staffName.localeCompare(b.staffName); });
  return { configured: true, asOf, windowDays, generatedAt, timeZone: TIME_ZONE, staff };
}
