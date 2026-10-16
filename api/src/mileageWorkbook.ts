import { BlobServiceClient } from '@azure/storage-blob';
import ExcelJS from 'exceljs';
import { GraphError, getSharePointSiteId, getUserDisplayName, graphGetBytes, graphIsConfigured } from './graphClient';
import { getAppointments, isCancellationStatus, type AppointmentRow } from './intakeQService';
import { employeeFolderName, uploadEmployeeFile } from './sharepointService';

/**
 * One living mileage workbook per provider, written from IntakeQ sessions.
 * A client records mileage only when the provider marks them In person and
 * enters an average miles-per-trip. Virtual visits do not add a row.
 * Client names stay in the SharePoint file. The admin confirms rows for billing.
 */

const RATE = 0.67;
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const SHEET = 'Mileage';

export type MileageRow = {
  id: string;
  staffName: string;
  staffEmail: string;
  date: string;
  client: string;
  service: string;
  miles: number;
  amount: number;
  intakeq: string;
  billing: string;
  appointmentId: string;
  path: string;
};

type Client = {
  name?: string;
  inPerson?: boolean | string;
  milesPerTrip?: string | number;
};

type Workspace = { ownerEmail?: string; clients?: Client[] };

const COLUMNS = ['Date', 'Client', 'Service', 'Miles', 'Amount', 'IntakeQ', 'Billing', 'Appointment Id'];

function denverToday(): string {
  return new Intl.DateTimeFormat('en-CA', {
    timeZone: 'America/Denver',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).format(new Date());
}

function windowStart(today: string): string {
  const parts = today.split('-').map(Number);
  const prev = new Date(Date.UTC(parts[0], parts[1] - 2, 1));
  return prev.toISOString().slice(0, 10);
}

function normName(value: string): string {
  return String(value || '').toLowerCase().replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
}

function namesMatch(left: string, right: string): boolean {
  const a = normName(left).split(' ').filter(Boolean);
  const b = normName(right).split(' ').filter(Boolean);
  if (!a.length || !b.length) return false;
  if (a.join(' ') === b.join(' ')) return true;
  if (a.length >= 2 && b.length >= 2 && a.every(function (token) { return b.indexOf(token) !== -1; })) return true;
  if (a.length >= 2 && b.length >= 2 && b.every(function (token) { return a.indexOf(token) !== -1; })) return true;
  return false;
}

function inPerson(client: Client): boolean {
  return client.inPerson === true || client.inPerson === 'true';
}

function milesOf(client: Client): number {
  const raw = Number(String(client.milesPerTrip == null ? '' : client.milesPerTrip).replace(/[^0-9.]/g, ''));
  if (!Number.isFinite(raw) || raw <= 0 || raw > 400) return 0;
  return Math.round(raw * 10) / 10;
}

function appointmentDate(when: string): string {
  const text = String(when || '').trim();
  const iso = text.match(/^(\d{4}-\d{2}-\d{2})/);
  if (iso) return iso[1];
  const ms = Date.parse(text);
  if (Number.isNaN(ms)) return '';
  return new Intl.DateTimeFormat('en-CA', {
    timeZone: 'America/Denver',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).format(new Date(ms));
}

function excelDate(iso: string): Date {
  const parts = iso.split('-').map(Number);
  return new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
}

function dateText(value: unknown): string {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const year = value.getUTCFullYear();
    const month = String(value.getUTCMonth() + 1).padStart(2, '0');
    const day = String(value.getUTCDate()).padStart(2, '0');
    return year + '-' + month + '-' + day;
  }
  const text = String(value || '').trim();
  const iso = text.match(/^(\d{4}-\d{2}-\d{2})/);
  return iso ? iso[1] : text.slice(0, 20);
}

function money(miles: number): number {
  return Math.round(miles * RATE * 100) / 100;
}

function fileName(folder: string): string {
  return folder + ' Mileage.xlsx';
}

function folderSegments(folder: string): string[] {
  return ['Employees', folder, 'Mileage', fileName(folder)];
}

async function readWorkbookBytes(segments: string[]): Promise<Uint8Array | null> {
  try {
    const siteId = await getSharePointSiteId();
    const path = '/sites/' + siteId + '/drive/root:/' + segments.map(encodeURIComponent).join('/') + ':/content';
    return await graphGetBytes(path);
  } catch (err) {
    if (err instanceof GraphError && (err.status === 404 || /not found|itemNotFound/i.test(err.message))) return null;
    throw err;
  }
}

export async function rowsFromWorkbook(bytes: Uint8Array, staffName: string, staffEmail: string, path: string): Promise<MileageRow[]> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(Buffer.from(bytes) as unknown as ExcelJS.Buffer);
  const sheet = workbook.getWorksheet(SHEET) || workbook.worksheets[0];
  if (!sheet) return [];
  let headerRow = 0;
  const columns: Record<string, number> = {};
  sheet.eachRow(function (row, index) {
    if (headerRow) return;
    const first = String(row.getCell(1).text || row.getCell(1).value || '').trim().toLowerCase();
    if (first !== 'date') return;
    headerRow = index;
    row.eachCell(function (cell, col) {
      columns[String(cell.text || cell.value || '').trim().toLowerCase()] = col;
    });
  });
  if (!headerRow) return [];
  const rows: MileageRow[] = [];
  sheet.eachRow(function (row, index) {
    if (index <= headerRow) return;
    const date = dateText(row.getCell(columns.date || 1).value);
    const client = String(row.getCell(columns.client || 2).text || row.getCell(columns.client || 2).value || '').trim();
    if (!date && !client) return;
    const miles = Number(row.getCell(columns.miles || 4).value || 0);
    const appointmentId = String(row.getCell(columns['appointment id'] || 8).text || row.getCell(columns['appointment id'] || 8).value || '').trim();
    const billing = String(row.getCell(columns.billing || 7).text || row.getCell(columns.billing || 7).value || '').trim() || 'Pending';
    rows.push({
      id: appointmentId || staffEmail + ':' + date + ':' + client,
      staffName,
      staffEmail,
      date,
      client,
      service: String(row.getCell(columns.service || 3).text || '').trim(),
      miles: Number.isFinite(miles) ? miles : 0,
      amount: money(Number.isFinite(miles) ? miles : 0),
      intakeq: String(row.getCell(columns.intakeq || 6).text || '').trim(),
      billing,
      appointmentId,
      path,
    });
  });
  return rows;
}

export async function buildMileageWorkbook(rows: MileageRow[]): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Emme Dashboard';
  const sheet = workbook.addWorksheet(SHEET);
  sheet.getCell(1, 1).value = 'Mileage recorded from IntakeQ in-person sessions';
  sheet.getCell(2, 1).value = 'Rate $' + RATE.toFixed(2) + ' per mile. Billing stays Pending until an admin confirms it.';
  sheet.mergeCells('A1:H1');
  sheet.mergeCells('A2:H2');
  sheet.getRow(1).font = { bold: true, size: 14 };
  const header = sheet.getRow(4);
  COLUMNS.forEach(function (name, index) {
    header.getCell(index + 1).value = name;
    header.getCell(index + 1).font = { bold: true };
  });
  rows.forEach(function (row, index) {
    const line = sheet.getRow(5 + index);
    const date = /^\d{4}-\d{2}-\d{2}$/.test(row.date) ? excelDate(row.date) : row.date;
    line.getCell(1).value = date;
    if (date instanceof Date) line.getCell(1).numFmt = 'mm/dd/yyyy';
    line.getCell(2).value = row.client;
    line.getCell(3).value = row.service;
    line.getCell(4).value = row.miles;
    line.getCell(4).numFmt = '0.0';
    line.getCell(5).value = row.amount;
    line.getCell(5).numFmt = '"$"#,##0.00';
    line.getCell(6).value = row.intakeq;
    line.getCell(7).value = row.billing || 'Pending';
    line.getCell(8).value = row.appointmentId;
  });
  sheet.columns = [
    { width: 14 }, { width: 28 }, { width: 28 }, { width: 10 },
    { width: 12 }, { width: 16 }, { width: 14 }, { width: 22 },
  ];
  sheet.autoFilter = { from: 'A4', to: 'H4' };
  sheet.views = [{ state: 'frozen', ySplit: 4 }];
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

export function mergeMileageRows(options: {
  existing: MileageRow[];
  appointments: AppointmentRow[];
  clients: Client[];
  staffName: string;
  staffEmail: string;
  path: string;
  today: string;
}): { rows: MileageRow[]; added: number } {
  const start = windowStart(options.today);
  const known = new Set(options.existing.map(function (row) { return row.appointmentId; }).filter(Boolean));
  const rows = options.existing.slice();
  let added = 0;
  options.appointments.forEach(function (appointment) {
    if (!appointment.id || known.has(appointment.id)) return;
    if (!options.staffEmail || appointment.practitionerEmail !== options.staffEmail) return;
    if (isCancellationStatus(appointment.status)) return;
    const date = appointmentDate(appointment.when);
    if (!date || date < start || date > options.today) return;
    const client = (options.clients || []).find(function (item) {
      return inPerson(item) && namesMatch(String(item.name || ''), appointment.clientName);
    });
    if (!client) return;
    const miles = milesOf(client);
    if (!miles) return;
    known.add(appointment.id);
    added += 1;
    rows.push({
      id: appointment.id,
      staffName: options.staffName,
      staffEmail: options.staffEmail,
      date,
      client: String(client.name || appointment.clientName),
      service: appointment.serviceName,
      miles,
      amount: money(miles),
      intakeq: appointment.status || 'Confirmed',
      billing: 'Pending',
      appointmentId: appointment.id,
      path: options.path,
    });
  });
  rows.sort(function (a, b) {
    return a.date.localeCompare(b.date) || a.client.localeCompare(b.client);
  });
  return { rows, added };
}

async function staffFolder(email: string): Promise<{ folder: string; staffName: string }> {
  const staffName = graphIsConfigured() && email ? await getUserDisplayName(email) : (email || 'Provider');
  return { folder: employeeFolderName(staffName), staffName };
}

async function saveRows(folder: string, rows: MileageRow[]): Promise<string> {
  const bytes = new Uint8Array(await buildMileageWorkbook(rows));
  const saved = await uploadEmployeeFile(['Employees', folder, 'Mileage'], fileName(folder), bytes, XLSX_TYPE);
  return saved.path;
}

export async function syncMileageWorkbook(workspace: Workspace, email: string, appointments?: AppointmentRow[]): Promise<{ ok: boolean; path?: string; added?: number; skipped?: boolean; error?: string }> {
  try {
    const staffEmail = String(email || workspace.ownerEmail || '').trim().toLowerCase();
    const named = await staffFolder(staffEmail);
    const segments = folderSegments(named.folder);
    const path = '/' + segments.join('/');
    const currentBytes = await readWorkbookBytes(segments);
    const existing = currentBytes ? await rowsFromWorkbook(currentBytes, named.staffName, staffEmail, path) : [];
    const today = denverToday();
    const visits = appointments || await getAppointments(windowStart(today), today);
    const merged = mergeMileageRows({
      existing,
      appointments: visits,
      clients: workspace.clients || [],
      staffName: named.staffName,
      staffEmail,
      path,
      today,
    });
    const hasInPerson = (workspace.clients || []).some(function (client) { return inPerson(client) && milesOf(client) > 0; });
    if (!merged.rows.length && !hasInPerson) return { ok: true, skipped: true, added: 0 };
    const savedPath = await saveRows(named.folder, merged.rows);
    return { ok: true, path: savedPath, added: merged.added };
  } catch (err) {
    const message = err instanceof Error ? err.message : 'Could not update the mileage workbook';
    if (/not configured/i.test(message)) return { ok: false, skipped: true, error: message.slice(0, 200) };
    return { ok: false, error: message.slice(0, 200) };
  }
}

async function listProviderWorkspaces(): Promise<Array<{ email: string; workspace: Workspace }>> {
  const conn = process.env.AZURE_STORAGE_CONNECTION_STRING || '';
  if (!conn) return [];
  const container = BlobServiceClient.fromConnectionString(conn).getContainerClient(process.env.AZURE_STORAGE_CONTAINER || 'emme-reports');
  const found: Array<{ email: string; workspace: Workspace }> = [];
  for await (const blob of container.listBlobsFlat({ prefix: 'schedules/by-email/' })) {
    if (!blob.name.endsWith('.json')) continue;
    const client = container.getBlockBlobClient(blob.name);
    const text = (await client.downloadToBuffer()).toString('utf8');
    let workspace: Workspace = {};
    try { workspace = JSON.parse(text) as Workspace; } catch (_) { workspace = {}; }
    const email = String(workspace.ownerEmail || '').trim().toLowerCase();
    if (email) found.push({ email, workspace });
  }
  return found;
}

export async function syncAllMileageLogs(): Promise<{ rows: MileageRow[]; error?: string }> {
  try {
    const providers = await listProviderWorkspaces();
    const today = denverToday();
    let appointments: AppointmentRow[] = [];
    let intakeError = '';
    try {
      appointments = await getAppointments(windowStart(today), today);
    } catch (err) {
      intakeError = err instanceof Error ? err.message : 'IntakeQ could not be read';
    }
    const rows: MileageRow[] = [];
    for (let i = 0; i < providers.length; i++) {
      const provider = providers[i];
      try {
        const result = await syncMileageWorkbook(provider.workspace, provider.email, appointments);
        if (!result.ok && result.error && !intakeError) intakeError = result.error;
        const named = await staffFolder(provider.email);
        const segments = folderSegments(named.folder);
        const bytes = await readWorkbookBytes(segments);
        if (!bytes) continue;
        const parsed = await rowsFromWorkbook(bytes, named.staffName, provider.email, result.path || '/' + segments.join('/'));
        parsed.forEach(function (row) { rows.push(row); });
      } catch (err) {
        if (!intakeError) intakeError = err instanceof Error ? err.message : 'Could not update one mileage spreadsheet';
      }
    }
    rows.sort(function (a, b) { return b.date.localeCompare(a.date) || a.staffName.localeCompare(b.staffName); });
    return intakeError ? { rows, error: intakeError.slice(0, 200) } : { rows };
  } catch (err) {
    const message = err instanceof Error ? err.message : 'Could not read mileage workbooks';
    return { rows: [], error: message.slice(0, 200) };
  }
}

export async function confirmMileageRow(staffEmail: string, appointmentId: string): Promise<{ ok: boolean; error?: string }> {
  const email = String(staffEmail || '').trim().toLowerCase();
  const id = String(appointmentId || '').trim();
  if (!email || !id) return { ok: false, error: 'Staff and appointment are required' };
  try {
    const named = await staffFolder(email);
    const segments = folderSegments(named.folder);
    const path = '/' + segments.join('/');
    const bytes = await readWorkbookBytes(segments);
    if (!bytes) return { ok: false, error: 'No mileage spreadsheet for this provider yet' };
    const rows = await rowsFromWorkbook(bytes, named.staffName, email, path);
    const row = rows.find(function (item) { return item.appointmentId === id; });
    if (!row) return { ok: false, error: 'That session is not on the mileage spreadsheet' };
    row.billing = 'Confirmed';
    await saveRows(named.folder, rows);
    return { ok: true };
  } catch (err) {
    const message = err instanceof Error ? err.message : 'Could not confirm mileage';
    return { ok: false, error: message.slice(0, 200) };
  }
}
