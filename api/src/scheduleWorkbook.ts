import ExcelJS from 'exceljs';
import { getUserDisplayName, graphIsConfigured } from './graphClient';
import { currentBoard } from './schedulingService';
import { employeeFolderName, uploadEmployeeFile } from './sharepointService';

/**
 * Writes one living Excel file per provider, in the same layout as
 * "Jenny Schedule.xlsx": a Current Schedule sheet with a table for each
 * weekday. Client names stay in the Notes column of that SharePoint file.
 * The Emme schedule database remains the source of truth. This file is the
 * copy the boss opens in SharePoint.
 */

const STATUSES = [
  'Open 4x',
  'Open 3x (1,2,3)',
  'Open 3x (2,3,4)',
  'Open 3x (1,3,4)',
  'Open 2x (1&3)',
  'Open 2x (2&4)',
  'Open 1x (1)',
  'Open 1x (2)',
  'Open 1x (3)',
  'Open 1x (4)',
  'Full',
  'Scheduling in Progress',
  'Lunch/Paperwork',
  'Makeup',
  'Not Open',
];

const DAY_ORDER = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];

const DAY_HEADER: Record<string, string> = {
  monday: 'Mon Appointment Time ',
  tuesday: 'Tues Appointment Time ',
  wednesday: 'Wed Appointment Time ',
  thursday: 'Thurs Appointment Time ',
  friday: 'Fri Appointment Time ',
};

type Slot = { time?: string; status?: string; notes?: string; lastDay?: string };
type Day = { name?: string; slots?: Slot[] };
type Workspace = { boards?: Array<{ name?: string; days?: Day[] }> };

const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

function dayKey(name: string): string {
  const text = String(name || '').trim().toLowerCase();
  const found = DAY_ORDER.find(function (day) { return text.startsWith(day.slice(0, 3)); });
  return found || text || 'day';
}

function dayLabel(key: string): string {
  return key.charAt(0).toUpperCase() + key.slice(1);
}

function timeHeader(key: string): string {
  return DAY_HEADER[key] || (dayLabel(key).slice(0, 3) + ' Appointment Time ');
}

function tableName(key: string, used: Set<string>): string {
  const base = dayLabel(key).replace(/[^A-Za-z0-9]/g, '') || 'Day';
  let name = base;
  let n = 2;
  while (used.has(name)) {
    name = base + n;
    n += 1;
  }
  used.add(name);
  return name;
}

export function spreadsheetStatus(status: string): string {
  const text = String(status || '').trim();
  const key = text.toLowerCase();
  if (key === 'lunch') return 'Lunch/Paperwork';
  if (key === 'scheduling in progress') return 'Scheduling in Progress';
  return text;
}

function timeFraction(text: string): number | null {
  const match = String(text || '').trim().match(/^(\d{1,2})(?::(\d{2}))?\s*(am|pm)?$/i);
  if (!match) return null;
  let hour = Number(match[1]);
  const minute = Number(match[2] || '0');
  const suffix = String(match[3] || '').toLowerCase();
  if (suffix === 'pm' && hour < 12) hour += 12;
  if (suffix === 'am' && hour === 12) hour = 0;
  if (hour > 23 || minute > 59) return null;
  return (hour * 60 + minute) / (24 * 60);
}

function formatExcelTime(value: unknown): string {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const hour = value.getUTCHours();
    const minute = value.getUTCMinutes();
    const suffix = hour >= 12 ? 'PM' : 'AM';
    const hour12 = hour % 12 || 12;
    return hour12 + ':' + String(minute).padStart(2, '0') + ' ' + suffix;
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    const minutes = Math.round((value % 1) * 24 * 60);
    const hour = Math.floor(minutes / 60) % 24;
    const minute = minutes % 60;
    const suffix = hour >= 12 ? 'PM' : 'AM';
    const hour12 = hour % 12 || 12;
    return hour12 + ':' + String(minute).padStart(2, '0') + ' ' + suffix;
  }
  return String(value || '').trim();
}

function weekdayFromHeader(text: string): string {
  const raw = String(text || '').trim();
  const lowered = raw.toLowerCase();
  if (lowered.indexOf('appointment') !== -1) {
    const prefix = lowered.replace(/appointment time/gi, '').trim();
    const found = DAY_ORDER.find(function (day) {
      return prefix === day || prefix.startsWith(day.slice(0, 3));
    });
    return found ? dayLabel(found) : '';
  }
  return DAY_ORDER.includes(lowered) ? dayLabel(lowered) : '';
}

export async function buildScheduleWorkbook(workspace: Workspace): Promise<Buffer> {
  const board = currentBoard(workspace.boards);
  const days = ((board && board.days) || []).slice().sort(function (a, b) {
    const ia = DAY_ORDER.indexOf(dayKey(String(a.name || '')));
    const ib = DAY_ORDER.indexOf(dayKey(String(b.name || '')));
    return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);
  });
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Emme Dashboard';
  const sheet = workbook.addWorksheet('Current Schedule');
  const lists = workbook.addWorksheet('Lists', { state: 'hidden' });
  STATUSES.forEach(function (status, index) {
    lists.getCell(index + 1, 1).value = status;
  });
  sheet.getColumn(1).width = 28;
  sheet.getColumn(2).width = 26;
  sheet.getColumn(3).width = 36;
  sheet.getColumn(4).width = 22;
  sheet.getColumn(6).width = 32;

  const usedNames = new Set<string>();
  let row = 1;
  days.forEach(function (day) {
    const key = dayKey(String(day.name || ''));
    const slots = (day.slots || []).filter(function (slot) {
      return String(slot.time || '').trim() || String(slot.status || '').trim() || String(slot.notes || '').trim();
    });
    const data = (slots.length ? slots : [{ time: '', status: '', notes: '', lastDay: '' }]).map(function (slot) {
      const fraction = timeFraction(String(slot.time || ''));
      return [
        fraction == null ? String(slot.time || '') : fraction,
        spreadsheetStatus(String(slot.status || '')),
        String(slot.notes || ''),
        String(slot.lastDay || ''),
      ];
    });
    const headerRow = row;
    sheet.getCell(headerRow, 6).value = 'Notes, Upcoming Changes';
    sheet.addTable({
      name: tableName(key, usedNames),
      ref: 'A' + headerRow,
      headerRow: true,
      columns: [
        { name: timeHeader(key), filterButton: true },
        { name: 'Status ', filterButton: true },
        { name: 'Notes (i.e. client names)', filterButton: true },
        { name: 'Date of Last Visit', filterButton: true },
      ],
      rows: data,
    });
    for (let i = 0; i < data.length; i++) {
      const cell = sheet.getCell(headerRow + 1 + i, 1);
      if (typeof cell.value === 'number') cell.numFmt = 'h:mm AM/PM';
    }
    const statusStart = headerRow + 1;
    const statusEnd = headerRow + data.length;
    const validations = sheet as ExcelJS.Worksheet & {
      dataValidations: { add: (address: string, rule: ExcelJS.DataValidation) => void };
    };
    validations.dataValidations.add('B' + statusStart + ':B' + statusEnd, {
      type: 'list',
      allowBlank: true,
      formulae: ['Lists!$A$1:$A$' + STATUSES.length],
      showErrorMessage: true,
      errorTitle: 'Status',
      error: 'Choose a status from the list.',
    });
    row = statusEnd + 2;
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

export async function openingsFromWorkbook(bytes: Uint8Array, staffName: string): Promise<{
  staffName: string;
  source: 'sharepoint-schedule-workbook';
  openings: Array<{ day: string; time: string; status: string }>;
}> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(Buffer.from(bytes) as unknown as ExcelJS.Buffer);
  const sheet = workbook.getWorksheet('Current Schedule') || workbook.worksheets[0];
  const openings: Array<{ day: string; time: string; status: string }> = [];
  if (!sheet) return { staffName, source: 'sharepoint-schedule-workbook', openings };
  let day = '';
  sheet.eachRow(function (sheetRow) {
    const first = String(sheetRow.getCell(1).text || sheetRow.getCell(1).value || '');
    const headerDay = weekdayFromHeader(first);
    if (headerDay) {
      day = headerDay;
      return;
    }
    const status = String(sheetRow.getCell(2).text || sheetRow.getCell(2).value || '').trim();
    const time = formatExcelTime(sheetRow.getCell(1).value);
    if (!day || !status || !/^open\s+[1-4]x/i.test(status)) return;
    openings.push({ day, time, status });
  });
  return { staffName, source: 'sharepoint-schedule-workbook', openings };
}

export async function syncScheduleWorkbook(workspace: Workspace, email: string): Promise<{ ok: boolean; path?: string; skipped?: boolean; error?: string }> {
  try {
    const staffName = graphIsConfigured() && email ? await getUserDisplayName(email) : (email || 'Provider');
    const folder = employeeFolderName(staffName);
    const fileName = folder + ' Schedule.xlsx';
    const bytes = new Uint8Array(await buildScheduleWorkbook(workspace));
    const saved = await uploadEmployeeFile(
      ['Employees', folder, 'Current Availability', 'Schedules'],
      fileName,
      bytes,
      XLSX_TYPE
    );
    return { ok: true, path: saved.path };
  } catch (err) {
    const message = err instanceof Error ? err.message : 'Could not update the SharePoint schedule';
    if (/not configured/i.test(message)) return { ok: false, skipped: true };
    return { ok: false, error: message.slice(0, 200) };
  }
}
