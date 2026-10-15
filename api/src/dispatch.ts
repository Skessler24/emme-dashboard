import { automationAuthorized, readPrincipal, requireAdmin, requireUser, resolveIsAdmin } from './auth';
import { GraphError } from './graphClient';
import { IntakeQError, complianceByStaff, getCancellations, getInvoices, getNoteSummaries } from './intakeQService';
import { asText, FnContext, HttpRequest, json, methodOf, queryOf, readBody, restParts } from './http';
import { availabilityForWindow } from './schedulingService';
import {
  LISTS,
  createItem,
  employeeFolderName,
  listFolder,
  listItems,
  onboardingFolders,
  readLatestScheduleSummaries,
  updateItem,
  uploadEmployeeFile,
} from './sharepointService';

function denverToday(): string {
  return new Intl.DateTimeFormat('en-CA', {
    timeZone: 'America/Denver',
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

function mondayOf(iso: string): string {
  const parts = iso.split('-').map(Number);
  const dt = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  const day = dt.getUTCDay();
  const shift = day === 0 ? -6 : 1 - day;
  return addDays(iso, shift);
}

function publicItem(item: { id: string; createdAt?: string; fields: Record<string, unknown> }) {
  return Object.assign({ id: item.id, createdAt: item.createdAt || '' }, item.fields);
}

function matchesEmail(fields: Record<string, unknown>, email: string): boolean {
  const wanted = email.trim().toLowerCase();
  const candidates = [fields.StaffEmail, fields.EmployeeEmail, fields.PractitionerEmail, fields.FromEmail];
  return candidates.some(function (value) { return String(value || '').trim().toLowerCase() === wanted; });
}

function safeError(err: unknown): { status: number; message: string } {
  if (err instanceof GraphError || err instanceof IntakeQError) {
    const status = err.status >= 400 && err.status < 600 ? err.status : 502;
    return { status: status === 401 || status === 403 ? 502 : status, message: err.message };
  }
  const message = err instanceof Error ? err.message : 'Request failed';
  return { status: 500, message: message.slice(0, 300) };
}

function windowDays(req: HttpRequest): number {
  const raw = Number(queryOf(req, 'days') || '14');
  if (!Number.isFinite(raw)) return 14;
  return Math.min(62, Math.max(1, Math.round(raw)));
}

async function me(context: FnContext, req: HttpRequest): Promise<void> {
  const principal = readPrincipal(req);
  if (!principal || !principal.email) {
    json(context, 401, { isAdmin: false, error: 'Sign in required' });
    return;
  }
  const admin = await resolveIsAdmin(principal.email);
  json(context, 200, {
    isAdmin: admin.isAdmin,
    email: principal.email,
    name: principal.name,
    adminCheck: admin.check,
  });
}

async function onboarding(context: FnContext, req: HttpRequest): Promise<void> {
  const admin = await requireAdmin(context, req);
  if (!admin) return;
  const parts = restParts(req);
  const employeeId = parts[0] || '';
  if (!employeeId) {
    json(context, 400, { error: 'Employee id is required' });
    return;
  }
  const verb = methodOf(req);
  if (verb === 'GET' && parts.length === 1) {
    const items = await listItems(LISTS.onboarding);
    const checklist = items.filter(function (item) {
      return matchesEmail(item.fields, employeeId) || String(item.fields.EmployeeId || '') === employeeId;
    }).map(publicItem);
    const employeeName = queryOf(req, 'employeeName') || String((checklist[0] && (checklist[0] as any).EmployeeName) || '');
    let current: Array<{ name: string; lastModified: string }> = [];
    let archived: Array<{ name: string; lastModified: string }> = [];
    if (employeeName) {
      const folders = onboardingFolders(employeeName);
      const currentFiles = await listFolder(folders.current);
      const archivedFiles = await listFolder(folders.archived);
      current = currentFiles.filter(function (file) { return !file.isFolder; });
      archived = archivedFiles.filter(function (file) { return !file.isFolder; });
    }
    json(context, 200, { employeeId, checklist, currentDocuments: current, archivedDocuments: archived });
    return;
  }
  if (verb === 'POST' && parts[1] === 'documents') {
    const body = readBody(req);
    const employeeName = asText(body.employeeName, 80);
    const fileName = asText(body.fileName, 180);
    const contentBase64 = asText(body.contentBase64, 6_000_000);
    const folder = asText(body.folder, 20).toLowerCase() === 'archived' ? 'archived' : 'current';
    if (!employeeName || !fileName || !contentBase64) {
      json(context, 400, { error: 'employeeName, fileName, and contentBase64 are required' });
      return;
    }
    if (/[\\/]/.test(fileName) || fileName.indexOf('..') !== -1) {
      json(context, 400, { error: 'File name is not valid' });
      return;
    }
    const bytes = Buffer.from(contentBase64, 'base64');
    if (!bytes.length || bytes.length > 4 * 1024 * 1024) {
      json(context, 400, { error: 'File must be between 1 byte and 4 MB' });
      return;
    }
    const folders = onboardingFolders(employeeName);
    const saved = await uploadEmployeeFile(folder === 'archived' ? folders.archived : folders.current, fileName, new Uint8Array(bytes), 'application/octet-stream');
    json(context, 201, { ok: true, file: saved });
    return;
  }
  json(context, 404, { error: 'Not found' });
}

async function referrals(context: FnContext, req: HttpRequest): Promise<void> {
  const admin = await requireAdmin(context, req);
  if (!admin) return;
  const verb = methodOf(req);
  if (verb === 'GET') {
    const items = await listItems(LISTS.referrals);
    json(context, 200, { referrals: items.map(publicItem) });
    return;
  }
  if (verb === 'POST') {
    const body = readBody(req);
    const referrerName = asText(body.referrerName, 200);
    if (!referrerName) {
      json(context, 400, { error: 'referrerName is required' });
      return;
    }
    const created = await createItem(LISTS.referrals, {
      Title: referrerName.slice(0, 250),
      ReferrerName: referrerName,
      ReceivedOn: asText(body.receivedOn, 40) || new Date().toISOString(),
      Status: asText(body.status, 40) || 'New',
      AssignedTo: asText(body.assignedTo, 200),
      Notes: asText(body.notes, 4000),
    });
    json(context, 201, { referral: publicItem(created) });
    return;
  }
  json(context, 405, { error: 'Method not allowed' });
}

async function trainings(context: FnContext, req: HttpRequest): Promise<void> {
  const admin = await requireAdmin(context, req);
  if (!admin) return;
  if (methodOf(req) !== 'GET') {
    json(context, 405, { error: 'Method not allowed' });
    return;
  }
  const items = await listItems(LISTS.trainings);
  json(context, 200, { trainings: items.map(publicItem) });
}

async function staff(context: FnContext, req: HttpRequest): Promise<void> {
  const admin = await requireAdmin(context, req);
  if (!admin) return;
  if (methodOf(req) !== 'GET') {
    json(context, 405, { error: 'Method not allowed' });
    return;
  }
  const parts = restParts(req);
  const staffId = (parts[0] || '').trim().toLowerCase();
  const action = parts[1] || '';
  if (!staffId || !action) {
    json(context, 400, { error: 'Staff id is required' });
    return;
  }
  if (action === 'trainings') {
    const items = await listItems(LISTS.staffTraining);
    json(context, 200, {
      staffId,
      trainings: items.filter(function (item) { return matchesEmail(item.fields, staffId); }).map(publicItem),
    });
    return;
  }
  if (action === 'open-notes' || action === 'open-evals') {
    const kind = action === 'open-evals' ? 'eval' : 'note';
    const notes = await getNoteSummaries();
    json(context, 200, {
      source: 'intakeq',
      persisted: false,
      staffId,
      notes: notes.filter(function (note) {
        return note.open && note.kind === kind && note.practitionerEmail === staffId;
      }),
    });
    return;
  }
  if (action === 'cancellations') {
    const end = queryOf(req, 'endDate') || denverToday();
    const start = queryOf(req, 'startDate') || addDays(end, -30);
    const rows = await getCancellations(start, end, staffId);
    json(context, 200, { source: 'intakeq', persisted: false, staffId, cancellations: rows });
    return;
  }
  json(context, 404, { error: 'Not found' });
}

function alertIsActive(fields: Record<string, unknown>, now: number): boolean {
  const active = fields.Active === true || fields.Active === 'true' || fields.Active === 1;
  if (!active) return false;
  const start = fields.StartsOn ? Date.parse(String(fields.StartsOn)) : NaN;
  const end = fields.EndsOn ? Date.parse(String(fields.EndsOn)) : NaN;
  if (!Number.isNaN(start) && start > now) return false;
  if (!Number.isNaN(end) && end < now) return false;
  return true;
}

async function adminAlerts(context: FnContext, req: HttpRequest): Promise<void> {
  if (methodOf(req) !== 'GET') {
    json(context, 405, { error: 'Method not allowed' });
    return;
  }
  const activeOnly = queryOf(req, 'active') === 'true';
  if (!activeOnly) {
    const admin = await requireAdmin(context, req);
    if (!admin) return;
  } else {
    const user = await requireUser(context, req);
    if (!user) return;
  }
  const items = await listItems(LISTS.alerts);
  const now = Date.now();
  const alerts = items.map(publicItem).filter(function (item) {
    if (!activeOnly) return true;
    return alertIsActive(item as unknown as Record<string, unknown>, now);
  });
  json(context, 200, { alerts });
}

async function staffMessages(context: FnContext, req: HttpRequest): Promise<void> {
  const verb = methodOf(req);
  const parts = restParts(req);
  if (verb === 'GET' && parts.length === 0) {
    const admin = await requireAdmin(context, req);
    if (!admin) return;
    const items = (await listItems(LISTS.messages)).map(publicItem);
    items.sort(function (a: any, b: any) {
      return String(b.createdAt || b.Created || '').localeCompare(String(a.createdAt || a.Created || ''));
    });
    json(context, 200, { messages: items.slice(0, 50) });
    return;
  }
  if (verb === 'POST' && parts.length === 0) {
    const user = await requireUser(context, req);
    if (!user) return;
    const body = readBody(req);
    const message = asText(body.message, 4000);
    if (!message) {
      json(context, 400, { error: 'A message is required' });
      return;
    }
    const created = await createItem(LISTS.messages, {
      Title: 'Staff message ' + new Date().toISOString().slice(0, 16),
      Message: message,
      FromEmail: user.email,
      FromName: user.name || user.email,
      Status: 'Open',
      Reply: '',
      RepliedBy: '',
      RepliedAt: '',
    });
    json(context, 201, { ok: true, id: created.id });
    return;
  }
  if (verb === 'POST' && parts[1] === 'reply') {
    const admin = await requireAdmin(context, req);
    if (!admin) return;
    const body = readBody(req);
    const reply = asText(body.reply, 4000);
    if (!reply) {
      json(context, 400, { error: 'A reply is required' });
      return;
    }
    await updateItem(LISTS.messages, parts[0], {
      Reply: reply,
      RepliedBy: admin.email || admin.name,
      RepliedAt: new Date().toISOString(),
      Status: 'Replied',
    });
    json(context, 200, { ok: true });
    return;
  }
  json(context, 404, { error: 'Not found' });
}

async function mileage(context: FnContext, req: HttpRequest): Promise<void> {
  const verb = methodOf(req);
  if (verb === 'GET') {
    const admin = await requireAdmin(context, req);
    if (!admin) return;
    const items = await listItems(LISTS.mileage);
    json(context, 200, { mileage: items.map(publicItem) });
    return;
  }
  if (verb === 'POST') {
    const automated = automationAuthorized(req);
    const user = automated ? null : await requireUser(context, req);
    if (!automated && !user) return;
    const body = readBody(req);
    const miles = Number(body.miles);
    const purpose = asText(body.purpose, 500);
    const tripDate = asText(body.tripDate, 40);
    if (!tripDate || !purpose || !Number.isFinite(miles) || miles < 0) {
      json(context, 400, { error: 'tripDate, miles, and purpose are required' });
      return;
    }
    let staffEmail = asText(body.staffEmail, 200).toLowerCase();
    let staffName = asText(body.staffName, 200);
    if (!automated && user) {
      const admin = await resolveIsAdmin(user.email);
      if (!admin.isAdmin) {
        staffEmail = user.email;
        staffName = user.name || user.email;
      }
    }
    if (!staffEmail) staffEmail = user ? user.email : '';
    if (!staffName) staffName = staffEmail;
    const created = await createItem(LISTS.mileage, {
      Title: staffName + ' ' + tripDate,
      StaffEmail: staffEmail,
      StaffName: staffName,
      TripDate: tripDate,
      Miles: miles,
      Purpose: purpose,
      Status: asText(body.status, 40) || 'Submitted',
    });
    json(context, 201, { mileage: publicItem(created) });
    return;
  }
  json(context, 405, { error: 'Method not allowed' });
}

async function cancellations(context: FnContext, req: HttpRequest): Promise<void> {
  const verb = methodOf(req);
  if (verb === 'GET') {
    const admin = await requireAdmin(context, req);
    if (!admin) return;
    if (queryOf(req, 'source') === 'intakeq') {
      const end = queryOf(req, 'endDate') || denverToday();
      const start = queryOf(req, 'startDate') || addDays(end, -30);
      const rows = await getCancellations(start, end);
      json(context, 200, { source: 'intakeq', persisted: false, cancellations: rows });
      return;
    }
    const items = await listItems(LISTS.cancellations);
    json(context, 200, { source: 'sharepoint', cancellations: items.map(publicItem) });
    return;
  }
  if (verb === 'POST') {
    const admin = await requireAdmin(context, req);
    if (!admin) return;
    const body = readBody(req);
    // POWER AUTOMATE nightly flow: POST { "sync": true } to copy recent
    // IntakeQ cancellations into the Cancellations list. Matched on ExternalId.
    if (body.sync === true) {
      const end = asText(body.endDate, 40) || denverToday();
      const start = asText(body.startDate, 40) || addDays(end, -30);
      const rows = await getCancellations(start, end);
      const existing = await listItems(LISTS.cancellations);
      const known = new Set(existing.map(function (item) { return String(item.fields.ExternalId || ''); }));
      let created = 0;
      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        if (!row.id || known.has(row.id)) continue;
        await createItem(LISTS.cancellations, {
          Title: (row.practitionerName || row.practitionerEmail || 'Cancellation') + ' ' + row.when.slice(0, 10),
          StaffEmail: row.practitionerEmail,
          StaffName: row.practitionerName,
          AppointmentDate: row.when,
          Status: row.status,
          ServiceName: row.serviceName,
          ClientName: row.clientName,
          ExternalId: row.id,
          Source: 'IntakeQ',
        });
        created += 1;
      }
      json(context, 200, { ok: true, created, scanned: rows.length });
      return;
    }
    json(context, 400, { error: 'Send { "sync": true } to update the Cancellations list from IntakeQ' });
    return;
  }
  json(context, 405, { error: 'Method not allowed' });
}

async function documentation(context: FnContext, req: HttpRequest): Promise<void> {
  const verb = methodOf(req);
  if (verb === 'GET') {
    const admin = await requireAdmin(context, req);
    if (!admin) return;
    if (queryOf(req, 'source') === 'intakeq') {
      const rows = await complianceByStaff();
      json(context, 200, {
        source: 'intakeq',
        persisted: false,
        weekOf: mondayOf(denverToday()),
        rows,
      });
      return;
    }
    const items = await listItems(LISTS.compliance);
    json(context, 200, { source: 'sharepoint', rows: items.map(publicItem) });
    return;
  }
  if (verb === 'POST') {
    const admin = await requireAdmin(context, req);
    if (!admin) return;
    // POWER AUTOMATE weekly flow: POST { "sync": true } to store this week's
    // open note / eval / report counts on DocumentationCompliance.
    const body = readBody(req);
    if (body.sync !== true) {
      json(context, 400, { error: 'Send { "sync": true } to store this week\'s documentation counts' });
      return;
    }
    const weekOf = asText(body.weekOf, 40) || mondayOf(denverToday());
    const rows = await complianceByStaff();
    const existing = await listItems(LISTS.compliance);
    let created = 0;
    let updated = 0;
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const match = existing.find(function (item) {
        return String(item.fields.StaffEmail || '').toLowerCase() === row.staffEmail && String(item.fields.WeekOf || '') === weekOf;
      });
      const fields = {
        Title: row.staffName + ' ' + weekOf,
        StaffEmail: row.staffEmail,
        StaffName: row.staffName,
        WeekOf: weekOf,
        OpenNotes: row.openNotes,
        OpenEvals: row.openEvals,
        OpenReports: row.openReports,
      };
      if (match) {
        await updateItem(LISTS.compliance, match.id, fields);
        updated += 1;
      } else {
        await createItem(LISTS.compliance, fields);
        created += 1;
      }
    }
    json(context, 200, { ok: true, weekOf, created, updated });
    return;
  }
  json(context, 405, { error: 'Method not allowed' });
}

async function billing(context: FnContext, req: HttpRequest): Promise<void> {
  const admin = await requireAdmin(context, req);
  if (!admin) return;
  if (methodOf(req) !== 'GET' || restParts(req)[0] !== 'invoices') {
    json(context, 404, { error: 'Not found' });
    return;
  }
  const end = queryOf(req, 'endDate') || denverToday();
  const start = queryOf(req, 'startDate') || addDays(end, -90);
  const invoices = await getInvoices({
    status: queryOf(req, 'status'),
    startDate: start,
    endDate: end,
    practitionerEmail: queryOf(req, 'practitionerEmail'),
  });
  json(context, 200, { source: 'intakeq', persisted: false, invoices });
}

async function scheduling(context: FnContext, req: HttpRequest): Promise<void> {
  if (methodOf(req) !== 'GET') {
    json(context, 405, { error: 'Method not allowed' });
    return;
  }
  const action = restParts(req)[0] || '';
  if (action === 'availability') {
    // Anonymous callers must present the automation key. Signed-in admins may preview.
    // Power Automate uses this JSON and writes summary files to SharePoint itself.
    const admin = await requireAdmin(context, req);
    if (!admin) return;
    const result = await availabilityForWindow(windowDays(req));
    if (!result.configured) {
      json(context, 503, {
        configured: false,
        error: 'Scheduling storage is not configured. Set AZURE_STORAGE_CONNECTION_STRING.',
      });
      return;
    }
    json(context, 200, {
      source: 'emme-scheduling-db',
      configured: true,
      asOf: result.asOf,
      windowDays: result.windowDays,
      generatedAt: result.generatedAt,
      timeZone: result.timeZone,
      staff: result.staff,
    });
    return;
  }
  if (action === 'openings') {
    const admin = await requireAdmin(context, req);
    if (!admin) return;
    const files = await readLatestScheduleSummaries();
    json(context, 200, { source: 'sharepoint-schedule-summaries', staff: files });
    return;
  }
  json(context, 404, { error: 'Not found' });
}

const AREAS: Record<string, (context: FnContext, req: HttpRequest) => Promise<void>> = {
  me,
  onboarding,
  referrals,
  trainings,
  staff,
  'admin-alerts': adminAlerts,
  'staff-messages': staffMessages,
  mileage,
  cancellations,
  'documentation-compliance': documentation,
  billing,
  scheduling,
};

export async function dispatch(context: FnContext, req: HttpRequest, area: string): Promise<void> {
  if (methodOf(req) === 'OPTIONS') {
    context.res = { status: 204, headers: { 'Content-Type': 'application/json' } };
    return;
  }
  const handler = AREAS[area];
  if (!handler) {
    json(context, 404, { error: 'Not found' });
    return;
  }
  try {
    await handler(context, req);
  } catch (err) {
    const parsed = safeError(err);
    json(context, parsed.status, { error: parsed.message });
  }
}
