import { resolveIsAdmin, requireUser } from './auth';
import { getUserDisplayName } from './graphClient';
import { FnContext, HttpRequest, json, methodOf, queryOf, readBody } from './http';
import { buildGrowthMapPdf, clientVisitTimes, monthStanding, plusHour, validWhen } from './reflectiveMap';
import { addCalendarEvent, notifyPerson, outlookBlocks, saveReadOnlyCopy } from './reflectiveOffice';
import {
  Appointment,
  MapConflictError,
  StoreNotReadyError,
  addAppointment,
  listAllAppointments,
  listAppointments,
  listProviderEmails,
  readMap,
  readScheduleWorkspace,
  saveMap,
  updateAppointment,
} from './reflectiveStore';

/**
 * Growth map, PDF, SharePoint copy, and appointment requests.
 * A provider can open only their own map. An administrator can open any provider.
 */

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

function niceName(email: string): string {
  const local = email.split('@')[0] || email;
  return local.replace(/[._-]+/g, ' ').replace(/\b[a-z]/g, function (c) { return c.toUpperCase(); });
}

async function personName(email: string, fallback: string): Promise<string> {
  try {
    const real = await getUserDisplayName(email);
    if (real && real.indexOf('@') === -1) return real;
  } catch (_) { /* the name made from the email is used */ }
  return fallback || niceName(email);
}

type Caller = { email: string; name: string; admin: boolean };

async function callerOf(context: FnContext, req: HttpRequest): Promise<Caller | null> {
  const user = await requireUser(context, req);
  if (!user) return null;
  const admin = await resolveIsAdmin(user.email);
  const email = user.email.trim().toLowerCase();
  return { email, name: user.name || email, admin: admin.isAdmin };
}

function targetEmail(caller: Caller, requested: string): string | null {
  const want = requested.trim().toLowerCase();
  if (!caller.admin) {
    if (want && want !== caller.email) return null;
    return caller.email;
  }
  if (!want) return caller.email;
  if (!/^[a-z0-9._+-]+@emmeco\.org$/.test(want)) return null;
  return want;
}

function fail(context: FnContext, err: unknown): void {
  if (err instanceof MapConflictError) {
    json(context, 409, { error: err.message, updatedAt: err.updatedAt, updatedBy: err.updatedBy });
    return;
  }
  if (err instanceof StoreNotReadyError) {
    json(context, 503, { error: err.message });
    return;
  }
  throw err;
}

async function placeOnCalendars(item: Appointment): Promise<string> {
  const problems: string[] = [];
  if (!item.adminEventId && item.adminEmail) {
    try {
      const id = await addCalendarEvent(item.adminEmail, item.start, item.end);
      if (!id) problems.push('Outlook did not add the event to the administrator calendar.');
      else item.adminEventId = id;
    } catch (err) { problems.push(err instanceof Error ? err.message : 'The admin calendar was not updated.'); }
  }
  if (!item.providerEventId && item.providerEmail) {
    try {
      const id = await addCalendarEvent(item.providerEmail, item.start, item.end);
      if (!id) problems.push('Outlook did not add the event to the provider calendar.');
      else item.providerEventId = id;
    } catch (err) { problems.push(err instanceof Error ? err.message : 'The provider calendar was not updated.'); }
  }
  item.status = item.adminEventId && item.providerEventId ? 'confirmed' : 'accepted';
  return problems.join(' ');
}

export async function handleReflectiveRequest(context: FnContext, req: HttpRequest): Promise<void> {
  const caller = await callerOf(context, req);
  if (!caller) return;
  const part = queryOf(req, 'part');
  const verb = methodOf(req);
  try {
    if (part === 'providers' && verb === 'GET') {
      if (!caller.admin) {
        json(context, 403, { error: 'Emme Admin access is required' });
        return;
      }
      const emails = await listProviderEmails();
      const today = denverToday();
      const all = await listAllAppointments();
      const providers = await Promise.all(emails.map(async function (email) {
        const mine = all.filter(function (item) { return item.providerEmail === email; });
        return { email, name: await personName(email, niceName(email)), standing: monthStanding(mine, today) };
      }));
      providers.sort(function (a, b) { return a.name.localeCompare(b.name); });
      json(context, 200, { providers });
      return;
    }

    const requested = queryOf(req, 'provider') || String(readBody(req).provider || '');
    const email = targetEmail(caller, requested);
    if (!email) {
      json(context, 403, { error: 'You can open only your own reflective supervision map.' });
      return;
    }
    const providerName = await personName(email, niceName(email));

    if (part === 'map' && verb === 'GET') {
      const map = await readMap(email, providerName);
      if (!map.name) map.name = providerName;
      json(context, 200, { role: caller.admin ? 'admin' : 'provider', me: caller.email, map });
      return;
    }

    if (part === 'map' && verb === 'PUT') {
      const body = readBody(req);
      const saved = await saveMap(email, providerName, body.map, caller.email, String(body.baseUpdatedAt || ''));
      json(context, 200, { role: caller.admin ? 'admin' : 'provider', me: caller.email, map: saved });
      return;
    }

    if (part === 'pdf' && verb === 'GET') {
      const map = await readMap(email, providerName);
      const pdf = buildGrowthMapPdf(map);
      const fileName = (map.name || 'Provider').replace(/[\\/:*?"<>|]/g, ' ') + ' - Reflective Supervision.pdf';
      context.res = {
        status: 200,
        headers: {
          'Content-Type': 'application/pdf',
          'Content-Disposition': 'attachment; filename="' + fileName.replace(/"/g, '') + '"',
          'Cache-Control': 'no-store',
        },
        body: pdf,
      };
      return;
    }

    if (part === 'copy' && verb === 'POST') {
      const map = await readMap(email, providerName);
      const copy = await saveReadOnlyCopy(map.name || providerName, email, buildGrowthMapPdf(map));
      json(context, 200, { copy });
      return;
    }

    if (part === 'sessions' && verb === 'GET') {
      const workspace = await readScheduleWorkspace(email);
      const start = denverToday();
      const outlook = await outlookBlocks(email, start, addDays(start, 14));
      json(context, 200, {
        timeZone: 'America/Denver',
        visits: clientVisitTimes(workspace, 14, start),
        outlook: outlook.events,
        outlookError: outlook.error,
      });
      return;
    }

    if (part === 'appointments' && verb === 'GET') {
      const appointments = await listAppointments(email);
      json(context, 200, {
        appointments,
        standing: monthStanding(appointments, denverToday()),
        role: caller.admin ? 'admin' : 'provider',
      });
      return;
    }

    if (part === 'appointments' && verb === 'POST') {
      await changeAppointment(context, req, caller, email, providerName);
      return;
    }

    json(context, 404, { error: 'Not found' });
  } catch (err) {
    fail(context, err);
  }
}

async function changeAppointment(context: FnContext, req: HttpRequest, caller: Caller, email: string, providerName: string): Promise<void> {
  const body = readBody(req);
  const action = String(body.action || '');
  if (action === 'propose') {
    if (!caller.admin) {
      json(context, 403, { error: 'An administrator sets the reflective supervision time.' });
      return;
    }
    const start = String(body.start || '');
    const end = String(body.end || '') || (validWhen(start) ? plusHour(start) : '');
    if (!validWhen(start) || !validWhen(end) || end <= start) {
      json(context, 400, { error: 'Enter a start and end in Mountain Time.' });
      return;
    }
    const saved = await addAppointment({
      providerEmail: email,
      providerName,
      adminEmail: caller.email,
      adminName: caller.name,
      start,
      end,
      suggestedStart: '',
      suggestedEnd: '',
      status: 'proposed',
      providerEventId: '',
      adminEventId: '',
    });
    const mailWarning = await notifyPerson(
      caller.email,
      email,
      'Reflective supervision time to confirm',
      'A reflective supervision time was proposed for ' + start.replace('T', ' ') + ' Mountain Time. Open Staff Expectations on the Emme dashboard to accept, decline, or suggest another time.',
    );
    json(context, 200, { appointment: saved, mailWarning });
    return;
  }

  const id = String(body.id || '').replace(/[^a-z0-9]/g, '');
  const current = (await listAppointments(email)).find(function (item) { return item.id === id; });
  if (!current) {
    json(context, 404, { error: 'That appointment was not found.' });
    return;
  }
  if (!caller.admin && current.providerEmail !== caller.email) {
    json(context, 403, { error: 'You can answer only your own reflective supervision time.' });
    return;
  }

  if (action === 'accept' || action === 'decline' || action === 'suggest') {
    if (caller.admin && caller.email !== current.providerEmail) {
      json(context, 403, { error: 'The provider accepts, declines, or suggests a different time.' });
      return;
    }
  }
  if ((action === 'confirm' || action === 'withdraw' || action === 'calendar') && !caller.admin) {
    json(context, 403, { error: 'An administrator confirms the time.' });
    return;
  }

  const suggestStart = String(body.start || '');
  const suggestEnd = String(body.end || '') || (validWhen(suggestStart) ? plusHour(suggestStart) : '');
  if (action === 'suggest' && (!validWhen(suggestStart) || !validWhen(suggestEnd) || suggestEnd <= suggestStart)) {
    json(context, 400, { error: 'Enter the day and time you want instead.' });
    return;
  }
  let mailWarning = '';
  let calendarWarning = '';
  let saved: Appointment | null;
  try {
    saved = await updateAppointment(id, function (item) {
    if (action === 'decline' && (item.status === 'proposed' || item.status === 'suggested')) item.status = 'declined';
    else if (action === 'suggest' && (item.status === 'proposed' || item.status === 'suggested')) {
      item.suggestedStart = suggestStart;
      item.suggestedEnd = suggestEnd;
      item.status = 'suggested';
    } else if (action === 'withdraw' && item.status !== 'confirmed') item.status = 'withdrawn';
    else if (action === 'accept' && item.status === 'proposed') item.status = 'accepted';
    else if (action === 'confirm' && item.status === 'suggested' && validWhen(item.suggestedStart) && validWhen(item.suggestedEnd)) {
      item.start = item.suggestedStart;
      item.end = item.suggestedEnd;
      item.status = 'accepted';
    } else if (action === 'calendar' && (item.status === 'accepted' || item.status === 'confirmed')) {
      /* The calendars are written again below when an event is still missing. */
    } else {
      throw new Error('That appointment cannot be changed that way.');
    }
  });
  } catch (err) {
    if (err instanceof Error && err.message === 'That appointment cannot be changed that way.') {
      json(context, 400, { error: err.message });
      return;
    }
    throw err;
  }
  if (!saved) {
    json(context, 404, { error: 'That appointment was not found.' });
    return;
  }

  if (saved.status === 'declined') {
    mailWarning = await notifyPerson(caller.email, saved.adminEmail, 'Reflective supervision time declined', 'The proposed reflective supervision time was declined. Open Reflective Supervision on the admin dashboard to choose another time.');
  } else if (saved.status === 'suggested') {
    mailWarning = await notifyPerson(caller.email, saved.adminEmail, 'Another reflective supervision time was suggested', 'A different reflective supervision time was suggested: ' + saved.suggestedStart.replace('T', ' ') + ' Mountain Time. Open Reflective Supervision on the admin dashboard to confirm it.');
  } else if ((action === 'accept' || action === 'confirm' || action === 'calendar') && (saved.status === 'accepted' || saved.status === 'confirmed')) {
    calendarWarning = await placeOnCalendars(saved);
    await updateAppointment(saved.id, function (item) {
      item.status = saved.status;
      item.adminEventId = saved.adminEventId;
      item.providerEventId = saved.providerEventId;
      item.start = saved.start;
      item.end = saved.end;
    });
    if (saved.status === 'confirmed') {
      mailWarning = await notifyPerson(saved.adminEmail, saved.providerEmail, 'Reflective supervision is on the calendar', 'The reflective supervision time is confirmed and is on both Outlook calendars: ' + saved.start.replace('T', ' ') + ' Mountain Time.');
    }
  }

  json(context, 200, { appointment: saved, mailWarning, calendarWarning });
}
