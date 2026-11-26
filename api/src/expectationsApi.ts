import { getNoteSummaries } from './intakeQService';
import { writeMileageDashboardCopy } from './mileageWorkbook';
import { listItems, LISTS } from './sharepointService';
import {
  OUTREACH_MESSAGE,
  VisitFrequency,
  addDays,
  cancellationClass,
  clientKey,
  denverDate,
  denverMonth,
  haversineMiles,
  hoursSince,
  incidentLimit,
  mileageRatio,
  noteIsLate,
  outsideZone,
  pastEndOfWeek,
  ratioIsHigh,
} from './expectationsRules';
import {
  BoardState,
  ExpectationsStoreError,
  FollowUp,
  TripRecord,
  TravelRequest,
  loadBoard,
  newId,
  updateBoard,
} from './expectationsStore';
import { geocodeAddress } from './expectationsZone';

type Person = { email: string; name: string };

function clip(value: unknown, max: number): string {
  return String(value == null ? '' : value).replace(/\u0000/g, '').trim().slice(0, max);
}

function emailOf(value: string): string {
  return clip(value, 200).toLowerCase();
}

function roundMiles(value: unknown): number {
  const miles = Number(value);
  if (!Number.isFinite(miles) || miles < 0 || miles > 400) return NaN;
  return Math.round(miles * 10) / 10;
}

function frequencyOf(value: unknown): VisitFrequency {
  const text = clip(value, 40);
  if (text === 'monthly' || text === 'twice-monthly') return text;
  return 'weekly';
}

function own<T extends { staffEmail: string }>(rows: T[], email: string, isAdmin: boolean): T[] {
  if (isAdmin) return rows;
  return rows.filter(function (row) { return row.staffEmail === email; });
}

function monthTrips(trips: TripRecord[], email: string, month: string): TripRecord[] {
  return trips.filter(function (trip) { return trip.staffEmail === email && trip.date.slice(0, 7) === month; });
}

function publicHome(state: BoardState, email: string) {
  const home = state.homes[email];
  if (!home) return null;
  return { address: home.address, mapLink: home.mapLink, saved: true };
}

async function notesFor(email: string, isAdmin: boolean, now: Date) {
  const notes = await getNoteSummaries();
  const late: Array<Record<string, unknown>> = [];
  const evals: Array<Record<string, unknown>> = [];
  const reports: Array<Record<string, unknown>> = [];
  notes.forEach(function (note) {
    if (!note.open) return;
    if (!isAdmin && note.practitionerEmail !== email) return;
    const row = {
      id: note.id,
      staffEmail: note.practitionerEmail,
      staffName: note.practitionerName || note.practitionerEmail,
      noteName: note.noteName,
      date: note.date,
      clientName: note.clientName,
      hoursOpen: hoursSince(note.date, now),
    };
    if (note.kind === 'note' && noteIsLate(note.date, now)) late.push(row);
    if (note.kind === 'eval') evals.push(Object.assign({}, row, { pastEndOfWeek: pastEndOfWeek(note.date, now) }));
    if (note.kind === 'report') reports.push(Object.assign({}, row, { pastEndOfWeek: pastEndOfWeek(note.date, now) }));
  });
  return { late, evals, reports };
}

async function cancellationsFor(email: string, isAdmin: boolean, state: BoardState, today: string) {
  const start = addDays(today, -183);
  const items = await listItems(LISTS.cancellations);
  const grouped = new Map<string, { client: string; staffName: string; staffEmail: string; productivity: number; incidents: number; rows: Array<Record<string, string>> }>();
  items.forEach(function (item) {
    const fields = item.fields;
    const staffEmail = emailOf(String(fields.StaffEmail || ''));
    if (!isAdmin && staffEmail !== email) return;
    const when = clip(fields.AppointmentDate, 40);
    const day = when.slice(0, 10);
    if (day && start && day < start) return;
    const status = clip(fields.Status, 80);
    const classified = cancellationClass(status);
    if (!classified.productivity && !classified.incident) return;
    const client = clip(fields.ClientName, 120);
    const key = clientKey(client) || 'unknown';
    let group = grouped.get(staffEmail + '|' + key);
    if (!group) {
      group = { client, staffName: clip(fields.StaffName, 120), staffEmail, productivity: 0, incidents: 0, rows: [] };
      grouped.set(staffEmail + '|' + key, group);
    }
    if (classified.productivity) group.productivity += 1;
    if (classified.incident) group.incidents += 1;
    group.rows.push({ date: day, status, kind: classified.kind, service: clip(fields.ServiceName, 120) });
  });
  return Array.from(grouped.entries()).map(function (entry) {
    const key = entry[0].split('|').slice(1).join('|');
    const group = entry[1];
    const frequency = state.frequency[entry[0]] || state.frequency[key] || 'weekly';
    const limit = incidentLimit(frequency);
    const outreach = state.outreach[entry[0]] || null;
    return {
      key: entry[0],
      client: group.client,
      staffName: group.staffName,
      staffEmail: group.staffEmail,
      frequency,
      limit,
      productivity: group.productivity,
      incidents: group.incidents,
      needsOutreach: group.incidents >= limit,
      outreach,
      message: group.incidents >= limit ? OUTREACH_MESSAGE : '',
      rows: group.rows,
    };
  }).sort(function (a, b) { return b.incidents - a.incidents || a.client.localeCompare(b.client); });
}

export async function expectationsView(user: Person, isAdmin: boolean, part: string): Promise<Record<string, unknown>> {
  const state = await loadBoard();
  const email = emailOf(user.email);
  const today = denverDate(new Date());
  if (part === 'email') {
    const checks = state.emailChecks[today] || {};
    return { checked: !!checks[email], day: today };
  }
  const month = denverMonth(new Date());
  const trips = own(state.trips, email, isAdmin);
  const mine = monthTrips(state.trips, email, month);
  const miles = mine.reduce(function (sum, trip) { return sum + trip.miles; }, 0);
  const hours = Number(state.hours[email + '|' + month] || 0);
  const ratio = mileageRatio(miles, hours);
  let notes: { late: Array<Record<string, unknown>>; evals: Array<Record<string, unknown>>; reports: Array<Record<string, unknown>> } = { late: [], evals: [], reports: [] };
  let notesError = '';
  try {
    notes = await notesFor(email, isAdmin, new Date());
  } catch (err) {
    notesError = err instanceof Error ? err.message.slice(0, 180) : 'IntakeQ could not be read';
  }
  let clients: Array<Record<string, unknown>> = [];
  let cancellationsError = '';
  try {
    clients = await cancellationsFor(email, isAdmin, state, today);
  } catch (err) {
    cancellationsError = err instanceof Error ? err.message.slice(0, 180) : 'Cancellations could not be read';
  }
  const followUps = state.followUps.filter(function (item) { return isAdmin || item.staffEmail === email; });
  return {
    role: isAdmin ? 'admin' : 'provider',
    me: email,
    day: today,
    month,
    zoneMiles: 5,
    measure: 'straight-line',
    home: publicHome(state, email),
    homes: isAdmin ? Object.keys(state.homes).map(function (key) {
      return { email: key, address: state.homes[key].address, mapLink: state.homes[key].mapLink, name: state.homes[key].name };
    }) : [],
    trips,
    monthMiles: miles,
    monthHours: hours,
    ratio,
    ratioHigh: ratioIsHigh(ratio),
    requests: own(state.requests, email, isAdmin),
    emailChecked: !!(state.emailChecks[today] && state.emailChecks[today][email]),
    notes,
    notesError,
    clients,
    cancellationsError,
    followUps,
    outreachMessage: OUTREACH_MESSAGE,
  };
}

function staffTarget(user: Person, isAdmin: boolean, requested: string): { email: string; name: string } {
  if (isAdmin && requested) return { email: emailOf(requested), name: '' };
  return { email: emailOf(user.email), name: user.name || user.email };
}

export async function expectationsAction(user: Person, isAdmin: boolean, body: Record<string, unknown>): Promise<Record<string, unknown>> {
  const action = clip(body.action, 40);
  const email = emailOf(user.email);
  if (action === 'expectations-home') {
    const address = clip(body.homeAddress, 180);
    const mapLink = clip(body.mapLink, 300);
    if (mapLink && !/^https:\/\//i.test(mapLink)) {
      throw new ExpectationsStoreError('The map link must start with https://', 400);
    }
    const point = await geocodeAddress(address);
    if (!point) throw new ExpectationsStoreError('That home address could not be found. Include the street, city, and ZIP.', 400);
    const target = staffTarget(user, isAdmin, clip(body.staffEmail, 200));
    await updateBoard(function (state) {
      state.homes[target.email] = {
        address,
        lat: point.lat,
        lon: point.lon,
        mapLink,
        name: target.name || (state.homes[target.email] && state.homes[target.email].name) || target.email,
      };
    });
    return { ok: true };
  }
  if (action === 'expectations-zone') {
    const clientAddress = clip(body.clientAddress, 180);
    const state = await loadBoard();
    const target = staffTarget(user, isAdmin, clip(body.staffEmail, 200));
    const home = state.homes[target.email];
    if (!home) throw new ExpectationsStoreError('Save the home address first.', 400);
    const point = await geocodeAddress(clientAddress);
    if (!point) throw new ExpectationsStoreError('That client address could not be found. Include the street, city, and ZIP.', 400);
    const miles = haversineMiles(home.lat, home.lon, point.lat, point.lon);
    return { miles, outside: outsideZone(miles), measure: 'straight-line' };
  }
  if (action === 'expectations-approval') {
    const clientAddress = clip(body.clientAddress, 180);
    const miles = roundMiles(body.miles);
    if (!clientAddress || !Number.isFinite(miles)) throw new ExpectationsStoreError('The address and miles are required.', 400);
    if (!outsideZone(miles)) return { ok: true, outside: false };
    const created = await updateBoard(function (state) {
      const existing = state.requests.find(function (item) {
        return item.staffEmail === email && item.status === 'pending' && item.clientAddress.toLowerCase() === clientAddress.toLowerCase();
      });
      if (existing) return;
      state.requests.push({
        id: newId('trip'),
        staffEmail: email,
        staffName: user.name || email,
        clientAddress,
        miles,
        status: 'pending',
        decidedBy: '',
        createdAt: new Date().toISOString(),
      });
    });
    const pending = created.requests.filter(function (item) { return item.staffEmail === email && item.status === 'pending'; });
    return { ok: true, outside: true, requests: pending };
  }
  if (action === 'expectations-decision') {
    if (!isAdmin) throw new ExpectationsStoreError('An administrator approves or denies travel.', 403);
    const id = clip(body.id, 40);
    const decision = clip(body.decision, 20);
    if (decision !== 'approved' && decision !== 'denied') throw new ExpectationsStoreError('Choose approve or deny.', 400);
    let found = false;
    await updateBoard(function (state) {
      state.requests.forEach(function (item) {
        if (item.id !== id) return;
        item.status = decision;
        item.decidedBy = email;
        found = true;
      });
    });
    if (!found) throw new ExpectationsStoreError('That request is no longer on the list.', 404);
    return { ok: true };
  }
  if (action === 'expectations-trip') {
    const target = staffTarget(user, isAdmin, clip(body.staffEmail, 200));
    const id = clip(body.id, 40);
    if (body.remove === true) {
      await updateBoard(function (state) {
        state.trips = state.trips.filter(function (trip) {
          if (trip.id !== id) return true;
          return !isAdmin && trip.staffEmail !== email;
        });
      });
      return { ok: true };
    }
    const miles = roundMiles(body.miles);
    const date = clip(body.date, 10);
    const startPlace = clip(body.startPlace, 160);
    const endPlace = clip(body.endPlace, 160);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !startPlace || !endPlace || !Number.isFinite(miles)) {
      throw new ExpectationsStoreError('Date, starting location, ending location, and miles are required.', 400);
    }
    await updateBoard(function (state) {
      const next: TripRecord = {
        id: id || newId('mile'),
        staffEmail: target.email,
        staffName: target.name || (state.homes[target.email] && state.homes[target.email].name) || target.email,
        date,
        startPlace,
        endPlace,
        miles,
      };
      const index = state.trips.findIndex(function (trip) { return trip.id === next.id; });
      if (index >= 0) {
        if (!isAdmin && state.trips[index].staffEmail !== email) return;
        next.staffEmail = state.trips[index].staffEmail;
        next.staffName = state.trips[index].staffName;
        state.trips[index] = next;
      } else state.trips.push(next);
    });
    return { ok: true };
  }
  if (action === 'expectations-hours') {
    const target = staffTarget(user, isAdmin, clip(body.staffEmail, 200));
    const hours = Number(body.hours);
    const month = clip(body.month, 7) || denverMonth(new Date());
    if (!/^\d{4}-\d{2}$/.test(month) || !Number.isFinite(hours) || hours < 0 || hours > 400) {
      throw new ExpectationsStoreError('Enter the hours worked this month.', 400);
    }
    await updateBoard(function (state) {
      state.hours[target.email + '|' + month] = Math.round(hours * 10) / 10;
    });
    return { ok: true };
  }
  if (action === 'expectations-copy') {
    const target = staffTarget(user, isAdmin, clip(body.staffEmail, 200));
    const state = await loadBoard();
    const month = denverMonth(new Date());
    const trips = monthTrips(state.trips, target.email, month);
    const hours = Number(state.hours[target.email + '|' + month] || 0);
    const home = state.homes[target.email];
    const saved = await writeMileageDashboardCopy(home ? home.name : target.name, target.email, trips, hours);
    if (!saved.ok) throw new ExpectationsStoreError(saved.error || 'The read-only copy could not be saved.', 400);
    return { ok: true, path: saved.path || '' };
  }
  if (action === 'expectations-email') {
    const today = denverDate(new Date());
    const checked = body.checked !== false;
    await updateBoard(function (state) {
      const day = state.emailChecks[today] || {};
      day[email] = checked;
      state.emailChecks[today] = day;
    });
    return { ok: true, checked, day: today };
  }
  if (action === 'expectations-frequency') {
    if (!isAdmin) throw new ExpectationsStoreError('An administrator sets how often the visits are scheduled.', 403);
    const key = clip(body.key, 160);
    if (!key) throw new ExpectationsStoreError('Choose a client first.', 400);
    const frequency = frequencyOf(body.frequency);
    await updateBoard(function (state) {
      state.frequency[key] = frequency;
    });
    return { ok: true };
  }
  if (action === 'expectations-followup') {
    if (!isAdmin) throw new ExpectationsStoreError('An administrator records the UNICORN or PAR follow-up.', 403);
    const noteId = clip(body.noteId, 80);
    const followUp = clip(body.followUp, 20);
    if (!noteId || (followUp !== 'unicorn' && followUp !== 'par' && followUp !== 'done')) {
      throw new ExpectationsStoreError('Choose UNICORN, PAR fax, or done.', 400);
    }
    const kind = clip(body.kind, 20) === 'report' ? 'report' : 'eval';
    await updateBoard(function (state) {
      const next: FollowUp = {
        noteId,
        staffEmail: emailOf(clip(body.staffEmail, 200)),
        noteName: clip(body.noteName, 120),
        date: clip(body.date, 40),
        kind,
        followUp,
      };
      const index = state.followUps.findIndex(function (item) { return item.noteId === noteId; });
      if (index >= 0) state.followUps[index] = next;
      else state.followUps.push(next);
    });
    return { ok: true };
  }
  if (action === 'expectations-outreach') {
    if (!isAdmin) throw new ExpectationsStoreError('An administrator records the email or the call.', 403);
    const key = clip(body.key, 160);
    const choice = clip(body.choice, 20);
    if (!key || (choice !== 'email' && choice !== 'call')) throw new ExpectationsStoreError('Choose email or a call.', 400);
    await updateBoard(function (state) {
      state.outreach[key] = { choice, at: new Date().toISOString(), by: email };
    });
    return { ok: true, message: choice === 'email' ? OUTREACH_MESSAGE : '' };
  }
  if (action === 'expectations-sync') {
    const view = await expectationsView(user, true, 'expectations');
    const due = Array.isArray(view.clients) ? (view.clients as Array<{ needsOutreach?: boolean; outreach?: unknown; message?: string; staffEmail?: string; client?: string }>).filter(function (item) {
      return item.needsOutreach && !item.outreach;
    }) : [];
    const late = view.notes && typeof view.notes === 'object' ? (view.notes as { late?: unknown[] }).late || [] : [];
    return {
      ok: true,
      day: view.day,
      lateNotes: Array.isArray(late) ? late.length : 0,
      outreachDue: due.length,
      messages: due.map(function (item) { return { staffEmail: item.staffEmail || '', client: item.client || '', message: item.message || OUTREACH_MESSAGE }; }),
    };
  }
  throw new ExpectationsStoreError('That mileage action is not recognized.', 400);
}

export function expectationsError(err: unknown): { status: number; message: string } {
  if (err instanceof ExpectationsStoreError) return { status: err.status, message: err.message };
  const message = err instanceof Error ? err.message : 'Request failed';
  return { status: 500, message: message.slice(0, 180) };
}

export function pendingRequest(requests: TravelRequest[], email: string, address: string): TravelRequest | undefined {
  return requests.find(function (item) {
    return item.staffEmail === email && item.status === 'pending' && item.clientAddress.toLowerCase() === address.toLowerCase();
  });
}
