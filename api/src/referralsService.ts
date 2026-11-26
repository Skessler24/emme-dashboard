import { BlobServiceClient } from '@azure/storage-blob';
import { randomBytes } from 'crypto';
import { getConfig } from './config';
import { asText, FnContext, HttpRequest, json, methodOf, queryOf, readBody, restParts } from './http';
import { openingsFromDashboard } from './schedulingService';
import { LISTS, createItem, listItems, updateItem } from './sharepointService';

/**
 * Referral checklist, message drafts, and opening counts.
 * The SharePoint list Referrals still holds the family name, dates, status,
 * assigned staff, and notes. The checklist lives beside it in blob storage
 * (referrals/records.json) because that list has no columns for steps.
 * Schedule openings are read live. They are not copied into this file.
 * Nothing here sends a message.
 */

const BLOB_NAME = 'referrals/records.json';
const TIME_ZONE = 'America/Denver';
const DAY_GOAL = 28;
const RESPONSE_DAYS = 15;
const OUTREACHES_BEFORE_BOTTOM = 2;
const OWNER = 'Administrator';
const SARA_BETH = 'Sara Beth';
const PAPERWORK_LINE = 'Please contact our referrals team at referrals@emmeco.org with any paperwork or questions.';

export const EI_FIRST_EMAIL = 'Hi Sam! SB & Jenny here from EMME. I received your information from your service coordinator with RMHS';
export const EI_LAST_MESSAGE = 'Hi hi Jenny, out of respect for you and your family today will be the last day I reach out to you. I will need to hear from you by end of today or I can let your service coordinator know there has been no response. For the record, we would love to support you and your family!';
export const SC_NO_RESPONSE_MESSAGE = 'Hey, haven\'t heard a response, here is the phone number, can you verify and check in if they want services?';
export const PRIVATE_FIRST_MESSAGE = 'Hi Jenny! Sara Beth & Jenny from EMME Therapy Services here. It looks like you reached out to us on our website for Speech and OT services. Do you have time to connect with us this week to talk about needs and next steps?';

const THERAPY_IDS = ['slp', 'ot', 'pt', 'di', 'psych', 'other'] as const;
const THERAPY_LABELS: Record<string, string> = {
  slp: 'Speech',
  ot: 'Occupational therapy',
  pt: 'Physical therapy',
  di: 'DI (Developmental Intervention)',
  psych: 'Psychology',
  other: 'Other',
};

/** The shared referrals mailbox is where families write in. It is not a provider. */
function isReferralsMailbox(email: string, name: string): boolean {
  const mail = email.trim().toLowerCase();
  const who = name.trim().toLowerCase().replace(/\s+/g, ' ');
  const local = mail.split('@')[0] || '';
  if (local === 'referrals' && mail.endsWith('@emmeco.org')) return true;
  return who === 'emmeco referrals' || who === 'emme referrals' || who === 'emmeco referral';
}

const EI_STEPS: Array<{ id: string; label: string }> = [
  { id: 'call', label: 'Call' },
  { id: 'mid', label: 'MID (Most Important Document), and put Medicaid and insurance on the PAR/HIA form' },
  { id: 'provider', label: 'Provider notification' },
  { id: 'welcome', label: 'Welcome email' },
  { id: 'sc', label: 'Service coordinator notified' },
  { id: 'days', label: 'Record the days it takes to place' },
  { id: 'remove', label: 'Remove from the referral list' },
];

export type ReferralSource = { id: string; label: string; how: string; note: string; paperwork: boolean };

export const REFERRAL_SOURCES: ReferralSource[] = [
  { id: 'ai', label: 'A&I Avenues', how: 'Early Intervention Colorado website', note: 'Log in and enter provider availability. Check often so a referral sent straight to a provider is not missed.', paperwork: true },
  { id: 'blue-peaks', label: 'Blue Peaks', how: 'Emailed to the inbox', note: 'Try to accept these when you can. They serve rural areas with few providers. Tracie Lopez is the coordinator who works with Blue Peaks.', paperwork: true },
  { id: 'cci', label: 'CCI', how: 'Emailed to the inbox', note: 'These are uncommon. Scott oversees CCI. Watch the inbox.', paperwork: true },
  { id: 'blue-sky', label: 'Colorado Blue Sky', how: 'Online call-out sheet', note: 'Check the call-out sheet daily. Enter the provider name, referrals@emmeco.org, and availability.', paperwork: true },
  { id: 'cornerstone', label: 'Cornerstone', how: 'Emailed to the inbox', note: 'Try to accept these when you can. They serve rural areas with few providers.', paperwork: true },
  { id: 'ddrc', label: 'DDRC', how: 'Referral lists by email', note: 'We rarely accept these. Only a few providers are credentialed. Set the email aside unless Kelly, Sara Beth, or Jenny are credentialed to pick up children through DDRC.', paperwork: true },
  { id: 'dp', label: 'Developmental Pathways', how: 'Emailed one at a time', note: 'These expire about 48 hours after they arrive. If we cannot accept one, reply and say we are unavailable. Everyone is credentialed with Developmental Pathways except Kellie, Kyla, Linda, and Dominique.', paperwork: true },
  { id: 'ecsdd', label: 'ECSDD', how: 'Emailed to the inbox', note: 'Watch the inbox and reply with provider availability.', paperwork: true },
  { id: 'envision', label: 'Envision', how: 'Online call-out sheet', note: 'Use the call-out sheet and list referrals@emmeco.org as the contact email.', paperwork: true },
  { id: 'foothills', label: 'Foothills Gateway', how: 'Online call-out sheet', note: 'Watch the call-out sheet and list the referrals email as the contact.', paperwork: true },
  { id: 'horizons', label: 'Horizons', how: 'Emailed to the inbox', note: 'Watch the inbox and reply with provider availability.', paperwork: true },
  { id: 'inspiration', label: 'Inspiration Field', how: 'Emailed to the inbox', note: 'Watch the inbox and reply with provider availability.', paperwork: true },
  { id: 'mvds', label: 'MVDS', how: 'Emailed to the inbox', note: 'Watch the inbox and reply with provider availability.', paperwork: true },
  { id: 'north-metro', label: 'North Metro', how: 'All-call emails', note: 'Review the all-call emails and reply with the providers who are available.', paperwork: true },
  { id: 'rmhs', label: 'RMHS', how: 'Online referral spreadsheet', note: 'Track every referral on the RMHS spreadsheet. Follow up every two days until it is placed and the paperwork is received, and keep following up after placement until the paperwork arrives.', paperwork: true },
  { id: 'website', label: 'Website or a family called', how: 'Phone, voicemail, or the website', note: 'This is a direct intake, not a broker placement. Collect the child\'s information, the pediatrician, and insurance, then start intake. Phone calls go on the call sheet for Sara Beth.', paperwork: false },
  { id: 'other', label: 'Other', how: 'Not listed', note: 'The written list does not name this source. Keep the name in the notes.', paperwork: false },
];

const PRIVATE_STEPS: Array<{ id: string; label: string }> = [
  { id: 'first', label: 'First contact' },
  { id: 'needs', label: 'Needs discussion' },
  { id: 'placement', label: 'Placement' },
  { id: 'follow', label: 'Follow up' },
  { id: 'removal', label: 'Removal' },
];

type DraftSeed = { id: string; label: string; body: string };
const EI_DRAFTS: DraftSeed[] = [
  { id: 'ei-first', label: 'First message to the Early Intervention family', body: EI_FIRST_EMAIL },
  { id: 'ei-last', label: 'Last message to a family who has not answered', body: EI_LAST_MESSAGE },
  { id: 'sc-no-response', label: 'Message to the service coordinator when the family has not answered', body: SC_NO_RESPONSE_MESSAGE },
];
const PRIVATE_DRAFTS: DraftSeed[] = [
  { id: 'private-first', label: 'First message to the private family', body: PRIVATE_FIRST_MESSAGE },
];
const ASSIGN_DRAFTS: DraftSeed[] = [
  { id: 'provider-notice', label: 'Provider notification', body: '' },
  { id: 'welcome', label: 'Welcome email', body: '' },
  { id: 'sc-assigned', label: 'Tell the service coordinator this family has a staff member', body: '' },
];

export type ReferralStep = { id: string; label: string; owner: string; doneOn: string };
export type ReferralDraft = {
  id: string;
  label: string;
  body: string;
  incomplete: boolean;
  status: 'draft' | 'approved';
  approvedBy: string;
  approvedOn: string;
};
export type ReferralTodo = {
  id: string;
  title: string;
  owner: string;
  dueOn: string;
  reminderOn: string;
  doneOn: string;
};
export type ReferralDetail = {
  id: string;
  familyName: string;
  kind: '' | 'ei' | 'private';
  status: string;
  serviceCoordinator: string;
  assignedTo: string;
  assignedEmail: string;
  notes: string;
  frequency: string;
  receivedOn: string;
  lastContactOn: string;
  serviceGridOn: string;
  placedOn: string;
  removedOn: string;
  source: string;
  outreachDates: string[];
  therapies: string[];
  steps: ReferralStep[];
  drafts: ReferralDraft[];
  todos: ReferralTodo[];
  updatedAt: string;
  updatedBy: string;
};

export type ListedReferral = { id: string; createdAt: string; fields: Record<string, unknown> };
export type ListPort = {
  list(): Promise<ListedReferral[]>;
  create(fields: Record<string, unknown>): Promise<ListedReferral>;
  update(id: string, fields: Record<string, unknown>): Promise<void>;
};

export type OpeningSlot = { date?: string; day: string; time: string; status: string };
export type OpeningStaff = { staffEmail: string; staffName: string; openings: OpeningSlot[] };
export type ProviderKind = { email: string; name: string; therapies: string[] };

type StoreFile = { version: 1; records: Record<string, ReferralDetail>; kinds: ProviderKind[] };
type Actor = { email: string; name: string };
type ActionResult = { ok: true; detail: ReferralDetail } | { ok: false; status: number; error: string };

export type ReferralDeps = {
  list?: ListPort;
  today?: string;
  loadAvailability?: () => Promise<{ configured: boolean; asOf?: string; staff: OpeningStaff[] }>;
};

const sharePointList: ListPort = {
  list: function () { return listItems(LISTS.referrals); },
  create: function (fields) { return createItem(LISTS.referrals, fields); },
  update: function (id, fields) { return updateItem(LISTS.referrals, id, fields); },
};

let writeChain: Promise<void> = Promise.resolve();

function locked<T>(work: () => Promise<T>): Promise<T> {
  const run = writeChain.then(work, work);
  writeChain = run.then(function () { return undefined; }, function () { return undefined; });
  return run;
}

function denverToday(now = new Date()): string {
  return new Intl.DateTimeFormat('en-CA', {
    timeZone: TIME_ZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).format(now);
}

export function parseIsoDate(iso: string): string {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return '';
  const parts = iso.split('-').map(Number);
  const dt = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  if (dt.getUTCFullYear() !== parts[0] || dt.getUTCMonth() !== parts[1] - 1 || dt.getUTCDate() !== parts[2]) return '';
  return iso;
}

export function addDays(iso: string, days: number): string {
  const parts = iso.split('-').map(Number);
  const dt = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  dt.setUTCDate(dt.getUTCDate() + days);
  return dt.toISOString().slice(0, 10);
}

export function mondayOf(iso: string): string {
  const parts = iso.split('-').map(Number);
  const dt = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  const day = dt.getUTCDay();
  const shift = day === 0 ? -6 : 1 - day;
  return addDays(iso, shift);
}

/** Working days after the start date, through the end date. Weekends are skipped. Holidays are not. */
export function workingDaysBetween(start: string, end: string): number | null {
  const from = parseIsoDate(start);
  const to = parseIsoDate(end);
  if (!from || !to || to < from) return null;
  let count = 0;
  let cursor = addDays(from, 1);
  while (cursor <= to) {
    const parts = cursor.split('-').map(Number);
    const day = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2])).getUTCDay();
    if (day !== 0 && day !== 6) count += 1;
    cursor = addDays(cursor, 1);
  }
  return count;
}

/** Calendar days after the start date, through the end date. */
export function calendarDaysAfter(start: string, end: string): number | null {
  const from = parseIsoDate(start);
  const to = parseIsoDate(end);
  if (!from || !to || to < from) return null;
  let count = 0;
  let cursor = addDays(from, 1);
  while (cursor <= to) {
    count += 1;
    cursor = addDays(cursor, 1);
  }
  return count;
}

export function sourceByValue(value: string): ReferralSource | undefined {
  const text = value.trim().toLowerCase();
  if (!text) return undefined;
  return REFERRAL_SOURCES.find(function (item) {
    return item.id === text || item.label.toLowerCase() === text;
  });
}

function normalizeSource(value: string): string {
  const found = sourceByValue(value);
  return found ? found.id : value.trim().slice(0, 200);
}

function actorName(actor: Actor): string {
  return asText(actor.name, 120) || asText(actor.email, 120) || OWNER;
}

function cleanTherapies(value: unknown): string[] {
  const list = Array.isArray(value) ? value : [];
  const picked: string[] = [];
  list.forEach(function (item) {
    const id = asText(item, 20);
    if (THERAPY_IDS.indexOf(id as typeof THERAPY_IDS[number]) !== -1 && picked.indexOf(id) === -1) picked.push(id);
  });
  return picked;
}

function cleanKind(value: unknown): '' | 'ei' | 'private' | null {
  const kind = asText(value, 20);
  if (!kind) return '';
  if (kind === 'ei' || kind === 'private') return kind;
  return null;
}

function templateBody(id: string): string {
  const all = EI_DRAFTS.concat(PRIVATE_DRAFTS, ASSIGN_DRAFTS);
  const found = all.find(function (item) { return item.id === id; });
  return found ? found.body : '';
}

function refreshDraft(draft: ReferralDraft): void {
  const body = draft.body.trim();
  if (!body) draft.incomplete = true;
  else if (draft.id === 'ei-first') draft.incomplete = body === EI_FIRST_EMAIL;
  else draft.incomplete = false;
  if (draft.status !== 'approved') {
    draft.status = 'draft';
    draft.approvedBy = '';
    draft.approvedOn = '';
  }
}

function makeDraft(seed: DraftSeed): ReferralDraft {
  const draft: ReferralDraft = {
    id: seed.id,
    label: seed.label,
    body: seed.body,
    incomplete: false,
    status: 'draft',
    approvedBy: '',
    approvedOn: '',
  };
  refreshDraft(draft);
  return draft;
}

function makeSteps(kind: 'ei' | 'private'): ReferralStep[] {
  const source = kind === 'ei' ? EI_STEPS : PRIVATE_STEPS;
  return source.map(function (step) {
    return { id: step.id, label: step.label, owner: OWNER, doneOn: '' };
  });
}

function blankDetail(id: string, familyName: string, today: string): ReferralDetail {
  return {
    id: id,
    familyName: familyName,
    kind: '',
    status: 'New',
    serviceCoordinator: '',
    assignedTo: '',
    assignedEmail: '',
    notes: '',
    frequency: '',
    receivedOn: today,
    lastContactOn: '',
    serviceGridOn: '',
    placedOn: '',
    removedOn: '',
    source: '',
    outreachDates: [],
    therapies: [],
    steps: [],
    drafts: [],
    todos: [],
    updatedAt: '',
    updatedBy: '',
  };
}

function sharePointFields(detail: ReferralDetail): Record<string, string> {
  return {
    Title: (detail.familyName || 'Referral').slice(0, 250),
    ReferrerName: detail.familyName.slice(0, 255),
    ReceivedOn: detail.receivedOn,
    Status: detail.status || 'Active',
    AssignedTo: detail.assignedTo.slice(0, 255),
    Notes: detail.notes.slice(0, 4000),
  };
}

function textField(fields: Record<string, unknown>, key: string): string {
  return asText(fields[key], 4000);
}

function fromSharePoint(item: ListedReferral, today: string): ReferralDetail {
  const fields = item.fields || {};
  const detail = blankDetail(item.id, textField(fields, 'ReferrerName') || textField(fields, 'Title'), today);
  detail.notes = textField(fields, 'Notes');
  detail.assignedTo = textField(fields, 'AssignedTo').slice(0, 200);
  detail.status = textField(fields, 'Status') || 'New';
  const received = parseIsoDate(textField(fields, 'ReceivedOn').slice(0, 10));
  detail.receivedOn = received || textField(fields, 'ReceivedOn').slice(0, 40) || today;
  return detail;
}

function coerceStored(raw: unknown, id: string, today: string): ReferralDetail | null {
  if (!raw || typeof raw !== 'object') return null;
  const row = raw as Partial<ReferralDetail>;
  const familyName = asText(row.familyName, 200);
  if (!familyName) return null;
  const kind = cleanKind(row.kind);
  const detail = blankDetail(id, familyName, today);
  detail.kind = kind || '';
  detail.status = asText(row.status, 40) || 'Active';
  detail.serviceCoordinator = asText(row.serviceCoordinator, 200);
  detail.assignedTo = asText(row.assignedTo, 200);
  detail.assignedEmail = asText(row.assignedEmail, 200);
  detail.notes = asText(row.notes, 4000);
  detail.frequency = asText(row.frequency, 80);
  detail.receivedOn = parseIsoDate(asText(row.receivedOn, 40)) || asText(row.receivedOn, 40) || today;
  detail.lastContactOn = parseIsoDate(asText(row.lastContactOn, 40));
  detail.serviceGridOn = parseIsoDate(asText(row.serviceGridOn, 40));
  detail.placedOn = parseIsoDate(asText(row.placedOn, 40));
  detail.removedOn = parseIsoDate(asText(row.removedOn, 40));
  detail.source = normalizeSource(asText(row.source, 200));
  detail.outreachDates = Array.isArray(row.outreachDates) ? row.outreachDates.map(function (item) {
    return parseIsoDate(asText(item, 40));
  }).filter(function (item) { return !!item; }).sort() : [];
  detail.therapies = cleanTherapies(row.therapies);
  detail.steps = Array.isArray(row.steps) ? row.steps.map(function (step) {
    return {
      id: asText(step && step.id, 40),
      label: asText(step && step.label, 200),
      owner: asText(step && step.owner, 120) || OWNER,
      doneOn: parseIsoDate(asText(step && step.doneOn, 40)),
    };
  }).filter(function (step) { return !!step.id; }) : [];
  detail.drafts = Array.isArray(row.drafts) ? row.drafts.map(function (draft) {
    const status = draft && draft.status === 'approved' ? 'approved' : 'draft';
    const next: ReferralDraft = {
      id: asText(draft && draft.id, 40),
      label: asText(draft && draft.label, 200),
      body: asText(draft && draft.body, 4000),
      incomplete: false,
      status: status,
      approvedBy: status === 'approved' ? asText(draft && draft.approvedBy, 120) : '',
      approvedOn: status === 'approved' ? parseIsoDate(asText(draft && draft.approvedOn, 40)) : '',
    };
    refreshDraft(next);
    return next;
  }).filter(function (draft) { return !!draft.id; }) : [];
  detail.todos = Array.isArray(row.todos) ? row.todos.map(function (todo) {
    return {
      id: asText(todo && todo.id, 80),
      title: asText(todo && todo.title, 300),
      owner: asText(todo && todo.owner, 120) || OWNER,
      dueOn: parseIsoDate(asText(todo && todo.dueOn, 40)),
      reminderOn: parseIsoDate(asText(todo && todo.reminderOn, 40)),
      doneOn: parseIsoDate(asText(todo && todo.doneOn, 40)),
    };
  }).filter(function (todo) { return !!todo.id && !!todo.title; }) : [];
  detail.updatedAt = asText(row.updatedAt, 40);
  detail.updatedBy = asText(row.updatedBy, 120);
  return detail;
}

function untouched(detail: ReferralDetail): boolean {
  if (detail.steps.some(function (step) { return !!step.doneOn; })) return false;
  if (detail.todos.length) return false;
  return detail.drafts.every(function (draft) {
    return draft.status !== 'approved' && draft.body === templateBody(draft.id);
  });
}

function startProcess(detail: ReferralDetail, kind: 'ei' | 'private'): void {
  detail.kind = kind;
  detail.steps = makeSteps(kind);
  detail.drafts = (kind === 'ei' ? EI_DRAFTS : PRIVATE_DRAFTS).map(makeDraft);
  if (detail.status === 'New' || !detail.status) detail.status = 'Active';
}

function addTodo(detail: ReferralDetail, todo: { id: string; title: string; owner: string; dueOn: string }): void {
  if (detail.todos.some(function (item) { return item.id === todo.id && !item.doneOn; })) return;
  const id = detail.todos.some(function (item) { return item.id === todo.id; })
    ? todo.id + '-' + String(detail.todos.length + 1)
    : todo.id;
  detail.todos.push({
    id: id,
    title: todo.title,
    owner: todo.owner || OWNER,
    dueOn: todo.dueOn,
    reminderOn: todo.dueOn,
    doneOn: '',
  });
}

function dropOpenSendTodo(detail: ReferralDetail, draftId: string): void {
  detail.todos = detail.todos.filter(function (todo) {
    if (todo.doneOn) return true;
    return todo.id !== 'send-' + draftId && todo.id.indexOf('send-' + draftId + '-') !== 0;
  });
}

function ensureAssignWork(detail: ReferralDetail, actor: Actor, today: string): void {
  ASSIGN_DRAFTS.forEach(function (seed) {
    if (!detail.drafts.some(function (draft) { return draft.id === seed.id; })) detail.drafts.push(makeDraft(seed));
  });
  const owner = actorName(actor);
  addTodo(detail, {
    id: 'assign-provider',
    title: 'Provider notification: write the message, approve it, and send it yourself.',
    owner: owner,
    dueOn: today,
  });
  addTodo(detail, {
    id: 'assign-welcome',
    title: 'Welcome email: write the message, approve it, and send it yourself.',
    owner: owner,
    dueOn: today,
  });
  addTodo(detail, {
    id: 'assign-sc',
    title: 'Tell the service coordinator this family has a staff member. This does not send by itself.',
    owner: owner,
    dueOn: today,
  });
}

function trackDaysTodo(detail: ReferralDetail, today: string): void {
  if (detail.kind !== 'ei' || !detail.serviceGridOn || !detail.placedOn) return;
  addTodo(detail, {
    id: 'tracking-doc',
    title: 'Add the days to place on the referrals spreadsheet. This screen does not open that spreadsheet.',
    owner: OWNER,
    dueOn: today,
  });
}

function applyStatusFromDates(detail: ReferralDetail): void {
  if (detail.removedOn || detail.status === 'Removed') return;
  if (detail.placedOn) detail.status = 'Placed';
  else if (detail.status === 'Placed') detail.status = 'Active';
}

function familyConfirmed(detail: ReferralDetail): boolean {
  return detail.status === 'Answered' || detail.status === 'Placed' || detail.status === 'Removed' || !!detail.placedOn || !!detail.removedOn;
}

export function shouldSitAtBottom(detail: ReferralDetail, today: string): boolean {
  if (detail.kind !== 'ei' || familyConfirmed(detail)) return false;
  if (detail.outreachDates.length < OUTREACHES_BEFORE_BOTTOM) return false;
  const first = detail.outreachDates.slice().sort()[0];
  if (!first) return false;
  return today > addDays(first, RESPONSE_DAYS);
}

function applyListPosition(detail: ReferralDetail, today: string): void {
  if (familyConfirmed(detail)) return;
  if (shouldSitAtBottom(detail, today)) detail.status = 'Bottom';
  else if (detail.status === 'Bottom') detail.status = 'Active';
}

function ensureOnce(detail: ReferralDetail, todo: { id: string; title: string; owner: string; dueOn: string }): void {
  const exists = detail.todos.some(function (item) {
    return item.id === todo.id || item.id.indexOf(todo.id + '-') === 0;
  });
  if (!exists) addTodo(detail, todo);
}

function openTodo(detail: ReferralDetail, prefix: string): ReferralTodo | undefined {
  return detail.todos.find(function (item) {
    return !item.doneOn && (item.id === prefix || item.id.indexOf(prefix + '-') === 0);
  });
}

function refreshFollowUps(detail: ReferralDetail, today: string): void {
  const source = sourceByValue(detail.source);
  if (source && source.id === 'rmhs' && detail.status !== 'Removed' && !detail.removedOn) {
    const title = 'RMHS: follow up on this referral. Do this every two days until it is placed and the paperwork is received, then keep following up until the paperwork arrives. This screen does not send the follow-up.';
    const open = openTodo(detail, 'rmhs-followup');
    const doneBefore = detail.todos.some(function (item) {
      return !!item.doneOn && (item.id === 'rmhs-followup' || item.id.indexOf('rmhs-followup-') === 0);
    });
    const anchor = detail.lastContactOn || detail.receivedOn || today;
    const due = doneBefore && !open ? addDays(today, 2) : addDays(anchor, 2);
    if (open) {
      if (!doneBefore) {
        open.dueOn = addDays(anchor, 2);
        open.reminderOn = open.dueOn;
      }
      open.title = title;
    } else {
      addTodo(detail, { id: 'rmhs-followup', title: title, owner: OWNER, dueOn: due });
    }
  }
  if (source && source.id === 'dp' && !familyConfirmed(detail)) {
    const due = detail.receivedOn ? addDays(detail.receivedOn, 2) : today;
    const title = 'Developmental Pathways referrals expire about 48 hours after they arrive. If we cannot accept this one, reply and say we are unavailable. This screen does not send that reply.';
    const open = openTodo(detail, 'dp-expire');
    if (open) {
      open.dueOn = due;
      open.reminderOn = due;
      open.title = title;
    } else {
      ensureOnce(detail, { id: 'dp-expire', title: title, owner: OWNER, dueOn: due });
    }
  }
  if (detail.kind === 'ei' && !familyConfirmed(detail) && detail.serviceGridOn) {
    const elapsed = workingDaysBetween(detail.serviceGridOn, today);
    if (elapsed != null && elapsed > DAY_GOAL) {
      ensureOnce(detail, {
        id: 'day-goal-up',
        title: 'The 28 working days from the service grid are up and the family has not confirmed. They can come off the list when you remove them.',
        owner: OWNER,
        dueOn: today,
      });
    }
  }
}

function settle(detail: ReferralDetail, today: string): void {
  refreshFollowUps(detail, today);
  applyListPosition(detail, today);
}

export function buildNewReferral(body: Record<string, unknown>, today: string): ActionResult {
  const familyName = asText(body.familyName, 200) || asText(body.referrerName, 200);
  if (!familyName) return { ok: false, status: 400, error: 'Family name is required.' };
  const kind = cleanKind(body.kind);
  if (kind === null) return { ok: false, status: 400, error: 'Choose Early Intervention or private.' };
  const detail = blankDetail('', familyName, today);
  detail.serviceCoordinator = asText(body.serviceCoordinator, 200);
  detail.assignedTo = asText(body.assignedTo, 200);
  detail.notes = asText(body.notes, 4000);
  detail.frequency = asText(body.frequency, 80);
  detail.source = normalizeSource(asText(body.source, 200));
  detail.therapies = cleanTherapies(body.therapies);
  detail.receivedOn = parseIsoDate(asText(body.receivedOn, 40)) || today;
  detail.serviceGridOn = parseIsoDate(asText(body.serviceGridOn, 40));
  detail.lastContactOn = parseIsoDate(asText(body.lastContactOn, 40));
  if (kind) startProcess(detail, kind);
  else detail.status = asText(body.status, 40) || 'New';
  trackDaysTodo(detail, today);
  settle(detail, today);
  return { ok: true, detail: detail };
}

function fail(status: number, error: string): ActionResult {
  return { ok: false, status: status, error: error };
}

export function applyReferralAction(detailIn: ReferralDetail, body: Record<string, unknown>, actor: Actor, today: string): ActionResult {
  const detail: ReferralDetail = JSON.parse(JSON.stringify(detailIn));
  const action = asText(body.action, 40) || 'save';

  if (action === 'save') {
    const familyName = asText(body.familyName, 200);
    if (!familyName) return fail(400, 'Family name is required.');
    const kind = cleanKind(body.kind);
    if (kind === null) return fail(400, 'Choose Early Intervention or private.');
    if (kind !== detail.kind) {
      if (kind && !detail.kind && detail.steps.length === 0) startProcess(detail, kind);
      else if (kind && untouched(detail)) startProcess(detail, kind);
      else if (kind !== detail.kind) return fail(400, 'Uncheck the steps and leave the messages as they started before switching Early Intervention and private.');
    }
    detail.familyName = familyName;
    detail.serviceCoordinator = asText(body.serviceCoordinator, 200);
    detail.notes = asText(body.notes, 4000);
    detail.frequency = asText(body.frequency, 80);
    detail.source = normalizeSource(asText(body.source, 200));
    detail.therapies = cleanTherapies(body.therapies);
    const received = parseIsoDate(asText(body.receivedOn, 40));
    if (asText(body.receivedOn, 40) && !received) return fail(400, 'Received date must be a real date.');
    if (received) detail.receivedOn = received;
    const last = asText(body.lastContactOn, 40);
    if (last && !parseIsoDate(last)) return fail(400, 'Last contact must be a real date.');
    detail.lastContactOn = parseIsoDate(last);
    const grid = asText(body.serviceGridOn, 40);
    if (grid && !parseIsoDate(grid)) return fail(400, 'Service grid date must be a real date.');
    detail.serviceGridOn = parseIsoDate(grid);
    const placed = asText(body.placedOn, 40);
    if (placed && !parseIsoDate(placed)) return fail(400, 'Placed date must be a real date.');
    detail.placedOn = parseIsoDate(placed);
    applyStatusFromDates(detail);
    trackDaysTodo(detail, today);
    settle(detail, today);
    return { ok: true, detail: detail };
  }

  if (action === 'outreach') {
    const typed = asText(body.on, 40);
    if (typed && !parseIsoDate(typed)) return fail(400, 'Outreach date must be a real date.');
    const on = parseIsoDate(typed) || today;
    detail.outreachDates.push(on);
    detail.outreachDates.sort();
    if (!detail.lastContactOn || on > detail.lastContactOn) detail.lastContactOn = on;
    settle(detail, today);
    return { ok: true, detail: detail };
  }

  if (action === 'step') {
    const stepId = asText(body.stepId, 40);
    const step = detail.steps.find(function (item) { return item.id === stepId; });
    if (!step) return fail(400, 'That step is not on this referral.');
    if (Object.prototype.hasOwnProperty.call(body, 'owner')) step.owner = asText(body.owner, 120) || OWNER;
    if (Object.prototype.hasOwnProperty.call(body, 'doneOn')) {
      const doneOn = asText(body.doneOn, 40);
      if (doneOn && !parseIsoDate(doneOn)) return fail(400, 'Step date must be a real date.');
      if (doneOn && stepId === 'days' && (!detail.serviceGridOn || !detail.placedOn)) {
        return fail(400, 'Enter the service grid date and the placed date, save them, then record the days.');
      }
      step.doneOn = parseIsoDate(doneOn);
      if (stepId === 'remove' || stepId === 'removal') {
        if (step.doneOn) {
          detail.status = 'Removed';
          detail.removedOn = step.doneOn;
        } else if (detail.status === 'Removed') {
          detail.status = detail.placedOn ? 'Placed' : 'Active';
          detail.removedOn = '';
        }
      }
      if (stepId === 'mid' && step.doneOn) {
        addTodo(detail, {
          id: 'par-hia',
          title: 'Put Medicaid and insurance on the PAR/HIA form. Then follow the prior authorization steps for that funding source. This screen does not fill out the form.',
          owner: step.owner || OWNER,
          dueOn: step.doneOn,
        });
      }
    }
    settle(detail, today);
    return { ok: true, detail: detail };
  }

  if (action === 'assign') {
    const assignedTo = asText(body.assignedTo, 200);
    const assignedEmail = asText(body.assignedEmail, 200).toLowerCase();
    if (!assignedTo) return fail(400, 'Choose a staff member.');
    if (assignedEmail && assignedEmail.indexOf('@') === -1) return fail(400, 'Staff email is not valid.');
    detail.assignedTo = assignedTo;
    detail.assignedEmail = assignedEmail;
    ensureAssignWork(detail, actor, today);
    settle(detail, today);
    return { ok: true, detail: detail };
  }

  if (action === 'answered') {
    const wasRemoved = detail.status === 'Removed' || !!detail.removedOn;
    if (wasRemoved) {
      detail.removedOn = '';
      detail.steps.forEach(function (step) {
        if (step.id === 'remove' || step.id === 'removal') step.doneOn = '';
      });
    }
    detail.status = 'Answered';
    addTodo(detail, {
      id: 'to-sara-beth',
      title: wasRemoved
        ? 'This family answered after being removed, so they are back on the active list. Send them to Sara Beth for onboarding.'
        : 'Send this family to Sara Beth for onboarding.',
      owner: SARA_BETH,
      dueOn: today,
    });
    settle(detail, today);
    return { ok: true, detail: detail };
  }

  if (action === 'save-draft' || action === 'approve') {
    const draft = detail.drafts.find(function (item) { return item.id === asText(body.draftId, 40); });
    if (!draft) return fail(400, 'That message is not on this referral.');
    if (Object.prototype.hasOwnProperty.call(body, 'body')) {
      const nextBody = asText(body.body, 4000);
      if (nextBody !== draft.body && draft.status === 'approved') {
        draft.status = 'draft';
        draft.approvedBy = '';
        draft.approvedOn = '';
        dropOpenSendTodo(detail, draft.id);
      }
      draft.body = nextBody;
      refreshDraft(draft);
    }
    if (action === 'save-draft') {
      settle(detail, today);
      return { ok: true, detail: detail };
    }
    refreshDraft(draft);
    if (draft.incomplete) return fail(400, 'Finish this message before you approve it. Nothing is sent from this screen.');
    draft.status = 'approved';
    draft.approvedBy = actorName(actor);
    draft.approvedOn = today;
    addTodo(detail, {
      id: 'send-' + draft.id,
      title: 'Send this message yourself: ' + draft.label + '. Approving it does not send it.',
      owner: actorName(actor),
      dueOn: today,
    });
    settle(detail, today);
    return { ok: true, detail: detail };
  }

  if (action === 'todo-done') {
    const todo = detail.todos.find(function (item) { return item.id === asText(body.todoId, 80); });
    if (!todo) return fail(400, 'That to-do is not on this referral.');
    if (!todo.doneOn) todo.doneOn = today;
    settle(detail, today);
    return { ok: true, detail: detail };
  }

  if (action === 'add-todo') {
    const title = asText(body.title, 300);
    const owner = asText(body.owner, 120);
    const dueOn = parseIsoDate(asText(body.dueOn, 40));
    if (!title || !owner || !dueOn) return fail(400, 'A to-do needs a task, an owner, and a due date.');
    detail.todos.push({
      id: 'custom-' + randomBytes(4).toString('hex'),
      title: title,
      owner: owner,
      dueOn: dueOn,
      reminderOn: dueOn,
      doneOn: '',
    });
    settle(detail, today);
    return { ok: true, detail: detail };
  }

  return fail(400, 'That action is not available.');
}

export function projectReferral(detail: ReferralDetail, today: string, createdAt = ''): Record<string, unknown> {
  const placedDays = detail.serviceGridOn && detail.placedOn ? workingDaysBetween(detail.serviceGridOn, detail.placedOn) : null;
  const elapsed = detail.kind === 'ei' && detail.serviceGridOn ? workingDaysBetween(detail.serviceGridOn, today) : null;
  const daysToPlace = placedDays;
  const daysOverGoal = detail.kind === 'ei' && (
    (daysToPlace != null && daysToPlace > DAY_GOAL) ||
    (!detail.placedOn && elapsed != null && elapsed > DAY_GOAL)
  );
  const source = sourceByValue(detail.source);
  const firstOutreach = detail.outreachDates.length ? detail.outreachDates.slice().sort()[0] : '';
  const respondBy = firstOutreach ? addDays(firstOutreach, RESPONSE_DAYS) : '';
  const responseDaysLeft = respondBy && today <= respondBy ? calendarDaysAfter(today, respondBy) : null;
  return Object.assign({}, detail, {
    createdAt: createdAt,
    daysToPlace: daysToPlace,
    daysOverGoal: daysOverGoal,
    elapsedWorkingDays: elapsed,
    remainingWorkingDays: elapsed == null ? null : DAY_GOAL - elapsed,
    stepsUnconfirmed: detail.kind === 'private',
    dayGoal: detail.kind === 'ei' ? DAY_GOAL : null,
    outreachCount: detail.outreachDates.length,
    outreachesBeforeBottom: OUTREACHES_BEFORE_BOTTOM,
    responseDays: RESPONSE_DAYS,
    firstOutreachOn: firstOutreach,
    respondBy: respondBy,
    responseDaysLeft: responseDaysLeft,
    atBottom: detail.status === 'Bottom',
    sourceLabel: source ? source.label : detail.source,
    sourceHow: source ? source.how : '',
    sourceNote: source ? source.note : '',
    paperworkLine: source && source.paperwork ? PAPERWORK_LINE : '',
    dpExpiresOn: source && source.id === 'dp' && detail.receivedOn ? addDays(detail.receivedOn, 2) : '',
    todos: detail.todos.map(function (todo) {
      return Object.assign({}, todo, { overdue: !todo.doneOn && !!todo.dueOn && todo.dueOn < today });
    }),
    ReferrerName: detail.familyName,
    Status: detail.status,
    AssignedTo: detail.assignedTo,
    ReceivedOn: detail.receivedOn,
    Notes: detail.notes,
    Title: detail.familyName,
  });
}

const DAY_SORT = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];

export function summarizeOpenings(options: { today: string; staff: OpeningStaff[]; kinds: ProviderKind[] }): Record<string, unknown> {
  const today = parseIsoDate(options.today) || options.today;
  const kindByEmail = new Map<string, string[]>();
  (options.kinds || []).forEach(function (kind) {
    kindByEmail.set(kind.email.trim().toLowerCase(), cleanTherapies(kind.therapies));
  });
  let total = 0;
  const therapyCounts = new Map<string, number>();
  const providers: Array<Record<string, unknown>> = [];
  const slots: Array<Record<string, unknown>> = [];

  (options.staff || []).forEach(function (person) {
    const email = asText(person.staffEmail, 200).toLowerCase();
    const staffName = person.staffName || email;
    if (isReferralsMailbox(email, staffName)) return;
    const therapyIds = kindByEmail.get(email) || [];
    const therapyLabels = therapyIds.length ? therapyIds.map(function (id) { return THERAPY_LABELS[id] || id; }) : ['Not set'];
    let spots = 0;
    (person.openings || []).forEach(function (opening) {
      const day = asText(opening.day, 20);
      const time = asText(opening.time, 40);
      const status = asText(opening.status, 80);
      if (!day && !time) return;
      total += 1;
      spots += 1;
      therapyLabels.forEach(function (label) {
        therapyCounts.set(label, (therapyCounts.get(label) || 0) + 1);
      });
      slots.push({
        staffName: staffName,
        staffEmail: email,
        therapies: therapyLabels,
        therapyIds: therapyIds,
        day: day,
        time: time,
        status: status,
        spot: day && time ? day + ' at ' + time : (day || time),
      });
    });
    providers.push({
      staffName: staffName,
      staffEmail: email,
      therapyIds: therapyIds,
      therapies: therapyLabels,
      spots: spots,
    });
  });

  providers.sort(function (a, b) { return String(a.staffName).localeCompare(String(b.staffName)); });
  slots.sort(function (a, b) {
    const dayA = DAY_SORT.indexOf(String(a.day).toLowerCase());
    const dayB = DAY_SORT.indexOf(String(b.day).toLowerCase());
    return (dayA < 0 ? 9 : dayA) - (dayB < 0 ? 9 : dayB)
      || String(a.time).localeCompare(String(b.time))
      || String(a.staffName).localeCompare(String(b.staffName));
  });
  const byTherapy = Array.from(therapyCounts.keys()).sort().map(function (label) {
    return { label: label, count: therapyCounts.get(label) || 0 };
  });
  const disciplineCount = byTherapy.filter(function (row) { return row.label !== 'Not set' && row.count > 0; }).length;
  return {
    source: 'emme-scheduling-db',
    timeZone: TIME_ZONE,
    today: today,
    total: total,
    disciplineCount: disciplineCount,
    unlabeled: therapyCounts.get('Not set') || 0,
    byTherapy: byTherapy,
    providers: providers,
    slots: slots,
  };
}

function emptyStore(): StoreFile {
  return { version: 1, records: {}, kinds: [] };
}

function storageReady(): boolean {
  return !!getConfig().storageConnection;
}

function container() {
  const conn = getConfig().storageConnection;
  if (!conn) throw new Error('Referral storage is not configured.');
  return BlobServiceClient.fromConnectionString(conn).getContainerClient(getConfig().storageContainer);
}

async function readStore(): Promise<StoreFile> {
  if (!storageReady()) return emptyStore();
  const blob = container().getBlockBlobClient(BLOB_NAME);
  if (!(await blob.exists())) return emptyStore();
  let parsed: unknown;
  try {
    parsed = JSON.parse((await blob.downloadToBuffer()).toString('utf8'));
  } catch (_) {
    throw new Error('Referral storage could not be read.');
  }
  if (!parsed || typeof parsed !== 'object') throw new Error('Referral storage could not be read.');
  const file = parsed as Partial<StoreFile>;
  const records = file.records && typeof file.records === 'object' ? file.records : null;
  if (!records) throw new Error('Referral storage could not be read.');
  const kinds = Array.isArray(file.kinds) ? file.kinds.map(function (kind) {
    return {
      email: asText(kind && kind.email, 200).toLowerCase(),
      name: asText(kind && kind.name, 200),
      therapies: cleanTherapies(kind && kind.therapies),
    };
  }).filter(function (kind) { return kind.email.indexOf('@') !== -1; }) : [];
  return { version: 1, records: records as Record<string, ReferralDetail>, kinds: kinds };
}

async function writeStore(file: StoreFile): Promise<void> {
  const target = container();
  await target.createIfNotExists();
  const payload = JSON.stringify(file);
  await target.getBlockBlobClient(BLOB_NAME).upload(payload, Buffer.byteLength(payload), {
    blobHTTPHeaders: { blobContentType: 'application/json' },
  });
}

function mergeDetail(item: ListedReferral, stored: ReferralDetail | undefined, today: string): ReferralDetail {
  const saved = stored ? coerceStored(stored, item.id, today) : null;
  if (saved) {
    saved.id = item.id;
    return saved;
  }
  return fromSharePoint(item, today);
}

function validId(id: string): boolean {
  return /^[A-Za-z0-9_-]{1,80}$/.test(id);
}

async function defaultAvailability(): Promise<{ configured: boolean; asOf?: string; staff: OpeningStaff[] }> {
  const result = await openingsFromDashboard();
  return {
    configured: result.configured,
    staff: result.staff.map(function (person) {
      return {
        staffEmail: person.email,
        staffName: person.employee,
        openings: (person.summary && person.summary.openings) || [],
      };
    }),
  };
}

function stamp(detail: ReferralDetail, actor: Actor): void {
  detail.updatedAt = new Date().toISOString();
  detail.updatedBy = actorName(actor);
}

export function memoryList(): ListPort & { rows: ListedReferral[] } {
  const rows: ListedReferral[] = [];
  let next = 1;
  return {
    rows: rows,
    async list() {
      return rows.map(function (row) { return { id: row.id, createdAt: row.createdAt, fields: Object.assign({}, row.fields) }; });
    },
    async create(fields) {
      const row = { id: String(next++), createdAt: new Date().toISOString(), fields: Object.assign({}, fields) };
      rows.push(row);
      return { id: row.id, createdAt: row.createdAt, fields: Object.assign({}, row.fields) };
    },
    async update(id, fields) {
      const row = rows.find(function (item) { return item.id === id; });
      if (!row) throw new Error('Referral was not found.');
      row.fields = Object.assign({}, row.fields, fields);
    },
  };
}

export async function handleReferrals(context: FnContext, req: HttpRequest, actor: Actor, deps?: ReferralDeps): Promise<void> {
  const list = (deps && deps.list) || sharePointList;
  const today = (deps && deps.today) || denverToday();
  const verb = methodOf(req);
  const body = readBody(req);

  if (verb === 'GET' && queryOf(req, 'view') === 'openings') {
    const load = (deps && deps.loadAvailability) || defaultAvailability;
    const loaded = await load();
    if (!loaded.configured) {
      json(context, 503, { error: 'Scheduling storage is not configured. Set AZURE_STORAGE_CONNECTION_STRING.' });
      return;
    }
    const store = await readStore();
    const summaryToday = parseIsoDate(loaded.asOf || '') || today;
    json(context, 200, summarizeOpenings({ today: summaryToday, staff: loaded.staff || [], kinds: store.kinds }));
    return;
  }

  if (verb === 'GET') {
    const items = await list.list();
    const referrals = await locked(async function () {
      const store = await readStore();
      let dirty = false;
      const rows = items.map(function (item) {
        const stored = store.records[item.id];
        const detail = mergeDetail(item, stored, today);
        const statusBefore = detail.status;
        const todosBefore = JSON.stringify(detail.todos);
        settle(detail, today);
        if (stored && (detail.status !== statusBefore || JSON.stringify(detail.todos) !== todosBefore)) {
          detail.updatedAt = stored.updatedAt;
          store.records[item.id] = detail;
          dirty = true;
        }
        return { detail: detail, createdAt: item.createdAt || '', statusBefore: statusBefore, hadRecord: !!stored };
      });
      if (dirty) {
        await writeStore(store);
        for (let i = 0; i < rows.length; i++) {
          if (rows[i].hadRecord && rows[i].detail.status !== rows[i].statusBefore) {
            await list.update(rows[i].detail.id, sharePointFields(rows[i].detail));
          }
        }
      }
      return rows.map(function (row) { return projectReferral(row.detail, today, row.createdAt); });
    });
    referrals.sort(function (a, b) {
      const rank = function (status: unknown) {
        if (status === 'Removed') return 3;
        if (status === 'Bottom') return 2;
        return 1;
      };
      const byRank = rank(a.status) - rank(b.status);
      if (byRank) return byRank;
      const received = a.status === 'Bottom'
        ? String(a.receivedOn || '').localeCompare(String(b.receivedOn || ''))
        : String(b.receivedOn || '').localeCompare(String(a.receivedOn || ''));
      return received || String(a.familyName || '').localeCompare(String(b.familyName || ''));
    });
    json(context, 200, { referrals: referrals, today: today, timeZone: TIME_ZONE, sources: REFERRAL_SOURCES });
    return;
  }

  if (verb !== 'POST') {
    json(context, 405, { error: 'Method not allowed' });
    return;
  }
  if (!storageReady()) {
    json(context, 503, { error: 'Referral storage is not configured. Set AZURE_STORAGE_CONNECTION_STRING.' });
    return;
  }

  const action = asText(body.action, 40);
  if (action === 'provider-kinds') {
    const incoming = Array.isArray(body.providers) ? body.providers : [];
    if (!incoming.length) {
      json(context, 400, { error: 'Choose a staff member and a kind of therapy.' });
      return;
    }
    const cleaned: ProviderKind[] = [];
    for (let i = 0; i < incoming.length; i++) {
      const row = incoming[i] as { email?: unknown; name?: unknown; therapies?: unknown };
      const email = asText(row && row.email, 200).toLowerCase();
      if (!email || email.indexOf('@') === -1) {
        json(context, 400, { error: 'Staff email is not valid.' });
        return;
      }
      cleaned.push({ email: email, name: asText(row && row.name, 200) || email, therapies: cleanTherapies(row && row.therapies) });
    }
    await locked(async function () {
      const store = await readStore();
      cleaned.forEach(function (kind) {
        const index = store.kinds.findIndex(function (item) { return item.email === kind.email; });
        if (index === -1) store.kinds.push(kind);
        else store.kinds[index] = kind;
      });
      await writeStore(store);
    });
    json(context, 200, { ok: true });
    return;
  }

  const id = restParts(req)[0] || asText(body.id, 80);
  if (!id) {
    if (action && action !== 'save') {
      json(context, 400, { error: 'Choose a referral.' });
      return;
    }
    const built = buildNewReferral(body, today);
    if (!built.ok) {
      json(context, built.status, { error: built.error });
      return;
    }
    const created = await list.create(sharePointFields(built.detail));
    built.detail.id = created.id;
    stamp(built.detail, actor);
    await locked(async function () {
      const store = await readStore();
      store.records[created.id] = built.detail;
      await writeStore(store);
    });
    json(context, 201, { referral: projectReferral(built.detail, today, created.createdAt || '') });
    return;
  }

  if (!validId(id)) {
    json(context, 400, { error: 'That referral was not found.' });
    return;
  }

  const outcome = await locked(async function (): Promise<{ status: number; error: string } | { detail: ReferralDetail; createdAt: string }> {
    const items = await list.list();
    const item = items.find(function (row) { return row.id === id; });
    if (!item) return { status: 404, error: 'That referral was not found.' };
    const store = await readStore();
    const current = mergeDetail(item, store.records[id], today);
    const base = asText(body.baseUpdatedAt, 40);
    if ((current.updatedAt || '') !== base) {
      return { status: 409, error: 'Someone else saved this referral. Reload the page, then make your changes again.' };
    }
    const applied = applyReferralAction(current, Object.assign({}, body, { action: action || 'save' }), actor, today);
    if (!applied.ok) return { status: applied.status, error: applied.error };
    stamp(applied.detail, actor);
    applied.detail.id = id;
    store.records[id] = applied.detail;
    await writeStore(store);
    await list.update(id, sharePointFields(applied.detail));
    return { detail: applied.detail, createdAt: item.createdAt || '' };
  });
  if ('error' in outcome) {
    json(context, outcome.status, { error: outcome.error });
    return;
  }
  json(context, 200, { referral: projectReferral(outcome.detail, today, outcome.createdAt) });
}
