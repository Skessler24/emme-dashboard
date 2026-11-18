import { NOT_EMME_MESSAGE, emailAllowed, readPrincipal, resolveIsAdmin } from './auth';
import { FnContext, HttpRequest, asText, json, methodOf, readBody, restParts } from './http';
import {
  CONNECTION_GUIDE,
  FILE_KINDS,
  GENERAL_QUESTIONS,
  GENERAL_SECTIONS,
  Question,
  SOURCE_KINDS,
  SOURCE_QUESTIONS,
  SOURCE_SECTIONS,
  questionsForKind,
} from './billingQuestions';
import {
  Answer,
  ConflictError,
  FundingSource,
  addFile,
  listFiles,
  readAnswers,
  readFile,
  removeFile,
  saveAnswers,
} from './billingIntakeStore';
import { assessSource, buildPrompt, missingItems } from './billingPrompt';
import { randomBytes } from 'crypto';

export const MAX_FILE_BYTES = 6 * 1024 * 1024;
export const MAX_FILES = 100;
const MAX_SOURCES = 40;
const MAX_TEXT = 6000;

export const FILE_TYPES: Record<string, string> = {
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12',
  xls: 'application/vnd.ms-excel',
  csv: 'text/csv',
  pdf: 'application/pdf',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  doc: 'application/msword',
  txt: 'text/plain',
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
};

/** Billing setup holds sensitive notes, so only a signed-in human administrator may use it (the automation key is not enough). */
export async function requireHumanAdmin(context: FnContext, req: HttpRequest): Promise<{ email: string; name: string } | null> {
  const principal = readPrincipal(req);
  if (!principal || !principal.email) {
    json(context, 401, { error: 'Sign in required' });
    return null;
  }
  if (!emailAllowed(principal.email)) {
    json(context, 403, { error: NOT_EMME_MESSAGE });
    return null;
  }
  const admin = await resolveIsAdmin(principal.email);
  if (!admin.isAdmin) {
    json(context, 403, { error: 'Emme Admin access is required' });
    return null;
  }
  return { email: principal.email, name: principal.name };
}

export function cleanAnswers(raw: unknown, questions: Question[]): Record<string, Answer> {
  const out: Record<string, Answer> = {};
  const source = raw && typeof raw === 'object' ? raw as Record<string, any> : {};
  questions.forEach(function (question) {
    const given = source[question.id];
    if (!given || typeof given !== 'object') return;
    const valid = (question.options || []).map(function (option) { return option.id; });
    let picked: string[] = Array.isArray(given.choice) ? given.choice.map(String) : [];
    picked = picked.filter(function (id, index) { return valid.indexOf(id) !== -1 && picked.indexOf(id) === index; });
    if (question.kind === 'one') picked = picked.slice(0, 1);
    if (question.kind === 'text') picked = [];
    const answer: Answer = { choice: picked, text: asText(given.text, MAX_TEXT), unsure: given.unsure === true };
    if (answer.choice.length || answer.text || answer.unsure) out[question.id] = answer;
  });
  return out;
}

function cleanSources(raw: unknown): FundingSource[] | null {
  if (!Array.isArray(raw)) return null;
  const seen = new Set<string>();
  const out: FundingSource[] = [];
  for (let i = 0; i < raw.length && out.length < MAX_SOURCES; i++) {
    const item = raw[i] as Record<string, any>;
    if (!item || typeof item !== 'object') continue;
    const kind = String(item.kind || '');
    if (!SOURCE_KINDS.some(function (k) { return k.id === kind; })) continue;
    let id = String(item.id || '');
    if (!/^[a-z0-9-]{3,40}$/.test(id) || id === 'general' || seen.has(id)) id = 'src-' + randomBytes(5).toString('hex');
    seen.add(id);
    out.push({
      id,
      kind,
      name: asText(item.name, 120) || 'Unnamed funding source',
      answers: cleanAnswers(item.answers, questionsForKind(kind)),
    });
  }
  return out;
}

export function safeFileName(name: string): string {
  const base = String(name || '').split(/[\\/]/).pop() || '';
  const cleaned = base.replace(/[^A-Za-z0-9._ ()-]/g, '_').replace(/\.{2,}/g, '.').trim().slice(0, 120);
  return cleaned || 'file';
}

export function extensionOf(name: string): string {
  const dot = name.lastIndexOf('.');
  return dot === -1 ? '' : name.slice(dot + 1).toLowerCase();
}

function questionBank() {
  return {
    generalSections: GENERAL_SECTIONS,
    generalQuestions: GENERAL_QUESTIONS,
    sourceSections: SOURCE_SECTIONS,
    sourceQuestions: SOURCE_QUESTIONS,
    sourceKinds: SOURCE_KINDS,
    fileKinds: FILE_KINDS,
    connectionGuide: CONNECTION_GUIDE,
    limits: { maxFileBytes: MAX_FILE_BYTES, fileTypes: Object.keys(FILE_TYPES) },
  };
}

async function readState(context: FnContext): Promise<void> {
  const [answers, files] = await Promise.all([readAnswers(), listFiles()]);
  json(context, 200, Object.assign({ answers, files }, questionBank()));
}

async function save(context: FnContext, req: HttpRequest, who: { email: string }): Promise<void> {
  const body = readBody(req);
  const sources = cleanSources(body.sources);
  if (!sources) {
    json(context, 400, { error: 'Send the list of funding sources' });
    return;
  }
  const general = cleanAnswers(body.general, GENERAL_QUESTIONS);
  try {
    const saved = await saveAnswers({ general, sources }, asText(body.baseUpdatedAt, 40), who.email);
    json(context, 200, { ok: true, updatedAt: saved.updatedAt, updatedBy: saved.updatedBy, sources: saved.sources.map(function (s) { return { id: s.id, name: s.name }; }) });
  } catch (err) {
    if (err instanceof ConflictError) {
      json(context, 409, { error: err.message, updatedBy: err.updatedBy, updatedAt: err.updatedAt });
      return;
    }
    throw err;
  }
}

async function upload(context: FnContext, req: HttpRequest, who: { email: string }): Promise<void> {
  const body = readBody(req);
  if (body.confirmedNoClientInfo !== true) {
    json(context, 400, { error: 'Confirm that the file has no real client information in it. Replace names, birth dates, and member numbers with fake ones first.' });
    return;
  }
  const name = safeFileName(asText(body.name, 300));
  const ext = extensionOf(name);
  if (!FILE_TYPES[ext]) {
    json(context, 400, { error: 'That kind of file is not allowed. Use one of: ' + Object.keys(FILE_TYPES).join(', ') + '.' });
    return;
  }
  const kind = asText(body.kind, 40);
  if (!FILE_KINDS.some(function (item) { return item.id === kind; })) {
    json(context, 400, { error: 'Choose what kind of file this is' });
    return;
  }
  const encoded = typeof body.contentBase64 === 'string' ? body.contentBase64.replace(/\s+/g, '') : '';
  if (!encoded || encoded.length > Math.ceil(MAX_FILE_BYTES * 4 / 3) + 8 || !/^[A-Za-z0-9+/]+={0,2}$/.test(encoded)) {
    json(context, 400, { error: 'The file is missing, damaged, or larger than ' + Math.round(MAX_FILE_BYTES / 1024 / 1024) + ' megabytes' });
    return;
  }
  const data = Buffer.from(encoded, 'base64');
  if (!data.length || data.length > MAX_FILE_BYTES) {
    json(context, 400, { error: 'The file is empty or larger than ' + Math.round(MAX_FILE_BYTES / 1024 / 1024) + ' megabytes' });
    return;
  }
  const scope = asText(body.scope, 40);
  if (scope !== 'general') {
    const answers = await readAnswers();
    if (!answers.sources.some(function (source) { return source.id === scope; })) {
      json(context, 400, { error: 'Save the funding source first, then upload its files' });
      return;
    }
  }
  const existing = await listFiles();
  if (existing.length >= MAX_FILES) {
    json(context, 400, { error: 'The file limit was reached. Remove files you no longer need.' });
    return;
  }
  const file = await addFile({
    scope,
    kind,
    note: asText(body.note, 300),
    name,
    contentType: FILE_TYPES[ext],
    uploadedBy: who.email,
  }, data);
  json(context, 200, { ok: true, file });
}

export function contentDisposition(name: string): string {
  return 'attachment; filename="' + name.replace(/"/g, '') + '"; filename*=UTF-8\'\'' + encodeURIComponent(name);
}

async function download(context: FnContext, id: string): Promise<void> {
  const found = await readFile(id);
  if (!found) {
    json(context, 404, { error: 'File not found' });
    return;
  }
  context.res = {
    status: 200,
    headers: {
      // Always a download, never shown in the browser, so a hostile file cannot run on the dashboard.
      'Content-Type': 'application/octet-stream',
      'Content-Disposition': contentDisposition(found.file.name),
      'X-Content-Type-Options': 'nosniff',
      'Cache-Control': 'no-store',
    },
    body: found.data,
    isRaw: true,
  } as FnContext['res'];
}

async function prompt(context: FnContext): Promise<void> {
  const [answers, files] = await Promise.all([readAnswers(), listFiles()]);
  const known = new Set(['general'].concat(answers.sources.map(function (s) { return s.id; })));
  json(context, 200, {
    prompt: buildPrompt(answers, files.filter(function (file) { return known.has(file.scope); }), new Date()),
    assessments: answers.sources.map(assessSource),
    missing: missingItems(answers),
  });
}

export async function handleBillingIntake(context: FnContext, req: HttpRequest): Promise<void> {
  if (methodOf(req) === 'OPTIONS') {
    context.res = { status: 204, headers: { 'Content-Type': 'application/json' } };
    return;
  }
  const who = await requireHumanAdmin(context, req);
  if (!who) return;
  try {
    const verb = methodOf(req);
    const parts = restParts(req);
    if (!parts.length && verb === 'GET') return await readState(context);
    if (!parts.length && verb === 'PUT') return await save(context, req, who);
    if (parts[0] === 'prompt' && verb === 'GET') return await prompt(context);
    if (parts[0] === 'files' && !parts[1] && verb === 'POST') return await upload(context, req, who);
    if (parts[0] === 'files' && parts[1] && verb === 'GET') return await download(context, parts[1]);
    if (parts[0] === 'files' && parts[1] && verb === 'DELETE') {
      const removed = await removeFile(parts[1]);
      json(context, removed ? 200 : 404, removed ? { ok: true } : { error: 'File not found' });
      return;
    }
    json(context, 404, { error: 'Not found' });
  } catch (err) {
    // Do not echo storage details back to the browser.
    const message = err instanceof Error && /not configured/i.test(err.message) ? err.message : 'Billing setup could not be saved or read. Try again in a moment.';
    json(context, 500, { error: message });
  }
}
