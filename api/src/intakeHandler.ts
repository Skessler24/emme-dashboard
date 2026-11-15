import { NOT_EMME_MESSAGE, emailAllowed, readPrincipal, resolveIsAdmin } from './auth';
import { FnContext, HttpRequest, asText, json, methodOf, readBody, restParts } from './http';
import { Question, TopicDef } from './intakeTypes';
import { getTopic } from './intakeTopics';
import {
  Answer,
  ConflictError,
  GroupEntry,
  addFile,
  listFiles,
  readAnswers,
  readFile,
  removeFile,
  saveAnswers,
} from './intakeStore';
import { assessAll, buildPrompt, missingItems, questionsForKind } from './intakePrompt';
import { randomBytes } from 'crypto';

export const MAX_FILE_BYTES = 6 * 1024 * 1024;
export const MAX_FILES = 100;
const MAX_SOURCES = 40;
const MAX_TEXT = 6000;

const FILE_TYPES: Record<string, string> = {
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
async function requireHumanAdmin(context: FnContext, req: HttpRequest): Promise<{ email: string; name: string } | null> {
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

function cleanAnswers(raw: unknown, questions: Question[]): Record<string, Answer> {
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

function cleanSources(topic: TopicDef, raw: unknown): GroupEntry[] | null {
  const group = topic.group;
  // A topic without a group never stores any entries; ignore whatever was sent.
  if (!group) return [];
  if (!Array.isArray(raw)) return null;
  const seen = new Set<string>();
  const out: GroupEntry[] = [];
  for (let i = 0; i < raw.length && out.length < MAX_SOURCES; i++) {
    const item = raw[i] as Record<string, any>;
    if (!item || typeof item !== 'object') continue;
    const kind = String(item.kind || '');
    if (!group.kinds.some(function (k) { return k.id === kind; })) continue;
    let id = String(item.id || '');
    if (!/^[a-z0-9-]{3,40}$/.test(id) || id === 'general' || seen.has(id)) id = 'src-' + randomBytes(5).toString('hex');
    seen.add(id);
    out.push({
      id,
      kind,
      name: asText(item.name, 120) || 'Unnamed ' + group.singular,
      answers: cleanAnswers(item.answers, questionsForKind(topic, kind)),
    });
  }
  return out;
}

export function safeFileName(name: string): string {
  const base = String(name || '').split(/[\\/]/).pop() || '';
  const cleaned = base.replace(/[^A-Za-z0-9._ ()-]/g, '_').replace(/\.{2,}/g, '.').trim().slice(0, 120);
  return cleaned || 'file';
}

function extensionOf(name: string): string {
  const dot = name.lastIndexOf('.');
  return dot === -1 ? '' : name.slice(dot + 1).toLowerCase();
}

function questionBank(topic: TopicDef) {
  const group = topic.group;
  return {
    topic: {
      id: topic.id,
      title: topic.title,
      guideIntro: topic.guideIntro,
      group: group ? {
        singular: group.singular,
        plural: group.plural,
        intro: group.intro,
        namePlaceholder: group.namePlaceholder,
        kinds: group.kinds,
        sections: group.sections,
        questions: group.questions,
      } : null,
      hasAssessment: !!topic.assess,
    },
    generalSections: topic.sections,
    generalQuestions: topic.questions,
    fileKinds: topic.fileKinds,
    connectionGuide: topic.guide,
    limits: { maxFileBytes: MAX_FILE_BYTES, fileTypes: Object.keys(FILE_TYPES) },
  };
}

async function readState(context: FnContext, topic: TopicDef): Promise<void> {
  const [answers, files] = await Promise.all([readAnswers(topic.storagePrefix), listFiles(topic.storagePrefix)]);
  json(context, 200, Object.assign({ answers, files }, questionBank(topic)));
}

async function save(context: FnContext, req: HttpRequest, topic: TopicDef, who: { email: string }): Promise<void> {
  const body = readBody(req);
  const sources = cleanSources(topic, body.sources);
  if (!sources) {
    json(context, 400, { error: 'Send the list of ' + (topic.group ? topic.group.plural : 'entries') });
    return;
  }
  const general = cleanAnswers(body.general, topic.questions);
  try {
    const saved = await saveAnswers(topic.storagePrefix, { general, sources }, asText(body.baseUpdatedAt, 40), who.email);
    json(context, 200, { ok: true, updatedAt: saved.updatedAt, updatedBy: saved.updatedBy, sources: saved.sources.map(function (s) { return { id: s.id, name: s.name }; }) });
  } catch (err) {
    if (err instanceof ConflictError) {
      json(context, 409, { error: err.message, updatedBy: err.updatedBy, updatedAt: err.updatedAt });
      return;
    }
    throw err;
  }
}

async function upload(context: FnContext, req: HttpRequest, topic: TopicDef, who: { email: string }): Promise<void> {
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
  if (!topic.fileKinds.some(function (item) { return item.id === kind; })) {
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
    const answers = await readAnswers(topic.storagePrefix);
    if (!answers.sources.some(function (source) { return source.id === scope; })) {
      json(context, 400, { error: 'Save the ' + (topic.group ? topic.group.singular : 'entry') + ' first, then upload its files' });
      return;
    }
  }
  const existing = await listFiles(topic.storagePrefix);
  if (existing.length >= MAX_FILES) {
    json(context, 400, { error: 'The file limit was reached. Remove files you no longer need.' });
    return;
  }
  const file = await addFile(topic.storagePrefix, {
    scope,
    kind,
    note: asText(body.note, 300),
    name,
    contentType: FILE_TYPES[ext],
    uploadedBy: who.email,
  }, data);
  json(context, 200, { ok: true, file });
}

function contentDisposition(name: string): string {
  return 'attachment; filename="' + name.replace(/"/g, '') + '"; filename*=UTF-8\'\'' + encodeURIComponent(name);
}

async function download(context: FnContext, topic: TopicDef, id: string): Promise<void> {
  const found = await readFile(topic.storagePrefix, id);
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

async function prompt(context: FnContext, topic: TopicDef): Promise<void> {
  const [answers, files] = await Promise.all([readAnswers(topic.storagePrefix), listFiles(topic.storagePrefix)]);
  const known = new Set(['general'].concat(answers.sources.map(function (s) { return s.id; })));
  json(context, 200, {
    prompt: buildPrompt(topic, answers, files.filter(function (file) { return known.has(file.scope); }), new Date()),
    assessments: assessAll(topic, answers),
    missing: missingItems(topic, answers),
  });
}

/**
 * Routes: /api/intake/{topic}/state (GET, PUT), /prompt (GET), /files (POST),
 * /files/{id} (GET, DELETE). Every address is spelled out so nothing depends on
 * an empty catch-all path.
 */
export async function handleIntake(context: FnContext, req: HttpRequest): Promise<void> {
  if (methodOf(req) === 'OPTIONS') {
    context.res = { status: 204, headers: { 'Content-Type': 'application/json' } };
    return;
  }
  const who = await requireHumanAdmin(context, req);
  if (!who) return;
  const topic = getTopic(asText(req.params && req.params.topic, 60));
  if (!topic) {
    json(context, 404, { error: 'Unknown setup topic' });
    return;
  }
  try {
    const verb = methodOf(req);
    const parts = restParts(req);
    const first = parts[0] || 'state';
    if (first === 'state' && !parts[1] && verb === 'GET') return await readState(context, topic);
    if (first === 'state' && !parts[1] && verb === 'PUT') return await save(context, req, topic, who);
    if (first === 'prompt' && verb === 'GET') return await prompt(context, topic);
    if (first === 'files' && !parts[1] && verb === 'POST') return await upload(context, req, topic, who);
    if (first === 'files' && parts[1] && verb === 'GET') return await download(context, topic, parts[1]);
    if (first === 'files' && parts[1] && verb === 'DELETE') {
      const removed = await removeFile(topic.storagePrefix, parts[1]);
      json(context, removed ? 200 : 404, removed ? { ok: true } : { error: 'File not found' });
      return;
    }
    json(context, 404, { error: 'Not found' });
  } catch (err) {
    // Do not echo storage details back to the browser.
    const message = err instanceof Error && /not configured/i.test(err.message) ? err.message : 'The setup answers could not be saved or read. Try again in a moment.';
    json(context, 500, { error: message });
  }
}
