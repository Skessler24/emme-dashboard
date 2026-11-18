import { FnContext, HttpRequest, asText, json, methodOf, readBody, restParts } from './http';
import { FILE_TYPES, MAX_FILE_BYTES, cleanAnswers, contentDisposition, extensionOf, requireHumanAdmin, safeFileName } from './billingIntake';
import { TILE_SETUPS, TileDef, tileSetup } from './tileSetupQuestions';
import { ConflictError, addFile, listFiles, readAnswers, readFile, removeFile, saveAnswers } from './tileSetupStore';
import { buildTilePrompt, missingItems, seededAnswers } from './tilePrompt';

export const MAX_TILE_FILES = 40;

/** What the browser needs to draw the form. The first-draft answers are sent as answers, not as part of the questions. */
function bank(tile: TileDef) {
  return {
    tile: { id: tile.id, title: tile.title },
    sections: tile.sections,
    questions: tile.questions.map(function (question) {
      const copy = Object.assign({}, question);
      delete copy.seed;
      return copy;
    }),
    fileKinds: tile.fileKinds,
    limits: { maxFileBytes: MAX_FILE_BYTES, fileTypes: Object.keys(FILE_TYPES) },
  };
}

async function readState(context: FnContext, tile: TileDef): Promise<void> {
  const [saved, files] = await Promise.all([readAnswers(tile.id), listFiles(tile.id)]);
  json(context, 200, Object.assign({
    answers: saved
      ? { updatedAt: saved.updatedAt, updatedBy: saved.updatedBy, answers: saved.answers, firstDraft: false }
      : { updatedAt: '', updatedBy: '', answers: seededAnswers(tile), firstDraft: true },
    files,
  }, bank(tile)));
}

async function save(context: FnContext, req: HttpRequest, tile: TileDef, who: { email: string }): Promise<void> {
  const body = readBody(req);
  const answers = cleanAnswers(body.answers, tile.questions);
  try {
    const saved = await saveAnswers(tile.id, answers, asText(body.baseUpdatedAt, 40), who.email);
    json(context, 200, { ok: true, updatedAt: saved.updatedAt, updatedBy: saved.updatedBy });
  } catch (err) {
    if (err instanceof ConflictError) {
      json(context, 409, { error: err.message, updatedBy: err.updatedBy, updatedAt: err.updatedAt });
      return;
    }
    throw err;
  }
}

async function upload(context: FnContext, req: HttpRequest, tile: TileDef, who: { email: string }): Promise<void> {
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
  if (!tile.fileKinds.some(function (item) { return item.id === kind; })) {
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
  if ((await listFiles(tile.id)).length >= MAX_TILE_FILES) {
    json(context, 400, { error: 'The file limit was reached. Remove files you no longer need.' });
    return;
  }
  const file = await addFile(tile.id, {
    kind,
    note: asText(body.note, 300),
    name,
    contentType: FILE_TYPES[ext],
    uploadedBy: who.email,
  }, data);
  json(context, 200, { ok: true, file });
}

async function download(context: FnContext, tile: TileDef, id: string): Promise<void> {
  const found = await readFile(tile.id, id);
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

async function prompt(context: FnContext, tile: TileDef): Promise<void> {
  const [saved, files] = await Promise.all([readAnswers(tile.id), listFiles(tile.id)]);
  const answers = saved ? saved.answers : seededAnswers(tile);
  json(context, 200, {
    prompt: buildTilePrompt(tile, saved, files, new Date()),
    missing: missingItems(tile, answers),
  });
}

export async function handleTileSetup(context: FnContext, req: HttpRequest): Promise<void> {
  if (methodOf(req) === 'OPTIONS') {
    context.res = { status: 204, headers: { 'Content-Type': 'application/json' } };
    return;
  }
  const who = await requireHumanAdmin(context, req);
  if (!who) return;
  try {
    const verb = methodOf(req);
    const parts = restParts(req);
    if (!parts.length && verb === 'GET') {
      json(context, 200, { tiles: TILE_SETUPS.map(function (tile) { return { id: tile.id, title: tile.title }; }) });
      return;
    }
    const tile = tileSetup(parts[0]);
    if (!tile) {
      json(context, 404, { error: 'Not found' });
      return;
    }
    const rest = parts.slice(1);
    if (!rest.length && verb === 'GET') return await readState(context, tile);
    if (!rest.length && verb === 'PUT') return await save(context, req, tile, who);
    if (rest[0] === 'prompt' && verb === 'GET') return await prompt(context, tile);
    if (rest[0] === 'files' && !rest[1] && verb === 'POST') return await upload(context, req, tile, who);
    if (rest[0] === 'files' && rest[1] && verb === 'GET') return await download(context, tile, rest[1]);
    if (rest[0] === 'files' && rest[1] && verb === 'DELETE') {
      const removed = await removeFile(tile.id, rest[1]);
      json(context, removed ? 200 : 404, removed ? { ok: true } : { error: 'File not found' });
      return;
    }
    json(context, 404, { error: 'Not found' });
  } catch (err) {
    // Do not echo storage details back to the browser.
    const message = err instanceof Error && /not configured/i.test(err.message) ? err.message : 'The setup could not be saved or read. Try again in a moment.';
    json(context, 500, { error: message });
  }
}
