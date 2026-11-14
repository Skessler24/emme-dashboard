import { BlobServiceClient } from '@azure/storage-blob';
import { randomBytes } from 'crypto';

/**
 * Setup answers and uploaded files for the Onboarding, Referrals, Staff
 * Information, and Staff Expectations tiles. They live in the same storage
 * account as schedules and messages, under tile-setup/<tile>/, one folder per
 * tile, so the tiles never see each other's answers.
 */

const CONTAINER = process.env.AZURE_STORAGE_CONTAINER || 'emme-reports';

export type Answer = { choice: string[]; text: string; unsure: boolean };

export type TileAnswers = {
  updatedAt: string;
  updatedBy: string;
  answers: Record<string, Answer>;
};

export type StoredFile = {
  id: string;
  kind: string;
  note: string;
  name: string;
  size: number;
  contentType: string;
  uploadedBy: string;
  uploadedAt: string;
};

export class ConflictError extends Error {
  constructor(public updatedBy: string, public updatedAt: string) {
    super('Someone else saved changes to this setup after you opened it. Reload the page to see their changes, then make your edits again.');
  }
}

let writeChain: Promise<void> = Promise.resolve();

function locked<T>(work: () => Promise<T>): Promise<T> {
  const run = writeChain.then(work, work);
  writeChain = run.then(function () { return undefined; }, function () { return undefined; });
  return run;
}

function answersName(tile: string): string { return 'tile-setup/' + tile + '/answers.json'; }
function indexName(tile: string): string { return 'tile-setup/' + tile + '/files.json'; }
function fileName(tile: string, id: string): string { return 'tile-setup/' + tile + '/files/' + id; }

function container() {
  const conn = process.env.AZURE_STORAGE_CONNECTION_STRING || '';
  if (!conn) throw new Error('Setup storage is not configured.');
  return BlobServiceClient.fromConnectionString(conn).getContainerClient(CONTAINER);
}

async function readJson<T>(name: string, fallback: T): Promise<T> {
  const blob = container().getBlockBlobClient(name);
  if (!(await blob.exists())) return fallback;
  try {
    return JSON.parse((await blob.downloadToBuffer()).toString('utf8')) as T;
  } catch (_) {
    return fallback;
  }
}

async function writeJson(name: string, value: unknown): Promise<void> {
  const target = container();
  await target.createIfNotExists();
  const payload = JSON.stringify(value);
  await target.getBlockBlobClient(name).upload(payload, Buffer.byteLength(payload), {
    blobHTTPHeaders: { blobContentType: 'application/json' },
  });
}

/** Returns null when nobody has saved this tile's answers yet. */
export async function readAnswers(tile: string): Promise<TileAnswers | null> {
  const stored = await readJson<TileAnswers | null>(answersName(tile), null);
  if (!stored || typeof stored !== 'object' || !stored.updatedAt) return null;
  return {
    updatedAt: String(stored.updatedAt),
    updatedBy: String(stored.updatedBy || ''),
    answers: stored.answers && typeof stored.answers === 'object' ? stored.answers : {},
  };
}

export async function listFiles(tile: string): Promise<StoredFile[]> {
  const rows = await readJson<StoredFile[]>(indexName(tile), []);
  return Array.isArray(rows) ? rows : [];
}

/**
 * Saves the answers. When answers were already saved by someone, the caller
 * must pass the updatedAt it last saw, so two people never overwrite each other.
 */
export function saveAnswers(
  tile: string,
  answers: Record<string, Answer>,
  baseUpdatedAt: string,
  by: string,
): Promise<TileAnswers> {
  return locked(async function () {
    const current = await readAnswers(tile);
    if (current && current.updatedAt !== baseUpdatedAt) {
      throw new ConflictError(current.updatedBy, current.updatedAt);
    }
    const saved: TileAnswers = { updatedAt: new Date().toISOString(), updatedBy: by, answers };
    await writeJson(answersName(tile), saved);
    return saved;
  });
}

export function addFile(
  tile: string,
  meta: Omit<StoredFile, 'id' | 'uploadedAt' | 'size'>,
  data: Buffer,
): Promise<StoredFile> {
  return locked(async function () {
    const file: StoredFile = Object.assign({}, meta, {
      id: randomBytes(8).toString('hex'),
      size: data.length,
      uploadedAt: new Date().toISOString(),
    });
    const target = container();
    await target.createIfNotExists();
    await target.getBlockBlobClient(fileName(tile, file.id)).upload(data, data.length, {
      blobHTTPHeaders: { blobContentType: 'application/octet-stream' },
    });
    const files = await listFiles(tile);
    files.push(file);
    await writeJson(indexName(tile), files);
    return file;
  });
}

export async function readFile(tile: string, id: string): Promise<{ file: StoredFile; data: Buffer } | null> {
  const files = await listFiles(tile);
  const file = files.find(function (row) { return row.id === id; });
  if (!file) return null;
  const blob = container().getBlockBlobClient(fileName(tile, file.id));
  if (!(await blob.exists())) return null;
  return { file, data: await blob.downloadToBuffer() };
}

export function removeFile(tile: string, id: string): Promise<boolean> {
  return locked(async function () {
    const files = await listFiles(tile);
    const file = files.find(function (row) { return row.id === id; });
    if (!file) return false;
    await container().getBlockBlobClient(fileName(tile, file.id)).deleteIfExists();
    await writeJson(indexName(tile), files.filter(function (row) { return row.id !== id; }));
    return true;
  });
}
