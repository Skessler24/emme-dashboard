import { BlobServiceClient } from '@azure/storage-blob';
import { randomBytes } from 'crypto';

/**
 * Setup answers and the files people upload with them, for every topic. They
 * live in the same storage account as schedules and messages. Each topic has
 * its own prefix (for example billing-intake/ or intake/onboarding/), so no
 * topic can read or overwrite another, and no SharePoint folder is needed.
 */

const CONTAINER = process.env.AZURE_STORAGE_CONTAINER || 'emme-reports';

export type Answer = { choice: string[]; text: string; unsure: boolean };

/** One block of a repeating group, such as a funding source. The stored property is still called "sources". */
export type GroupEntry = {
  id: string;
  kind: string;
  name: string;
  answers: Record<string, Answer>;
};

export type IntakeAnswers = {
  updatedAt: string;
  updatedBy: string;
  general: Record<string, Answer>;
  sources: GroupEntry[];
};

export type StoredFile = {
  id: string;
  /** "general" or a group entry id. */
  scope: string;
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

function container() {
  const conn = process.env.AZURE_STORAGE_CONNECTION_STRING || '';
  if (!conn) throw new Error('Setup storage is not configured.');
  return BlobServiceClient.fromConnectionString(conn).getContainerClient(CONTAINER);
}

function answersName(prefix: string): string { return prefix + 'answers.json'; }
function indexName(prefix: string): string { return prefix + 'files.json'; }
function fileName(prefix: string, id: string): string { return prefix + 'files/' + id; }

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

export function emptyAnswers(): IntakeAnswers {
  return { updatedAt: '', updatedBy: '', general: {}, sources: [] };
}

export async function readAnswers(prefix: string): Promise<IntakeAnswers> {
  const stored = await readJson<IntakeAnswers | null>(answersName(prefix), null);
  if (!stored || typeof stored !== 'object') return emptyAnswers();
  return {
    updatedAt: String(stored.updatedAt || ''),
    updatedBy: String(stored.updatedBy || ''),
    general: stored.general || {},
    sources: Array.isArray(stored.sources) ? stored.sources : [],
  };
}

export async function listFiles(prefix: string): Promise<StoredFile[]> {
  const rows = await readJson<StoredFile[]>(indexName(prefix), []);
  return Array.isArray(rows) ? rows : [];
}

/**
 * Saves the answers. When the answers were already saved by someone, the caller
 * must pass the updatedAt it last saw, so two people never overwrite each other.
 * Files that belong to a group entry that was removed are deleted too.
 */
export function saveAnswers(
  prefix: string,
  next: { general: Record<string, Answer>; sources: GroupEntry[] },
  baseUpdatedAt: string,
  by: string,
): Promise<IntakeAnswers> {
  return locked(async function () {
    const current = await readAnswers(prefix);
    if (current.updatedAt && current.updatedAt !== baseUpdatedAt) {
      throw new ConflictError(current.updatedBy, current.updatedAt);
    }
    const saved: IntakeAnswers = {
      updatedAt: new Date().toISOString(),
      updatedBy: by,
      general: next.general,
      sources: next.sources,
    };
    await writeJson(answersName(prefix), saved);
    const keep = new Set(['general'].concat(saved.sources.map(function (s) { return s.id; })));
    const files = await listFiles(prefix);
    const orphans = files.filter(function (file) { return !keep.has(file.scope); });
    if (orphans.length) {
      const target = container();
      for (let i = 0; i < orphans.length; i++) {
        await target.getBlockBlobClient(fileName(prefix, orphans[i].id)).deleteIfExists();
      }
      await writeJson(indexName(prefix), files.filter(function (file) { return keep.has(file.scope); }));
    }
    return saved;
  });
}

export function addFile(
  prefix: string,
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
    await target.getBlockBlobClient(fileName(prefix, file.id)).upload(data, data.length, {
      blobHTTPHeaders: { blobContentType: 'application/octet-stream' },
    });
    const files = await listFiles(prefix);
    files.push(file);
    await writeJson(indexName(prefix), files);
    return file;
  });
}

export async function readFile(prefix: string, id: string): Promise<{ file: StoredFile; data: Buffer } | null> {
  const files = await listFiles(prefix);
  const file = files.find(function (row) { return row.id === id; });
  if (!file) return null;
  const blob = container().getBlockBlobClient(fileName(prefix, file.id));
  if (!(await blob.exists())) return null;
  return { file, data: await blob.downloadToBuffer() };
}

export function removeFile(prefix: string, id: string): Promise<boolean> {
  return locked(async function () {
    const files = await listFiles(prefix);
    const file = files.find(function (row) { return row.id === id; });
    if (!file) return false;
    await container().getBlockBlobClient(fileName(prefix, file.id)).deleteIfExists();
    await writeJson(indexName(prefix), files.filter(function (row) { return row.id !== id; }));
    return true;
  });
}
