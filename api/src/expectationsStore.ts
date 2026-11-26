import { BlobServiceClient } from '@azure/storage-blob';
import { randomBytes } from 'crypto';
import { VisitFrequency } from './expectationsRules';

/**
 * Live mileage board, travel approvals, and reminder checks.
 * Stored under expectations/ so schedules and other tiles are left alone.
 * Note text is never written here.
 */

const CONTAINER = process.env.AZURE_STORAGE_CONTAINER || 'emme-reports';
const BLOB = 'expectations/staff-board.json';

export type HomeRecord = {
  address: string;
  lat: number;
  lon: number;
  mapLink: string;
  name: string;
};

export type TripRecord = {
  id: string;
  staffEmail: string;
  staffName: string;
  date: string;
  startPlace: string;
  endPlace: string;
  miles: number;
};

export type TravelRequest = {
  id: string;
  staffEmail: string;
  staffName: string;
  clientAddress: string;
  miles: number;
  status: 'pending' | 'approved' | 'denied';
  decidedBy: string;
  createdAt: string;
};

export type FollowUp = {
  noteId: string;
  staffEmail: string;
  noteName: string;
  date: string;
  kind: 'eval' | 'report';
  followUp: 'unicorn' | 'par' | 'done';
};

export type OutreachRecord = {
  choice: 'email' | 'call';
  at: string;
  by: string;
};

export type BoardState = {
  homes: Record<string, HomeRecord>;
  trips: TripRecord[];
  hours: Record<string, number>;
  requests: TravelRequest[];
  emailChecks: Record<string, Record<string, boolean>>;
  frequency: Record<string, VisitFrequency>;
  followUps: FollowUp[];
  outreach: Record<string, OutreachRecord>;
};

export class ExpectationsStoreError extends Error {
  status: number;
  constructor(message: string, status: number) {
    super(message);
    this.status = status;
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
  if (!conn) throw new ExpectationsStoreError('Mileage storage is not configured.', 503);
  return BlobServiceClient.fromConnectionString(conn).getContainerClient(CONTAINER);
}

export function emptyBoard(): BoardState {
  return { homes: {}, trips: [], hours: {}, requests: [], emailChecks: {}, frequency: {}, followUps: [], outreach: {} };
}

export function newId(prefix: string): string {
  return prefix + '-' + randomBytes(4).toString('hex');
}

async function readBoard(): Promise<BoardState> {
  const blob = container().getBlockBlobClient(BLOB);
  if (!(await blob.exists())) return emptyBoard();
  try {
    const parsed = JSON.parse((await blob.downloadToBuffer()).toString('utf8')) as Partial<BoardState>;
    const blank = emptyBoard();
    return {
      homes: parsed.homes || blank.homes,
      trips: Array.isArray(parsed.trips) ? parsed.trips : [],
      hours: parsed.hours || blank.hours,
      requests: Array.isArray(parsed.requests) ? parsed.requests : [],
      emailChecks: parsed.emailChecks || blank.emailChecks,
      frequency: parsed.frequency || blank.frequency,
      followUps: Array.isArray(parsed.followUps) ? parsed.followUps : [],
      outreach: parsed.outreach || blank.outreach,
    };
  } catch (_) {
    return emptyBoard();
  }
}

async function writeBoard(state: BoardState): Promise<void> {
  const target = container();
  await target.createIfNotExists();
  const blob = target.getBlockBlobClient(BLOB);
  const body = JSON.stringify(state);
  await blob.upload(body, Buffer.byteLength(body), {
    blobHTTPHeaders: { blobContentType: 'application/json' },
    metadata: { kind: 'expectations' },
  });
}

export function updateBoard(change: (state: BoardState) => void): Promise<BoardState> {
  return locked(async function () {
    const state = await readBoard();
    change(state);
    await writeBoard(state);
    return state;
  });
}

export function loadBoard(): Promise<BoardState> {
  return readBoard();
}
