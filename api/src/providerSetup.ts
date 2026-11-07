import { GraphError, getSharePointSiteId, getUserDisplayName, graphGet, graphGetText, graphIsConfigured } from './graphClient';

/**
 * Finds a provider's existing SharePoint folder by their emmeco.org email.
 * This only reads folders that are already there. It does not create folders or files.
 * An optional emme-provider.json in the folder can name the email directly.
 * Otherwise the folder or workbook name is matched to the provider's name and email.
 */

export const MILEAGE_ROOT = ['EMME', 'PAYROLL & MILEAGE', 'Mileage'];
export const SCHEDULE_ROOT = ['EMME', 'LEADERSHIP TEAM', 'EMME Processes', 'Referral Process', 'Provider Schedules'];
export const RS_OBSERVATION_ROOT = ['EMME', 'LEADERSHIP TEAM', 'EMME Processes', 'RS_Observation Process'];
const MARKER = 'emme-provider.json';

export type FoundProviderFile = {
  folder: string;
  fileName: string;
  path: string;
  matchedBy: 'email' | 'name';
};

type DriveItem = { name: string; isFolder: boolean };

export function emmeEmail(value: string): string {
  const email = String(value || '').trim().toLowerCase();
  return /@emmeco\.org$/.test(email) ? email : '';
}

function normName(value: string): string {
  return String(value || '').toLowerCase().replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
}

function compact(value: string): string {
  return normName(value).replace(/\s/g, '');
}

function shortCode(name: string): string {
  const parts = normName(name).split(' ').filter(Boolean);
  if (parts.length < 2) return parts[0] || '';
  return parts[0].slice(0, 3) + parts[parts.length - 1].slice(0, 3);
}

function labelOf(value: string): string {
  return compact(String(value || '')
    .replace(/\.xlsx$/i, '')
    .replace(/\b(schedule|mileage|reimbursement|form|20\d{2})\b/gi, ' '));
}

export function providerMatchScore(candidate: string, staffName: string, email: string): number {
  const folder = labelOf(candidate);
  const name = compact(staffName);
  const code = shortCode(staffName);
  const local = compact((email || '').split('@')[0] || '');
  if (!folder) return 0;
  if (local && folder === local) return 100;
  if (name && folder === name) return 95;
  if (code && folder === code) return 90;
  return 0;
}

let librarySiteCache = '';

async function librarySiteId(): Promise<string> {
  if (librarySiteCache) return librarySiteCache;
  const tries: string[] = [];
  try {
    const personal = await graphGet('/sites/emmeco031-my.sharepoint.com:/personal/jenriat_emmeco_org?$select=id');
    if (personal && personal.id) tries.push(String(personal.id));
  } catch (_) { /* the configured site is the fallback */ }
  try {
    const configured = await getSharePointSiteId();
    if (tries.indexOf(configured) === -1) tries.push(configured);
  } catch (_) { /* the personal library may still be reachable */ }
  if (!tries.length) throw new GraphError(503, 'SharePoint is not configured.');
  const folder = MILEAGE_ROOT.map(encodeURIComponent).join('/');
  for (let i = 0; i < tries.length; i++) {
    try {
      await graphGet('/sites/' + tries[i] + '/drive/root:/' + folder + ':/children?$select=name&$top=1');
      librarySiteCache = tries[i];
      return librarySiteCache;
    } catch (err) {
      if (err instanceof GraphError && (err.status === 401 || err.status === 403 || err.status === 404)) continue;
      throw err;
    }
  }
  librarySiteCache = tries[0];
  return librarySiteCache;
}

function drivePath(segments: string[]): string {
  return segments.map(encodeURIComponent).join('/');
}

async function listChildren(segments: string[]): Promise<DriveItem[]> {
  const siteId = await librarySiteId();
  const path = '/sites/' + siteId + '/drive/root:/' + drivePath(segments) + ':/children?$select=name,folder,file';
  try {
    const data = await graphGet(path);
    return ((data && data.value) || []).map(function (item: { name?: string; folder?: unknown }) {
      return { name: String(item.name || ''), isFolder: !!item.folder };
    });
  } catch (err) {
    if (err instanceof GraphError && (err.status === 404 || /itemNotFound/i.test(err.message))) return [];
    throw err;
  }
}

async function markerEmail(segments: string[]): Promise<string> {
  const siteId = await librarySiteId();
  const path = '/sites/' + siteId + '/drive/root:/' + drivePath(segments.concat([MARKER])) + ':/content';
  try {
    const text = await graphGetText(path);
    const parsed = JSON.parse(text) as { email?: string };
    return emmeEmail(String(parsed && parsed.email || ''));
  } catch (_) {
    return '';
  }
}

function workbookName(files: DriveItem[], kind: 'mileage' | 'schedule'): string {
  const books = files.filter(function (file) {
    return !file.isFolder && /\.xlsx$/i.test(file.name) && file.name.indexOf('~$') !== 0;
  });
  const preferred = books.find(function (file) {
    return kind === 'mileage' ? /mileage/i.test(file.name) : /schedule/i.test(file.name);
  });
  return (preferred || books[0] || { name: '' }).name;
}

export async function findExistingProviderFile(root: string[], email: string, displayName: string, kind: 'mileage' | 'schedule'): Promise<FoundProviderFile | null> {
  const providerEmail = emmeEmail(email);
  const entries = await listChildren(root);
  let emailFolder = '';
  for (let i = 0; i < entries.length; i++) {
    const item = entries[i];
    if (!item.isFolder) continue;
    const marked = await markerEmail(root.concat([item.name]));
    if (marked && marked === providerEmail) {
      emailFolder = item.name;
      break;
    }
  }
  let bestFolder = emailFolder;
  let bestScore = emailFolder ? 100 : 0;
  let matchedBy: 'email' | 'name' = emailFolder ? 'email' : 'name';
  if (!bestFolder) {
    entries.forEach(function (item) {
      if (!item.isFolder) return;
      const score = Math.max(providerMatchScore(item.name, displayName, providerEmail), providerMatchScore(item.name, displayName, ''));
      if (score > bestScore) {
        bestScore = score;
        bestFolder = item.name;
        matchedBy = providerEmail && compact(item.name) === compact(providerEmail.split('@')[0] || '') ? 'email' : 'name';
      }
    });
  }
  if (bestFolder && (emailFolder || bestScore >= 90)) {
    const inside = await listChildren(root.concat([bestFolder]));
    const fileName = workbookName(inside, kind);
    const segments = fileName ? root.concat([bestFolder, fileName]) : root.concat([bestFolder]);
    return { folder: bestFolder, fileName, path: '/' + segments.join('/'), matchedBy };
  }
  const rootBooks = entries.filter(function (item) { return !item.isFolder; });
  let bestFile = '';
  let fileScore = 0;
  rootBooks.forEach(function (item) {
    const score = providerMatchScore(item.name, displayName, providerEmail);
    if (score > fileScore) {
      fileScore = score;
      bestFile = item.name;
    }
  });
  if (!bestFile || fileScore < 90 || !/\.xlsx$/i.test(bestFile)) return null;
  return {
    folder: '',
    fileName: bestFile,
    path: '/' + root.concat([bestFile]).join('/'),
    matchedBy: 'name',
  };
}

export async function searchProviderLibraries(email: string, displayName: string): Promise<{
  email: string;
  name: string;
  mileage: FoundProviderFile | null;
  schedule: FoundProviderFile | null;
}> {
  const providerEmail = emmeEmail(email);
  if (!providerEmail) throw new GraphError(400, 'Use the provider’s emmeco.org email.');
  let name = String(displayName || '').trim();
  if (!name && graphIsConfigured()) name = await getUserDisplayName(providerEmail);
  if (!name) name = providerEmail;
  const mileage = await findExistingProviderFile(MILEAGE_ROOT, providerEmail, name, 'mileage');
  const schedule = await findExistingProviderFile(SCHEDULE_ROOT, providerEmail, name, 'schedule');
  return { email: providerEmail, name, mileage, schedule };
}

export type ObservationItem = {
  name: string;
  kind: 'folder' | 'file';
  lastModified: string;
  webUrl: string;
  size: number | null;
};

function assertFolderSegment(segment: string): string {
  const name = String(segment || '').trim();
  if (!name || name === '.' || name === '..' || /[\\/]/.test(name) || name.length > 180) {
    throw new GraphError(400, 'Folder path is not valid');
  }
  return name;
}

let observationSiteCache = '';

async function folderOnSite(siteId: string, segments: string[]): Promise<{ webUrl: string } | null> {
  try {
    const hit = await graphGet('/sites/' + siteId + '/drive/root:/' + drivePath(segments) + '?$select=id,webUrl,folder');
    if (!hit || !hit.folder) return null;
    return { webUrl: String(hit.webUrl || '') };
  } catch (err) {
    if (err instanceof GraphError && (err.status === 404 || /itemNotFound/i.test(err.message))) return null;
    throw err;
  }
}

/**
 * Reflective Supervision files live in
 * EMME/LEADERSHIP TEAM/EMME Processes/RS_Observation Process.
 * The personal library is checked first, then the site already used for mileage.
 */
async function observationSiteId(): Promise<string> {
  if (observationSiteCache) return observationSiteCache;
  const tries: string[] = [];
  try {
    const personal = await graphGet('/sites/emmeco031-my.sharepoint.com:/personal/jenriat_emmeco_org?$select=id');
    if (personal && personal.id) tries.push(String(personal.id));
  } catch (_) { /* the mileage library site is the fallback */ }
  try {
    const known = await librarySiteId();
    if (tries.indexOf(known) === -1) tries.push(known);
  } catch (_) { /* the personal library may still contain the folder */ }
  let last: unknown = null;
  for (let i = 0; i < tries.length; i++) {
    try {
      const found = await folderOnSite(tries[i], RS_OBSERVATION_ROOT);
      if (found) {
        observationSiteCache = tries[i];
        return observationSiteCache;
      }
    } catch (err) {
      last = err;
      if (err instanceof GraphError && (err.status === 401 || err.status === 403 || err.status === 404)) continue;
      throw err;
    }
  }
  if (last instanceof GraphError && last.status !== 404) throw last;
  throw new GraphError(404, 'The Reflective Supervision folder was not found in SharePoint.');
}

async function listDriveChildren(siteId: string, segments: string[]): Promise<ObservationItem[]> {
  let next: string | null = '/sites/' + siteId + '/drive/root:/' + drivePath(segments) +
    ':/children?$select=name,lastModifiedDateTime,folder,file,webUrl,size&$top=200';
  const items: ObservationItem[] = [];
  for (let page = 0; page < 10 && next; page++) {
    const data = await graphGet(next);
    const value = (data && data.value) || [];
    value.forEach(function (item: { name?: string; lastModifiedDateTime?: string; folder?: unknown; webUrl?: string; size?: number }) {
      const name = String(item.name || '');
      if (!name || name.indexOf('~$') === 0) return;
      const isFolder = !!item.folder;
      items.push({
        name,
        kind: isFolder ? 'folder' : 'file',
        lastModified: String(item.lastModifiedDateTime || ''),
        webUrl: String(item.webUrl || ''),
        size: isFolder || typeof item.size !== 'number' ? null : item.size,
      });
    });
    next = data && data['@odata.nextLink'] ? String(data['@odata.nextLink']) : null;
  }
  items.sort(function (a, b) {
    if (a.kind !== b.kind) return a.kind === 'folder' ? -1 : 1;
    return a.name.localeCompare(b.name, undefined, { sensitivity: 'base' });
  });
  return items;
}

export async function listObservationProcess(relative: string[]): Promise<{ path: string; webUrl: string; items: ObservationItem[] }> {
  const extra = (relative || []).map(assertFolderSegment);
  const siteId = await observationSiteId();
  const segments = RS_OBSERVATION_ROOT.concat(extra);
  const folder = await folderOnSite(siteId, segments);
  if (!folder) throw new GraphError(404, 'That Reflective Supervision folder was not found.');
  const items = await listDriveChildren(siteId, segments);
  return { path: '/' + segments.join('/'), webUrl: folder.webUrl, items };
}
