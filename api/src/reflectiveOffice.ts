import { GraphError, graphGet, graphSend } from './graphClient';
import { copyFileName, matchingFolder } from './reflectiveMap';

/**
 * SharePoint snapshot and Outlook for reflective supervision.
 * Calendar events and emails say that a meeting exists. They do not include what was said.
 */

const SHARE_URL = 'https://emmeco031.sharepoint.com/:f:/s/Emmeteam/IgC76zwHwa13Rp6tgPf4ysldAf8v9GKA9LPp1ML4qvUVUl8?e=4J5wWy';

type ShareFolder = { driveId: string; folderId: string; name: string };

let shareCache: { folder: ShareFolder; expires: number } | null = null;

export function resetReflectiveOfficeCache(): void {
  shareCache = null;
}

function shareId(url: string): string {
  return 'u!' + Buffer.from(url, 'utf8').toString('base64').replace(/=+$/g, '').replace(/\//g, '_').replace(/\+/g, '-');
}

function graphMessage(err: unknown): string {
  if (err instanceof GraphError) return err.message;
  return 'Microsoft 365 did not accept that request.';
}

async function rsProvidersFolder(): Promise<ShareFolder> {
  if (shareCache && shareCache.expires > Date.now()) return shareCache.folder;
  const item = await graphGet('/shares/' + shareId(SHARE_URL) + '/driveItem?$select=id,name,parentReference');
  const driveId = String(item && item.parentReference && item.parentReference.driveId || '');
  const folderId = String(item && item.id || '');
  if (!driveId || !folderId) throw new GraphError(502, 'The RS Providers folder could not be opened.');
  const folder = { driveId, folderId, name: String(item.name || 'RS Providers') };
  shareCache = { folder, expires: Date.now() + 5 * 60 * 1000 };
  return folder;
}

export type CopyResult = { ok: boolean; path: string; webUrl: string; error: string };

/** Replaces the provider's read-only PDF in the RS Providers folder. */
export async function saveReadOnlyCopy(providerName: string, email: string, pdf: Buffer): Promise<CopyResult> {
  try {
    const root = await rsProvidersFolder();
    const children = await graphGet('/drives/' + root.driveId + '/items/' + root.folderId + '/children?$select=id,name,folder&$top=200');
    const folders = ((children && children.value) || [])
      .filter(function (item: { folder?: unknown; id?: string; name?: string }) { return item && item.folder && item.id; })
      .map(function (item: { id: string; name?: string }) { return { id: String(item.id), name: String(item.name || '') }; });
    const matched = matchingFolder(folders, providerName, email);
    const parentId = matched ? matched.id : root.folderId;
    const parentName = matched ? matched.name : root.name;
    const fileName = copyFileName(providerName);
    const saved = await graphSend('/drives/' + root.driveId + '/items/' + parentId + ':/' + encodeURIComponent(fileName) + ':/content', {
      method: 'PUT',
      rawBody: pdf,
      contentType: 'application/pdf',
    });
    return {
      ok: true,
      path: parentName + '/' + fileName,
      webUrl: String(saved && saved.webUrl || ''),
      error: '',
    };
  } catch (err) {
    return { ok: false, path: '', webUrl: '', error: graphMessage(err) };
  }
}

export async function addCalendarEvent(email: string, start: string, end: string): Promise<string> {
  const data = await graphSend('/users/' + encodeURIComponent(email) + '/events', {
    method: 'POST',
    jsonBody: {
      subject: 'Reflective supervision',
      body: {
        contentType: 'Text',
        content: 'Confirmed in the Emme dashboard. What was said in the session is not written on this calendar event.',
      },
      start: { dateTime: start + ':00', timeZone: 'America/Denver' },
      end: { dateTime: end + ':00', timeZone: 'America/Denver' },
      isReminderOn: true,
      reminderMinutesBeforeStart: 60,
    },
  });
  return String(data && data.id || '');
}

export async function notifyPerson(fromEmail: string, toEmail: string, subject: string, text: string): Promise<string> {
  if (!fromEmail || !toEmail || fromEmail === toEmail) return '';
  try {
    await graphSend('/users/' + encodeURIComponent(fromEmail) + '/sendMail', {
      method: 'POST',
      jsonBody: {
        message: {
          subject,
          body: { contentType: 'Text', content: text },
          toRecipients: [{ emailAddress: { address: toEmail } }],
        },
        saveToSentItems: false,
      },
    });
    return '';
  } catch (err) {
    return graphMessage(err);
  }
}

export type OutlookBlock = { start: string; end: string; subject: string };

/** The provider's Outlook events, so a reflective supervision time is not booked on top of a client visit. */
export async function outlookBlocks(email: string, startDate: string, endDate: string): Promise<{ events: OutlookBlock[]; error: string }> {
  try {
    const path = '/users/' + encodeURIComponent(email) + '/calendarView?startDateTime=' + encodeURIComponent(startDate + 'T00:00:00')
      + '&endDateTime=' + encodeURIComponent(endDate + 'T23:59:59')
      + '&$select=subject,start,end,showAs&$top=40';
    const data = await graphGet(path, { Prefer: 'outlook.timezone="America/Denver"' });
    const events: OutlookBlock[] = [];
    ((data && data.value) || []).forEach(function (item: { subject?: string; start?: { dateTime?: string }; end?: { dateTime?: string }; showAs?: string }) {
      const show = String(item.showAs || '').toLowerCase();
      if (show === 'free') return;
      events.push({
        start: String(item.start && item.start.dateTime || ''),
        end: String(item.end && item.end.dateTime || ''),
        subject: String(item.subject || 'Busy').slice(0, 120),
      });
    });
    return { events, error: '' };
  } catch (err) {
    return { events: [], error: graphMessage(err) };
  }
}
