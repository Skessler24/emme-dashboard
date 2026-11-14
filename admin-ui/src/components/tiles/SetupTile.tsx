import { useCallback, useEffect, useState } from 'react';
import { api } from '../../api';

type Status = {
  graphConfigured: boolean;
  sharePointSettings: { hostname: string; sitePath: string };
  site: { name: string; webUrl: string } | null;
  siteError: string;
  lists: Array<{ name: string; exists: boolean; missingColumns: string[] }>;
  folders: Array<{ label: string; path: string; usedFor: string; onSharedSite: boolean; onPersonalOneDrive: boolean }>;
  usingPersonalOneDrive: boolean;
  nextSteps: string[];
  ready: boolean;
};

type CreateResult = {
  created: { createdLists: string[]; addedColumns: Array<{ list: string; columns: string[] }> };
  status: Status;
};

export function SetupTile() {
  const [status, setStatus] = useState<Status | null>(null);
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const [note, setNote] = useState('');

  const load = useCallback(() => {
    setError('');
    return api<Status>('/api/setup')
      .then(setStatus)
      .catch((err: Error) => setError(err.message));
  }, []);

  useEffect(() => { void load(); }, [load]);

  async function createLists() {
    setBusy(true);
    setNote('');
    setError('');
    try {
      const result = await api<CreateResult>('/api/setup', { method: 'POST', body: {} });
      setStatus(result.status);
      const made = result.created.createdLists.length;
      const added = result.created.addedColumns.reduce((sum, item) => sum + item.columns.length, 0);
      setNote(made || added ? 'Created ' + made + ' list' + (made === 1 ? '' : 's') + (added ? ' and added ' + added + ' missing column' + (added === 1 ? '' : 's') : '') + '.' : 'Everything was already in place.');
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not create the lists');
    } finally {
      setBusy(false);
    }
  }

  const listsReady = !!status && status.lists.length > 0 && status.lists.every((l) => l.exists && l.missingColumns.length === 0);

  return (
    <section>
      <h1>Setup</h1>
      <p className="lede">Checks that the Emme SharePoint site, lists, and folders are ready. Nothing here changes until you press a button.</p>
      {error ? <p className="error">{error}</p> : null}
      {note ? <p className="lede">{note}</p> : null}
      {!status && !error ? <p className="lede">Checking…</p> : null}
      {status ? (
        <>
          <p><strong>{status.ready ? 'All set.' : 'Not finished yet.'}</strong></p>
          {status.nextSteps.length ? (
            <ol>
              {status.nextSteps.map((step) => <li key={step}>{step}</li>)}
            </ol>
          ) : null}

          <h2>SharePoint site</h2>
          {status.site ? (
            <p>{status.site.name} — <a href={status.site.webUrl} target="_blank" rel="noreferrer">{status.site.webUrl}</a></p>
          ) : (
            <p className="error">{status.siteError || 'Not connected.'}</p>
          )}
          <p className="lede">
            Azure settings: hostname {status.sharePointSettings.hostname || '(not set)'}, site path {status.sharePointSettings.sitePath || '(not set)'}.
          </p>

          <h2>Folders</h2>
          <table>
            <thead><tr><th>Folder</th><th>Used for</th><th>Emme site</th><th>Personal OneDrive</th></tr></thead>
            <tbody>
              {status.folders.map((folder) => (
                <tr key={folder.path}>
                  <td>{folder.path}</td>
                  <td>{folder.usedFor}</td>
                  <td>{folder.onSharedSite ? 'Found' : 'Not found'}</td>
                  <td>{folder.onPersonalOneDrive ? 'Found (old copy)' : '-'}</td>
                </tr>
              ))}
            </tbody>
          </table>

          <h2>Lists</h2>
          <table>
            <thead><tr><th>List</th><th>Status</th></tr></thead>
            <tbody>
              {status.lists.map((list) => (
                <tr key={list.name}>
                  <td>{list.name}</td>
                  <td>
                    {!list.exists ? 'Missing' : list.missingColumns.length ? 'Missing columns: ' + list.missingColumns.join(', ') : 'Ready'}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <div className="row">
            <button type="button" onClick={() => void createLists()} disabled={busy || !status.site || listsReady}>
              {busy ? 'Creating…' : listsReady ? 'Lists are ready' : 'Create missing lists'}
            </button>
            <button type="button" onClick={() => void load()} disabled={busy}>Check again</button>
          </div>
        </>
      ) : null}
    </section>
  );
}
