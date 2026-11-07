import { useEffect, useState } from 'react';
import { api } from '../../api';

type Item = {
  name: string;
  kind: 'folder' | 'file';
  lastModified: string;
  webUrl: string;
  size: number | null;
};

type FolderResponse = {
  path: string;
  webUrl: string;
  items: Item[];
};

const ROOT = '/reflective-supervision';
const SHAREPOINT_FOLDER =
  'https://emmeco031-my.sharepoint.com/personal/jenriat_emmeco_org/_layouts/15/onedrive.aspx?id=%2Fpersonal%2Fjenriat%5Femmeco%5Forg%2FDocuments%2FEMME%2FLEADERSHIP%20TEAM%2FEMME%20Processes%2FRS%5FObservation%20Process';

function relativeFromHash(hash: string): string[] {
  const path = (hash || '#/').replace(/^#/, '');
  if (path !== ROOT && !path.startsWith(ROOT + '/')) return [];
  return path.slice(ROOT.length).split('/').filter(Boolean).map(function (part) {
    try { return decodeURIComponent(part); } catch (_) { return part; }
  });
}

function folderHref(parts: string[]): string {
  if (!parts.length) return '#' + ROOT;
  return '#' + ROOT + '/' + parts.map(encodeURIComponent).join('/');
}

function when(value: string): string {
  if (!value) return '';
  const dt = new Date(value);
  if (Number.isNaN(dt.getTime())) return value;
  return dt.toLocaleString('en-US', { timeZone: 'America/Denver', dateStyle: 'medium', timeStyle: 'short' });
}

function sizeLabel(bytes: number | null): string {
  if (bytes == null) return '';
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(bytes < 10 * 1024 ? 1 : 0) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

export function ReflectiveSupervisionTile() {
  const [hash, setHash] = useState(function () { return window.location.hash || '#/'; });
  const [items, setItems] = useState<Item[]>([]);
  const [webUrl, setWebUrl] = useState('');
  const [folderPath, setFolderPath] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(true);
  const relative = relativeFromHash(hash);
  const openUrl = webUrl || (relative.length === 0 ? SHAREPOINT_FOLDER : '');

  useEffect(function () {
    const onChange = function () { setHash(window.location.hash || '#/'); };
    window.addEventListener('hashchange', onChange);
    return function () { window.removeEventListener('hashchange', onChange); };
  }, []);

  useEffect(function () {
    let cancelled = false;
    const parts = relativeFromHash(hash);
    const query = parts.length ? '?path=' + encodeURIComponent(parts.join('/')) : '';
    setLoading(true);
    setError('');
    api<FolderResponse>('/api/reflective-supervision' + query)
      .then(function (data) {
        if (cancelled) return;
        setItems(data.items || []);
        setWebUrl(data.webUrl || '');
        setFolderPath(data.path || '');
      })
      .catch(function (err: Error) {
        if (cancelled) return;
        setItems([]);
        setWebUrl('');
        setFolderPath('');
        setError(err.message);
      })
      .finally(function () {
        if (!cancelled) setLoading(false);
      });
    return function () { cancelled = true; };
  }, [hash]);

  return (
    <section>
      <h1>Reflective Supervision</h1>
      <p className="lede">Observation process files from SharePoint. Folders open here. Files open in SharePoint.</p>
      <p className="row">
        <a href={'#' + ROOT}>Reflective Supervision</a>
        {relative.map(function (part, index) {
          const parts = relative.slice(0, index + 1);
          return (
            <span key={parts.join('/')}>
              {' / '}
              <a href={folderHref(parts)}>{part}</a>
            </span>
          );
        })}
      </p>
      {folderPath ? <p className="muted">{folderPath}</p> : null}
      {openUrl ? <p><a className="btn btn--ghost" href={openUrl} target="_blank" rel="noopener noreferrer">Open this folder in SharePoint</a></p> : null}
      {error ? <p className="error">{error}</p> : null}
      {loading ? <p className="muted">Loading the SharePoint folder…</p> : null}
      {!loading && !error && items.length === 0 ? <p className="empty">This folder is empty.</p> : null}
      {items.length > 0 ? (
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Name</th>
                <th>Type</th>
                <th>Updated</th>
                <th>Size</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {items.map(function (item) {
                const next = relative.concat([item.name]);
                return (
                  <tr key={item.kind + ':' + item.name}>
                    <td>
                      {item.kind === 'folder'
                        ? <a href={folderHref(next)}>{item.name}</a>
                        : item.webUrl
                          ? <a href={item.webUrl} target="_blank" rel="noopener noreferrer">{item.name}</a>
                          : item.name}
                    </td>
                    <td>{item.kind === 'folder' ? 'Folder' : 'File'}</td>
                    <td>{when(item.lastModified)}</td>
                    <td>{sizeLabel(item.size)}</td>
                    <td>
                      {item.webUrl
                        ? <a href={item.webUrl} target="_blank" rel="noopener noreferrer">Open</a>
                        : null}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      ) : null}
    </section>
  );
}
