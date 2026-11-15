import { useState } from 'react';
import { api } from '../../api';
import type { Option, StoredFile } from './types';

function toBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || '').split(',')[1] || '');
    reader.onerror = () => reject(new Error('The file could not be read'));
    reader.readAsDataURL(file);
  });
}

function sizeText(bytes: number): string {
  return bytes > 1024 * 1024 ? (bytes / 1024 / 1024).toFixed(1) + ' MB' : Math.max(1, Math.round(bytes / 1024)) + ' KB';
}

export function FilePanel(props: {
  scope: string;
  files: StoredFile[];
  fileKinds: Option[];
  maxFileBytes: number;
  fileTypes: string[];
  /** The topic's address, for example /api/intake/onboarding. */
  apiBase: string;
  /** Saves answers first so a group entry (for example a funding source) exists on the server before its files are uploaded. */
  ensureSaved: () => Promise<boolean>;
  onChanged: () => void;
}) {
  const mine = props.files.filter((file) => file.scope === props.scope);
  const [kind, setKind] = useState('');
  const [note, setNote] = useState('');
  const [confirmed, setConfirmed] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [fileKey, setFileKey] = useState(0);
  const [chosen, setChosen] = useState<File | null>(null);

  const kindLabel = (id: string) => props.fileKinds.find((item) => item.id === id)?.label || id;

  async function upload() {
    setError('');
    if (!chosen) return setError('Choose a file first.');
    if (!kind) return setError('Choose what kind of file this is.');
    if (!confirmed) return setError('Confirm that the file has no real client information.');
    if (chosen.size > props.maxFileBytes) return setError('That file is larger than ' + Math.round(props.maxFileBytes / 1024 / 1024) + ' megabytes.');
    setBusy(true);
    try {
      if (!(await props.ensureSaved())) throw new Error('Your answers could not be saved yet, so the file was not uploaded. Try again in a moment.');
      await api(props.apiBase + '/files', {
        method: 'POST',
        body: { scope: props.scope, kind, note, name: chosen.name, contentBase64: await toBase64(chosen), confirmedNoClientInfo: true },
      });
      setChosen(null);
      setNote('');
      setConfirmed(false);
      setFileKey((key) => key + 1);
      props.onChanged();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Upload failed');
    } finally {
      setBusy(false);
    }
  }

  async function remove(file: StoredFile) {
    if (!window.confirm('Remove "' + file.name + '"?')) return;
    setError('');
    try {
      await api(props.apiBase + '/files/' + encodeURIComponent(file.id), { method: 'DELETE' });
      props.onChanged();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not remove the file');
    }
  }

  return (
    <div className="setup-files">
      <h4>Templates and sample files</h4>
      <p className="muted">
        Upload blank templates and samples that show how things work. <strong>Before uploading, replace every real client or staff name, birth date, member number, and diagnosis with fake ones, or black them out.</strong>
      </p>
      {mine.length ? (
        <ul className="setup-files__list">
          {mine.map((file) => (
            <li key={file.id}>
              <a href={props.apiBase + '/files/' + encodeURIComponent(file.id)}>{file.name}</a>
              <span className="muted"> · {kindLabel(file.kind)} · {sizeText(file.size)} · added by {file.uploadedBy}</span>
              {file.note ? <div className="muted">{file.note}</div> : null}
              <button type="button" className="btn btn--ghost" onClick={() => remove(file)}>Remove</button>
            </li>
          ))}
        </ul>
      ) : <p className="empty">No files yet.</p>}
      <div className="row">
        <label>
          What is this file?
          <select value={kind} onChange={(event) => setKind(event.target.value)}>
            <option value="">Choose…</option>
            {props.fileKinds.map((item) => <option key={item.id} value={item.id}>{item.label}</option>)}
          </select>
        </label>
        <label>
          Short note (optional)
          <input value={note} maxLength={300} onChange={(event) => setNote(event.target.value)} placeholder="What should we know about it?" />
        </label>
        <label>
          File
          <input
            key={fileKey}
            type="file"
            accept={props.fileTypes.map((type) => '.' + type).join(',')}
            onChange={(event) => setChosen(event.target.files && event.target.files[0] ? event.target.files[0] : null)}
          />
        </label>
      </div>
      <label className="setup-check">
        <input type="checkbox" checked={confirmed} onChange={(event) => setConfirmed(event.target.checked)} />
        <span>This file has no real client information in it.</span>
      </label>
      <div className="row">
        <button type="button" className="btn" disabled={busy} onClick={upload}>{busy ? 'Uploading…' : 'Upload file'}</button>
        <span className="muted">Allowed: {props.fileTypes.join(', ')}. Up to {Math.round(props.maxFileBytes / 1024 / 1024)} megabytes each.</span>
      </div>
      {error ? <p className="error">{error}</p> : null}
    </div>
  );
}
