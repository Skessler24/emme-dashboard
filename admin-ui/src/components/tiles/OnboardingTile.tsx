import { useState, type FormEvent } from 'react';
import { api } from '../../api';
import { DataTable } from '../DataTable';

type Docs = { name: string; lastModified: string };
type FoundFile = { folder: string; fileName: string; path: string; matchedBy: string };

function folderLine(label: string, found: FoundFile | null): string {
  if (!found) return 'No existing ' + label.toLowerCase() + ' folder for this email.';
  if (!found.fileName) return label + ' folder found: ' + found.path + ' (' + found.matchedBy + '). No workbook in that folder yet.';
  return label + ': ' + found.path + ' (' + found.matchedBy + ')';
}

export function OnboardingTile() {
  const [employeeId, setEmployeeId] = useState('');
  const [employeeName, setEmployeeName] = useState('');
  const [checklist, setChecklist] = useState<Array<Record<string, unknown>>>([]);
  const [currentDocs, setCurrentDocs] = useState<Docs[]>([]);
  const [archivedDocs, setArchivedDocs] = useState<Docs[]>([]);
  const [folder, setFolder] = useState('current');
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [folders, setFolders] = useState<{ mileage: FoundFile | null; schedule: FoundFile | null; email: string; name: string } | null>(null);

  async function load(event?: FormEvent) {
    event?.preventDefault();
    setError('');
    setNotice('');
    try {
      const data = await api<{ checklist: Array<Record<string, unknown>>; currentDocuments: Docs[]; archivedDocuments: Docs[] }>(
        '/api/onboarding/' + encodeURIComponent(employeeId.trim()) + '?employeeName=' + encodeURIComponent(employeeName.trim())
      );
      setChecklist(data.checklist || []);
      setCurrentDocs(data.currentDocuments || []);
      setArchivedDocs(data.archivedDocuments || []);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not load onboarding');
    }
  }

  async function findFolders() {
    const email = employeeId.trim().toLowerCase();
    if (!/@emmeco\.org$/.test(email)) {
      setError('Enter the provider’s emmeco.org email before searching.');
      setFolders(null);
      return;
    }
    setError('');
    setNotice('');
    try {
      const data = await api<{ found: { mileage: FoundFile | null; schedule: FoundFile | null; email: string; name: string } }>(
        '/api/onboarding/' + encodeURIComponent(email) + '/folders?employeeName=' + encodeURIComponent(employeeName.trim())
      );
      setFolders(data.found);
      setNotice('Existing folders checked for ' + email + '.');
    } catch (err) {
      setFolders(null);
      setError(err instanceof Error ? err.message : 'Could not search SharePoint');
    }
  }

  async function upload(file: File | undefined) {
    if (!file || !employeeName.trim()) {
      setError('Enter the employee name before uploading.');
      return;
    }
    const contentBase64 = await fileToBase64(file);
    setError('');
    try {
      await api('/api/onboarding/' + encodeURIComponent(employeeId.trim() || employeeName.trim()) + '/documents', {
        method: 'POST',
        body: { employeeName: employeeName.trim(), fileName: file.name, contentBase64, folder },
      });
      setNotice('Uploaded ' + file.name);
      if (employeeId.trim()) await load();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Upload failed');
    }
  }

  return (
    <section>
      <h1>Onboarding</h1>
      <p className="lede">Checklist rows come from the OnboardingChecklist list. Mileage and schedule folders are matched to the provider’s emmeco.org email. Finding them does not create anything new.</p>
      <form className="row" onSubmit={load}>
        <label>Emme email<input value={employeeId} onChange={(event) => setEmployeeId(event.target.value)} placeholder="name@emmeco.org" required /></label>
        <label>Provider name<input value={employeeName} onChange={(event) => setEmployeeName(event.target.value)} placeholder="First Last" /></label>
        <button className="btn" type="submit">Load</button>
        <button className="btn btn--ghost" type="button" onClick={findFolders}>Find existing folders</button>
      </form>
      {folders ? (
        <div>
          <p className="muted">Searched existing admin folders for {folders.email}. Nothing new was created.</p>
          <p>{folderLine('Mileage', folders.mileage)}</p>
          <p>{folderLine('Schedule', folders.schedule)}</p>
        </div>
      ) : null}
      {error ? <p className="error">{error}</p> : null}
      {notice ? <p className="muted">{notice}</p> : null}
      <DataTable
        rows={checklist}
        empty="No checklist items for this employee."
        columns={[
          { key: 'Title', label: 'Item' },
          { key: 'Status', label: 'Status' },
          { key: 'DueDate', label: 'Due' },
          { key: 'EmployeeName', label: 'Employee' },
        ]}
      />
      <h2>Documents</h2>
      <div className="row">
        <label>
          Save to
          <select value={folder} onChange={(event) => setFolder(event.target.value)}>
            <option value="current">Current documents</option>
            <option value="archived">Archived documents</option>
          </select>
        </label>
        <label>
          File
          <input type="file" onChange={(event) => upload(event.target.files?.[0])} />
        </label>
      </div>
      <h2>Current</h2>
      <DataTable rows={currentDocs} empty="No current documents." columns={[{ key: 'name', label: 'File' }, { key: 'lastModified', label: 'Updated' }]} />
      <h2>Archived</h2>
      <DataTable rows={archivedDocs} empty="No archived documents." columns={[{ key: 'name', label: 'File' }, { key: 'lastModified', label: 'Updated' }]} />
    </section>
  );
}

function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error('Could not read the file'));
    reader.onload = () => {
      const value = String(reader.result || '');
      resolve(value.slice(value.indexOf(',') + 1));
    };
    reader.readAsDataURL(file);
  });
}
