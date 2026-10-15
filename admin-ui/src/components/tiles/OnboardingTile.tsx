import { useState, type FormEvent } from 'react';
import { api } from '../../api';
import { DataTable } from '../DataTable';

type Docs = { name: string; lastModified: string };

export function OnboardingTile() {
  const [employeeId, setEmployeeId] = useState('');
  const [employeeName, setEmployeeName] = useState('');
  const [checklist, setChecklist] = useState<Array<Record<string, unknown>>>([]);
  const [currentDocs, setCurrentDocs] = useState<Docs[]>([]);
  const [archivedDocs, setArchivedDocs] = useState<Docs[]>([]);
  const [folder, setFolder] = useState('current');
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');

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
      <p className="lede">Checklist rows come from the OnboardingChecklist list. Documents stay in the employee SharePoint folder.</p>
      <form className="row" onSubmit={load}>
        <label>Employee email<input value={employeeId} onChange={(event) => setEmployeeId(event.target.value)} required /></label>
        <label>Folder name<input value={employeeName} onChange={(event) => setEmployeeName(event.target.value)} placeholder="First Last" /></label>
        <button className="btn" type="submit">Load</button>
      </form>
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
