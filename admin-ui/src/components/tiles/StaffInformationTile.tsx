import { useEffect, useState, type FormEvent } from 'react';
import { api, type Row } from '../../api';
import { StaffMessageAdmin } from '../StaffMessageAdmin';
import { DataTable } from '../DataTable';

export function StaffInformationTile() {
  const [catalog, setCatalog] = useState<Row[]>([]);
  const [staffId, setStaffId] = useState('');
  const [trainings, setTrainings] = useState<Row[]>([]);
  const [notes, setNotes] = useState<Row[]>([]);
  const [evals, setEvals] = useState<Row[]>([]);
  const [error, setError] = useState('');

  useEffect(() => {
    api<{ trainings: Row[] }>('/api/trainings')
      .then((data) => setCatalog(data.trainings || []))
      .catch((err: Error) => setError(err.message));
  }, []);

  async function lookup(event: FormEvent) {
    event.preventDefault();
    const id = staffId.trim().toLowerCase();
    if (!id) return;
    setError('');
    const path = '/api/staff/' + encodeURIComponent(id);
    try {
      const [trainingData, noteData, evalData] = await Promise.all([
        api<{ trainings: Row[] }>(path + '/trainings'),
        api<{ notes: Row[] }>(path + '/open-notes'),
        api<{ notes: Row[] }>(path + '/open-evals'),
      ]);
      setTrainings(trainingData.trainings || []);
      setNotes(noteData.notes || []);
      setEvals(evalData.notes || []);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not load staff');
    }
  }

  return (
    <section>
      <h1>Staff information</h1>
      <p className="lede">Training records live in SharePoint. Open notes and evals are read live from IntakeQ.</p>
      <h2>Training catalog</h2>
      {error ? <p className="error">{error}</p> : null}
      <DataTable
        rows={catalog}
        empty="No trainings in the Trainings list."
        columns={[
          { key: 'Title', label: 'Training' },
          { key: 'Required', label: 'Required' },
          { key: 'DueDate', label: 'Due' },
          { key: 'Description', label: 'Description' },
        ]}
      />
      <form className="row" onSubmit={lookup}>
        <label>Staff email<input value={staffId} onChange={(event) => setStaffId(event.target.value)} required /></label>
        <button className="btn" type="submit">Look up</button>
      </form>
      <h2>Training status</h2>
      <DataTable rows={trainings} empty="No training status for this person." columns={[{ key: 'TrainingName', label: 'Training' }, { key: 'Status', label: 'Status' }, { key: 'CompletedOn', label: 'Completed' }]} />
      <h2>Open notes</h2>
      <DataTable rows={notes} empty="No open notes." columns={[{ key: 'noteName', label: 'Note' }, { key: 'clientName', label: 'Client' }, { key: 'date', label: 'Date' }, { key: 'status', label: 'Status' }]} />
      <h2>Open evals</h2>
      <DataTable rows={evals} empty="No open evals." columns={[{ key: 'noteName', label: 'Eval' }, { key: 'clientName', label: 'Client' }, { key: 'date', label: 'Date' }, { key: 'status', label: 'Status' }]} />
      <StaffMessageAdmin />
    </section>
  );
}
