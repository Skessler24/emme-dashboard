import { useEffect, useState, type FormEvent } from 'react';
import { api, type Row } from '../../api';
import { DataTable } from '../DataTable';

export function ReferralsTile() {
  const [rows, setRows] = useState<Row[]>([]);
  const [error, setError] = useState('');
  const [referrerName, setReferrerName] = useState('');
  const [assignedTo, setAssignedTo] = useState('');
  const [notes, setNotes] = useState('');

  function load() {
    api<{ referrals: Row[] }>('/api/referrals')
      .then((data) => { setRows(data.referrals || []); setError(''); })
      .catch((err: Error) => setError(err.message));
  }

  useEffect(() => { load(); }, []);

  async function submit(event: FormEvent) {
    event.preventDefault();
    setError('');
    try {
      await api('/api/referrals', {
        method: 'POST',
        body: { referrerName, assignedTo, notes, status: 'New' },
      });
      setReferrerName('');
      setAssignedTo('');
      setNotes('');
      load();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not save referral');
    }
  }

  return (
    <section>
      <h1>Referrals</h1>
      <p className="lede">Stored on the Referrals SharePoint list.</p>
      <form className="panel" onSubmit={submit}>
        <div className="row">
          <label>Referrer<input value={referrerName} onChange={(event) => setReferrerName(event.target.value)} required /></label>
          <label>Assigned to<input value={assignedTo} onChange={(event) => setAssignedTo(event.target.value)} /></label>
        </div>
        <label>Notes<textarea value={notes} onChange={(event) => setNotes(event.target.value)} /></label>
        <div className="row"><button className="btn" type="submit">Add referral</button></div>
      </form>
      {error ? <p className="error">{error}</p> : null}
      <DataTable
        rows={rows}
        empty="No referrals yet."
        columns={[
          { key: 'ReferrerName', label: 'Referrer' },
          { key: 'Status', label: 'Status' },
          { key: 'AssignedTo', label: 'Assigned' },
          { key: 'ReceivedOn', label: 'Received' },
          { key: 'Notes', label: 'Notes' },
        ]}
      />
    </section>
  );
}
