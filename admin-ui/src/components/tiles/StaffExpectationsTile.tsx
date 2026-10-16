import { useEffect, useState, type FormEvent } from 'react';
import { api, type Row } from '../../api';
import { DataTable } from '../DataTable';

type SummaryFile = {
  employee: string;
  fileName: string;
  path: string;
  summary?: {
    openings?: Array<{ date?: string; day?: string; time?: string; status?: string }>;
    counts?: { open?: number };
  };
};

export function StaffExpectationsTile() {
  const [openings, setOpenings] = useState<Row[]>([]);
  const [mileage, setMileage] = useState<Row[]>([]);
  const [cancellations, setCancellations] = useState<Row[]>([]);
  const [compliance, setCompliance] = useState<Row[]>([]);
  const [error, setError] = useState('');
  const [tripDate, setTripDate] = useState('');
  const [miles, setMiles] = useState('');
  const [purpose, setPurpose] = useState('');

  function loadLists() {
    Promise.all([
      api<{ staff: SummaryFile[] }>('/api/scheduling/openings'),
      api<{ mileage: Row[] }>('/api/mileage'),
      api<{ cancellations: Row[] }>('/api/cancellations'),
      api<{ rows: Row[] }>('/api/documentation-compliance'),
    ]).then(([openingData, mileageData, cancellationData, complianceData]) => {
      const flat: Row[] = [];
      (openingData.staff || []).forEach((file) => {
        const slots = file.summary?.openings || [];
        if (!slots.length) {
          flat.push({ id: file.path, employee: file.employee, fileName: file.fileName, date: '', time: '', status: 'No openings' });
          return;
        }
        slots.forEach((slot, index) => {
          flat.push({
            id: file.path + ':' + index,
            employee: file.employee,
            date: slot.date || '',
            day: slot.day || '',
            time: slot.time || '',
            status: slot.status || '',
          });
        });
      });
      setOpenings(flat);
      setMileage(mileageData.mileage || []);
      setCancellations(cancellationData.cancellations || []);
      setCompliance(complianceData.rows || []);
      setError('');
    }).catch((err: Error) => setError(err.message));
  }

  useEffect(() => { loadLists(); }, []);

  async function submitMileage(event: FormEvent) {
    event.preventDefault();
    setError('');
    try {
      await api('/api/mileage', { method: 'POST', body: { tripDate, miles: Number(miles), purpose } });
      setTripDate('');
      setMiles('');
      setPurpose('');
      loadLists();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not save mileage');
    }
  }

  return (
    <section>
      <h1>Staff expectations</h1>
      <p className="lede">
        Each provider has one schedule spreadsheet in SharePoint. It updates when they save their schedule in the app.
      </p>
      {error ? <p className="error">{error}</p> : null}
      <h2>Openings</h2>
      <DataTable
        rows={openings}
        empty="No provider spreadsheets yet. They appear after a provider saves a schedule."
        columns={[
          { key: 'employee', label: 'Staff' },
          { key: 'date', label: 'Date' },
          { key: 'day', label: 'Day' },
          { key: 'time', label: 'Time' },
          { key: 'status', label: 'Status' },
        ]}
      />
      <h2>Mileage</h2>
      <form className="row" onSubmit={submitMileage}>
        <label>Date<input type="date" value={tripDate} onChange={(event) => setTripDate(event.target.value)} required /></label>
        <label>Miles<input value={miles} onChange={(event) => setMiles(event.target.value)} inputMode="decimal" required /></label>
        <label>Purpose<input value={purpose} onChange={(event) => setPurpose(event.target.value)} required /></label>
        <button className="btn" type="submit">Add</button>
      </form>
      <DataTable rows={mileage} empty="No mileage entries." columns={[{ key: 'StaffName', label: 'Staff' }, { key: 'TripDate', label: 'Date' }, { key: 'Miles', label: 'Miles' }, { key: 'Purpose', label: 'Purpose' }, { key: 'Status', label: 'Status' }]} />
      <h2>Cancellations</h2>
      <p className="muted">The Cancellations list is filled from IntakeQ by a scheduled Power Automate flow.</p>
      <DataTable rows={cancellations} empty="No cancellations stored yet." columns={[{ key: 'StaffName', label: 'Staff' }, { key: 'AppointmentDate', label: 'When' }, { key: 'Status', label: 'Status' }, { key: 'ClientName', label: 'Client' }, { key: 'ServiceName', label: 'Service' }]} />
      <h2>Documentation</h2>
      <p className="muted">Weekly counts of open notes, evals, and reports. Note text stays in IntakeQ.</p>
      <DataTable rows={compliance} empty="No weekly documentation summary yet." columns={[{ key: 'StaffName', label: 'Staff' }, { key: 'WeekOf', label: 'Week' }, { key: 'OpenNotes', label: 'Open notes' }, { key: 'OpenEvals', label: 'Open evals' }, { key: 'OpenReports', label: 'Open reports' }]} />
    </section>
  );
}
