import { useEffect, useState } from 'react';
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
  const [mileageNote, setMileageNote] = useState('');

  function loadLists() {
    Promise.all([
      api<{ staff: SummaryFile[] }>('/api/scheduling/openings'),
      api<{ mileage: Row[]; syncError?: string }>('/api/mileage'),
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
      setMileageNote(mileageData.syncError || '');
      setCancellations(cancellationData.cancellations || []);
      setCompliance(complianceData.rows || []);
      setError('');
    }).catch((err: Error) => setError(err.message));
  }

  useEffect(() => { loadLists(); }, []);

  async function confirmMileage(row: Row) {
    setError('');
    try {
      await api('/api/mileage', {
        method: 'POST',
        body: { action: 'confirm', staffEmail: row.staffEmail, appointmentId: row.appointmentId },
      });
      loadLists();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not confirm mileage');
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
      <p className="muted">Each provider has a mileage spreadsheet in SharePoint. In-person sessions from IntakeQ are added with that client’s average miles. Confirm a row when it is ready for billing.</p>
      {mileageNote ? <p className="error">{mileageNote}</p> : null}
      {mileage.length === 0 ? <p className="empty">No mileage yet. It appears after a provider marks a client In person and IntakeQ shows a session.</p> : (
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Staff</th><th>Date</th><th>Client</th><th>Service</th><th>Miles</th><th>Amount</th><th>Billing</th><th></th>
              </tr>
            </thead>
            <tbody>
              {mileage.map((row) => (
                <tr key={String(row.id || row.appointmentId)}>
                  <td>{String(row.staffName || '')}</td>
                  <td>{String(row.date || '')}</td>
                  <td>{String(row.client || '')}</td>
                  <td>{String(row.service || '')}</td>
                  <td>{String(row.miles ?? '')}</td>
                  <td>{typeof row.amount === 'number' ? '$' + row.amount.toFixed(2) : String(row.amount || '')}</td>
                  <td>{String(row.billing || '')}</td>
                  <td>
                    {row.billing === 'Confirmed' ? null : (
                      <button type="button" className="btn btn--ghost" onClick={() => confirmMileage(row)}>Confirm</button>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      <h2>Cancellations</h2>
      <p className="muted">The Cancellations list is filled from IntakeQ by a scheduled Power Automate flow.</p>
      <DataTable rows={cancellations} empty="No cancellations stored yet." columns={[{ key: 'StaffName', label: 'Staff' }, { key: 'AppointmentDate', label: 'When' }, { key: 'Status', label: 'Status' }, { key: 'ClientName', label: 'Client' }, { key: 'ServiceName', label: 'Service' }]} />
      <h2>Documentation</h2>
      <p className="muted">Weekly counts of open notes, evals, and reports. Note text stays in IntakeQ.</p>
      <DataTable rows={compliance} empty="No weekly documentation summary yet." columns={[{ key: 'StaffName', label: 'Staff' }, { key: 'WeekOf', label: 'Week' }, { key: 'OpenNotes', label: 'Open notes' }, { key: 'OpenEvals', label: 'Open evals' }, { key: 'OpenReports', label: 'Open reports' }]} />
    </section>
  );
}
