import { useEffect, useState } from 'react';
import { api, type Row } from '../../api';
import { DataTable } from '../DataTable';
import { TileWithSetup } from '../setup/TileSetup';
import { ExpectationsBoard } from './ExpectationsBoard';

type SummaryFile = {
  employee: string;
  fileName: string;
  path: string;
  summary?: {
    openings?: Array<{ date?: string; day?: string; time?: string; status?: string }>;
    counts?: { open?: number };
  };
};

type WeekCounts = { week: string; notes: number; evals: number; reports: number };

function asCount(value: unknown): number {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function countLabel(count: number, singular: string, plural: string): string {
  return count + ' ' + (count === 1 ? singular : plural);
}

function latestWeek(rows: Row[]): WeekCounts | null {
  let week = '';
  rows.forEach((row) => {
    const value = String(row.WeekOf || '');
    if (value > week) week = value;
  });
  if (!week) return null;
  const inWeek = rows.filter((row) => String(row.WeekOf || '') === week);
  return {
    week,
    notes: inWeek.reduce((sum, row) => sum + asCount(row.OpenNotes), 0),
    evals: inWeek.reduce((sum, row) => sum + asCount(row.OpenEvals), 0),
    reports: inWeek.reduce((sum, row) => sum + asCount(row.OpenReports), 0),
  };
}

function StaffExpectationsOverview() {
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

  const week = latestWeek(compliance);

  return (
    <section>
      <h1>Staff expectations</h1>
      <p className="lede">
        Openings come straight from each provider’s schedule in the app, so they are current the moment a provider saves.
        The mileage address check, trip table, open notes, and cancellation review are first. The lists underneath are the workbook, the cancellations list, and the weekly counts.
      </p>
      {error ? <p className="error">{error}</p> : null}
      <ExpectationsBoard />
      <h2>Openings</h2>
      <DataTable
        rows={openings}
        empty="No provider schedules yet. They appear after a provider saves a schedule."
        columns={[
          { key: 'employee', label: 'Staff' },
          { key: 'date', label: 'Date' },
          { key: 'day', label: 'Day' },
          { key: 'time', label: 'Time' },
          { key: 'status', label: 'Status' },
        ]}
      />
      <h2>Mileage</h2>
      <p className="muted">Each provider’s mileage form is in SharePoint under Payroll & Mileage. A nightly check adds a row when the client is marked In person, the visit is on the provider’s Outlook calendar, and IntakeQ shows the appointment happened. Confirm a row when it is ready for billing. Travel outside 5 straight-line miles is approved or denied in the section above.</p>
      {mileageNote ? <p className="error">{mileageNote}</p> : null}
      {mileage.length === 0 ? <p className="empty">No mileage yet. It appears after a provider marks a client In person, enters average miles, and both Outlook and IntakeQ show the visit.</p> : (
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
      <p className="muted">These rows are kept for the quarterly productivity review. A scheduled check copies cancellations from IntakeQ into the Cancellations list. The review above applies the attendance rules. This table is every stored cancellation.</p>
      <DataTable rows={cancellations} empty="No cancellations stored yet." columns={[{ key: 'StaffName', label: 'Staff' }, { key: 'AppointmentDate', label: 'When' }, { key: 'Status', label: 'Status' }, { key: 'ClientName', label: 'Client' }, { key: 'ServiceName', label: 'Service' }]} />
      <h2>Documentation</h2>
      <p className="muted">Notes are expected to be locked within 48 hours of the session. Evaluations and reports follow the end of the week. The table below is the weekly counts. Note text stays in IntakeQ. The open list above is the live check.</p>
      <div className="panel" role="note" aria-label="Evaluation and report counts">
        {week ? (
          <p>Week of {week.week}: {countLabel(week.notes, 'open note', 'open notes')}, {countLabel(week.evals, 'open evaluation', 'open evaluations')}, {countLabel(week.reports, 'open report', 'open reports')}.</p>
        ) : (
          <p>Open notes, open evaluations, and open reports appear here after the weekly count is saved.</p>
        )}
        <p className="muted">An open evaluation or report that is still open after Friday of its week is past the end of the week. Finished ones are recorded above for the UNICORN portal or the Medicaid PAR fax.</p>
      </div>
      <DataTable rows={compliance} empty="No weekly documentation summary yet." columns={[{ key: 'StaffName', label: 'Staff' }, { key: 'WeekOf', label: 'Week' }, { key: 'OpenNotes', label: 'Open notes' }, { key: 'OpenEvals', label: 'Open evals' }, { key: 'OpenReports', label: 'Open reports' }]} />
      <h2>Reminders</h2>
      <p className="muted">The email check, open notes, travel outside 5 miles, and too many cancellations show on this page. A Power Automate flow sends the matching email at 6:00 AM Denver time. This website does not run its own timer.</p>
    </section>
  );
}

export function StaffExpectationsTile() {
  return (
    <TileWithSetup hashId="expectations" tileId="staff-expectations" label="Staff Expectations">
      <StaffExpectationsOverview />
    </TileWithSetup>
  );
}
