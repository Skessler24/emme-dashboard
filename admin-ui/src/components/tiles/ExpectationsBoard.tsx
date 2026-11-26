import { useEffect, useMemo, useState, type FormEvent } from 'react';
import { api } from '../../api';

type Trip = { id: string; staffEmail: string; staffName: string; date: string; startPlace: string; endPlace: string; miles: number };
type RequestRow = { id: string; staffEmail: string; staffName: string; clientAddress: string; miles: number; status: string };
type NoteRow = { id: string; staffEmail: string; staffName: string; noteName: string; date: string; clientName: string; hoursOpen: number | null; pastEndOfWeek?: boolean };
type ClientRow = {
  key: string;
  client: string;
  staffName: string;
  staffEmail: string;
  frequency: string;
  limit: number;
  productivity: number;
  incidents: number;
  needsOutreach: boolean;
  outreach: { choice: string; at: string } | null;
  message: string;
  rows: Array<{ date: string; status: string; kind: string; service: string }>;
};
type FollowUp = { noteId: string; followUp: string };
type Board = {
  role: string;
  me: string;
  day: string;
  month: string;
  home: { address: string; mapLink: string } | null;
  trips: Trip[];
  monthMiles: number;
  monthHours: number;
  ratio: number | null;
  ratioHigh: boolean;
  requests: RequestRow[];
  emailChecked: boolean;
  notes: { late: NoteRow[]; evals: NoteRow[]; reports: NoteRow[] };
  notesError: string;
  clients: ClientRow[];
  cancellationsError: string;
  followUps: FollowUp[];
};

const EMPTY: Board = {
  role: '', me: '', day: '', month: '', home: null, trips: [], monthMiles: 0, monthHours: 0,
  ratio: null, ratioHigh: false, requests: [], emailChecked: false,
  notes: { late: [], evals: [], reports: [] }, notesError: '', clients: [], cancellationsError: '', followUps: [],
};

function followOf(board: Board, id: string): string {
  const found = board.followUps.find((item) => item.noteId === id);
  return found ? found.followUp : '';
}

export function ExpectationsBoard() {
  const [board, setBoard] = useState<Board>(EMPTY);
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [loading, setLoading] = useState(true);
  const [homeAddress, setHomeAddress] = useState('');
  const [mapLink, setMapLink] = useState('');
  const [clientAddress, setClientAddress] = useState('');
  const [zone, setZone] = useState<{ miles: number; outside: boolean } | null>(null);
  const [trip, setTrip] = useState({ date: '', startPlace: '', endPlace: '', miles: '' });
  const [hours, setHours] = useState('');
  const [filter, setFilter] = useState('');
  const [celebrate, setCelebrate] = useState(false);

  function load() {
    setLoading(true);
    api<Board>('/api/mileage?part=expectations')
      .then((data) => {
        setBoard(data);
        setHomeAddress((data.home && data.home.address) || '');
        setMapLink((data.home && data.home.mapLink) || '');
        setHours(data.monthHours ? String(data.monthHours) : '');
        setError('');
      })
      .catch((err: Error) => setError(err.message))
      .finally(() => setLoading(false));
  }

  useEffect(() => { load(); }, []);

  async function post(body: Record<string, unknown>, done: string) {
    setError('');
    setNotice('');
    try {
      await api('/api/mileage', { method: 'POST', body });
      setNotice(done);
      load();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not save');
    }
  }

  async function checkZone(event: FormEvent) {
    event.preventDefault();
    setError('');
    setZone(null);
    try {
      const result = await api<{ miles: number; outside: boolean }>('/api/mileage', {
        method: 'POST',
        body: { action: 'expectations-zone', clientAddress },
      });
      setZone(result);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not measure that address');
    }
  }

  async function toggleEmail(checked: boolean) {
    setCelebrate(checked);
    try {
      await api('/api/mileage', { method: 'POST', body: { action: 'expectations-email', checked } });
      setBoard((current) => ({ ...current, emailChecked: checked }));
    } catch (err) {
      setCelebrate(false);
      setError(err instanceof Error ? err.message : 'Could not save the email check');
    }
  }

  const shownClients = useMemo(() => {
    const needle = filter.trim().toLowerCase();
    if (!needle) return board.clients || [];
    return (board.clients || []).filter((row) => row.client.toLowerCase().indexOf(needle) !== -1 || row.staffName.toLowerCase().indexOf(needle) !== -1);
  }, [board.clients, filter]);

  const admin = board.role === 'admin';

  return (
    <div className="exp-board">
      <div className={'exp-email' + (celebrate ? ' exp-celebrate' : '')}>
        <label className="exp-check">
          <input type="checkbox" checked={!!board.emailChecked} onChange={(event) => { void toggleEmail(event.target.checked); }} />
          <span>Did you check and reply to your emails today?</span>
        </label>
        {board.emailChecked ? <p className="exp-burst">Checked for today.</p> : <p className="muted">This resets each morning. It does not open anyone’s inbox.</p>}
      </div>
      {error ? <p className="error">{error}</p> : null}
      {notice ? <p className="muted">{notice}</p> : null}
      {loading ? <p className="muted">Loading mileage, notes, and cancellations…</p> : null}

      <h2>Can this visit be in person?</h2>
      <p className="muted">Enter the client’s address. The distance is a straight line from the home address saved here to that address. More than 5 miles needs an administrator to approve the travel. Driving miles for the month are the miles you type in the trip table, the same columns as the mileage spreadsheet.</p>
      <form className="row" onSubmit={(event) => { event.preventDefault(); void post({ action: 'expectations-home', homeAddress, mapLink }, 'Home address saved.'); }}>
        <label>Home address
          <input value={homeAddress} onChange={(event) => setHomeAddress(event.target.value)} required />
        </label>
        <label>Map link
          <input value={mapLink} onChange={(event) => setMapLink(event.target.value)} placeholder="https://" />
        </label>
        <button className="btn" type="submit">Save home</button>
        {board.home && board.home.mapLink ? <a className="btn btn--ghost" href={board.home.mapLink} target="_blank" rel="noreferrer">Open map</a> : null}
      </form>
      <form className="row" onSubmit={(event) => { void checkZone(event); }}>
        <label>Client address
          <input value={clientAddress} onChange={(event) => setClientAddress(event.target.value)} required />
        </label>
        <button className="btn btn--ghost" type="submit">Check 5 miles</button>
      </form>
      {zone ? (
        <div className="panel" role="status">
          <p>{zone.miles} straight-line miles. {zone.outside ? 'This address is outside the 5-mile range.' : 'This address is inside the 5-mile range.'}</p>
          {zone.outside ? (
            <button className="btn" type="button" onClick={() => { void post({ action: 'expectations-approval', clientAddress, miles: zone.miles }, 'Sent for administrator approval.'); }}>Send for admin approval</button>
          ) : null}
        </div>
      ) : null}

      <h3>Travel waiting for a decision</h3>
      {(board.requests || []).filter((row) => row.status === 'pending').length === 0 ? <p className="empty">No travel is waiting.</p> : (
        <div className="table-wrap">
          <table>
            <thead><tr><th>Staff</th><th>Address</th><th>Miles</th><th>Status</th>{admin ? <th></th> : null}</tr></thead>
            <tbody>
              {board.requests.filter((row) => row.status === 'pending').map((row) => (
                <tr key={row.id}>
                  <td>{row.staffName}</td>
                  <td>{row.clientAddress}</td>
                  <td>{row.miles}</td>
                  <td>Waiting</td>
                  {admin ? (
                    <td>
                      <button className="btn btn--ghost" type="button" onClick={() => { void post({ action: 'expectations-decision', id: row.id, decision: 'approved' }, 'Travel approved.'); }}>Approve</button>
                      <button className="btn btn--ghost" type="button" onClick={() => { void post({ action: 'expectations-decision', id: row.id, decision: 'denied' }, 'Travel denied.'); }}>Deny</button>
                    </td>
                  ) : null}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <h2>Mileage table</h2>
      <p className="muted">Date, starting location, ending location, and miles. The total and the monthly ratio update when you save. A ratio above 1.8 is highlighted. Save as PDF uses the print window. The read-only SharePoint copy is named Dashboard Mileage Copy and does not replace the original form.</p>
      <form className="row" onSubmit={(event) => { event.preventDefault(); void post({ action: 'expectations-trip', date: trip.date, startPlace: trip.startPlace, endPlace: trip.endPlace, miles: Number(trip.miles) }, 'Trip saved.'); setTrip({ date: '', startPlace: '', endPlace: '', miles: '' }); }}>
        <label>Date<input type="date" value={trip.date} onChange={(event) => setTrip({ ...trip, date: event.target.value })} required /></label>
        <label>Starting location<input value={trip.startPlace} onChange={(event) => setTrip({ ...trip, startPlace: event.target.value })} required /></label>
        <label>Ending location<input value={trip.endPlace} onChange={(event) => setTrip({ ...trip, endPlace: event.target.value })} required /></label>
        <label>Miles<input inputMode="decimal" value={trip.miles} onChange={(event) => setTrip({ ...trip, miles: event.target.value })} required /></label>
        <button className="btn" type="submit">Add trip</button>
      </form>
      <form className="row" onSubmit={(event) => { event.preventDefault(); void post({ action: 'expectations-hours', hours: Number(hours), month: board.month }, 'Hours saved.'); }}>
        <label>Monthly hours of productivity
          <input inputMode="decimal" value={hours} onChange={(event) => setHours(event.target.value)} />
        </label>
        <button className="btn btn--ghost" type="submit">Save hours</button>
        <button className="btn btn--ghost" type="button" onClick={() => window.print()}>Save as PDF</button>
        <button className="btn btn--ghost" type="button" onClick={() => { void post({ action: 'expectations-copy' }, 'Read-only copy saved in Payroll & Mileage.'); }}>Save read-only copy</button>
      </form>
      <p className={board.ratioHigh ? 'exp-ratio exp-ratio--high' : 'exp-ratio'}>
        This month: {board.monthMiles} miles, {board.monthHours} hours, ratio {board.ratio == null ? '—' : board.ratio}.
        {board.ratioHigh ? ' Above 1.8.' : ''}
      </p>
      <div className="table-wrap exp-print">
        <table>
          <thead><tr><th>Staff</th><th>Date</th><th>Starting location</th><th>Ending location</th><th>Miles</th><th></th></tr></thead>
          <tbody>
            {(board.trips || []).map((row) => (
              <tr key={row.id}>
                <td>{row.staffName}</td>
                <td>{row.date}</td>
                <td>{row.startPlace}</td>
                <td>{row.endPlace}</td>
                <td>{row.miles}</td>
                <td><button className="btn btn--ghost" type="button" onClick={() => { void post({ action: 'expectations-trip', id: row.id, remove: true }, 'Trip removed.'); }}>Remove</button></td>
              </tr>
            ))}
          </tbody>
        </table>
        {(board.trips || []).length === 0 ? <p className="empty">No trips saved in the dashboard yet.</p> : null}
      </div>

      <h2>Open notes</h2>
      <p className="muted">A note stays here when IntakeQ still has it open more than 48 hours after the note date. Administrators see every staff member. A provider sees only their own. Refresh checks IntakeQ now. The 6:00 AM check is a Power Automate flow.</p>
      <button className="btn btn--ghost" type="button" onClick={() => load()}>Refresh from IntakeQ</button>
      {board.notesError ? <p className="error">{board.notesError}</p> : null}
      <NoteTable rows={board.notes.late || []} empty="No notes are open past 48 hours." />

      <h2>Evaluations and reports</h2>
      <p className="muted">These follow the end of the week. Saturday means the Friday of that note’s week has passed. Everyone can see the ones that apply to them. When an evaluation or report is finished, an administrator records that it needs the UNICORN portal for early intervention, or a PAR fax to the doctor for Medicaid.</p>
      <h3>Open evaluations</h3>
      <DocTable rows={board.notes.evals || []} admin={admin} followOf={(id) => followOf(board, id)} onFollow={(row, followUp) => { void post({ action: 'expectations-followup', noteId: row.id, followUp, kind: 'eval', staffEmail: row.staffEmail, noteName: row.noteName, date: row.date }, 'Follow-up saved.'); }} />
      <h3>Open reports</h3>
      <DocTable rows={board.notes.reports || []} admin={admin} followOf={(id) => followOf(board, id)} onFollow={(row, followUp) => { void post({ action: 'expectations-followup', noteId: row.id, followUp, kind: 'report', staffEmail: row.staffEmail, noteName: row.noteName, date: row.date }, 'Follow-up saved.'); }} />

      <h2>Cancellations and productivity</h2>
      <p className="muted">A cancellation counts for productivity when it was not rescheduled, whatever the reason. A no-show or a late cancellation (less than 24 hours) is an attendance incident. In a rolling six months the limit is 4 for weekly visits, 2 for twice a month, and 1 for monthly. Advance notice of 48 hours or more is not an incident. When the limit is reached, the family message is ready and an administrator can send it or call.</p>
      {board.cancellationsError ? <p className="error">{board.cancellationsError}</p> : null}
      <label>Filter by client or staff
        <input value={filter} onChange={(event) => setFilter(event.target.value)} />
      </label>
      {shownClients.length === 0 ? <p className="empty">No matching cancellations in the last six months.</p> : shownClients.map((row) => (
        <article className="exp-client" key={row.key}>
          <h3>{row.client || 'Unknown client'}</h3>
          <p>{row.staffName}: {row.productivity} cancellations without a reschedule, {row.incidents} of {row.limit} attendance incidents ({row.frequency}).</p>
          {row.needsOutreach ? <p className="exp-ratio--high">{row.message}</p> : null}
          {admin ? (
            <div className="row">
              <label>Visit frequency
                <select value={row.frequency} onChange={(event) => { void post({ action: 'expectations-frequency', key: row.key, frequency: event.target.value }, 'Frequency saved.'); }}>
                  <option value="weekly">Weekly</option>
                  <option value="twice-monthly">Twice a month</option>
                  <option value="monthly">Monthly</option>
                </select>
              </label>
              {row.needsOutreach ? (
                <>
                  <button className="btn" type="button" onClick={() => { void post({ action: 'expectations-outreach', key: row.key, choice: 'email' }, 'Marked to send the email.'); }}>Send admin email</button>
                  <button className="btn btn--ghost" type="button" onClick={() => { void post({ action: 'expectations-outreach', key: row.key, choice: 'call' }, 'Marked for a phone call.'); }}>Call in person</button>
                </>
              ) : null}
            </div>
          ) : null}
          {row.outreach ? <p className="muted">Recorded: {row.outreach.choice === 'call' ? 'phone call' : 'email'}.</p> : null}
          <div className="table-wrap">
            <table>
              <thead><tr><th>Date</th><th>Status</th><th>Kind</th><th>Service</th></tr></thead>
              <tbody>
                {row.rows.map((visit, index) => (
                  <tr key={row.key + index}><td>{visit.date}</td><td>{visit.status}</td><td>{visit.kind}</td><td>{visit.service}</td></tr>
                ))}
              </tbody>
            </table>
          </div>
        </article>
      ))}
    </div>
  );
}

function NoteTable(props: { rows: NoteRow[]; empty: string }) {
  if (!props.rows.length) return <p className="empty">{props.empty}</p>;
  return (
    <div className="table-wrap">
      <table>
        <thead><tr><th>Staff</th><th>Note</th><th>Client</th><th>Date</th><th>Hours open</th></tr></thead>
        <tbody>
          {props.rows.map((row) => (
            <tr key={row.id || row.staffEmail + row.date}>
              <td>{row.staffName}</td><td>{row.noteName}</td><td>{row.clientName}</td><td>{row.date}</td><td>{row.hoursOpen ?? ''}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function DocTable(props: { rows: NoteRow[]; admin: boolean; followOf: (id: string) => string; onFollow: (row: NoteRow, followUp: string) => void }) {
  if (!props.rows.length) return <p className="empty">None open.</p>;
  return (
    <div className="table-wrap">
      <table>
        <thead><tr><th>Staff</th><th>Name</th><th>Client</th><th>Date</th><th>End of week</th>{props.admin ? <th>Follow-up</th> : null}</tr></thead>
        <tbody>
          {props.rows.map((row) => (
            <tr key={row.id || row.noteName + row.date}>
              <td>{row.staffName}</td>
              <td>{row.noteName}</td>
              <td>{row.clientName}</td>
              <td>{row.date}</td>
              <td>{row.pastEndOfWeek ? 'Past Friday' : 'This week'}</td>
              {props.admin ? (
                <td>
                  <span className="muted">{props.followOf(row.id) || 'Not recorded'}</span>
                  <button className="btn btn--ghost" type="button" onClick={() => props.onFollow(row, 'unicorn')}>UNICORN</button>
                  <button className="btn btn--ghost" type="button" onClick={() => props.onFollow(row, 'par')}>PAR fax</button>
                  <button className="btn btn--ghost" type="button" onClick={() => props.onFollow(row, 'done')}>Done</button>
                </td>
              ) : null}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
