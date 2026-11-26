import { useEffect, useState, type FormEvent, type ReactNode } from 'react';
import { ApiError, api } from '../../api';
import { TileWithSetup } from '../setup/TileSetup';

type Step = { id: string; label: string; owner: string; doneOn: string };
type Draft = {
  id: string;
  label: string;
  body: string;
  incomplete: boolean;
  status: 'draft' | 'approved';
  approvedBy: string;
  approvedOn: string;
};
type Todo = { id: string; title: string; owner: string; dueOn: string; reminderOn: string; doneOn: string; overdue?: boolean };
type Referral = {
  id: string;
  familyName: string;
  kind: '' | 'ei' | 'private';
  status: string;
  serviceCoordinator: string;
  assignedTo: string;
  assignedEmail: string;
  notes: string;
  receivedOn: string;
  lastContactOn: string;
  serviceGridOn: string;
  placedOn: string;
  removedOn: string;
  source: string;
  therapies: string[];
  steps: Step[];
  drafts: Draft[];
  todos: Todo[];
  updatedAt: string;
  updatedBy: string;
  daysToPlace: number | null;
  daysOverGoal: boolean;
  elapsedWorkingDays: number | null;
  remainingWorkingDays: number | null;
  stepsUnconfirmed: boolean;
  outreachCount?: number;
  respondBy?: string;
  responseDaysLeft?: number | null;
  atBottom?: boolean;
  sourceLabel?: string;
  sourceHow?: string;
  sourceNote?: string;
  paperworkLine?: string;
  dpExpiresOn?: string;
  firstOutreachOn?: string;
};

type SourceInfo = { id: string; label: string; how: string; note: string; paperwork: boolean };

type Provider = { staffName: string; staffEmail: string; therapyIds: string[]; therapies: string[]; week: number; next30: number };
type Slot = { staffName: string; staffEmail: string; therapies: string[]; therapyIds: string[]; date: string; day: string; time: string };
type Openings = {
  today: string;
  weekStart: string;
  weekEnd: string;
  through: string;
  total: number;
  weekTotal: number;
  disciplineCount: number;
  unlabeled: number;
  goalMax: number;
  overGoal: boolean;
  byMonth: Array<{ month: string; label: string; count: number }>;
  byTherapy: Array<{ label: string; count: number }>;
  providers: Provider[];
  slots: Slot[];
};

const THERAPIES = [
  { id: 'slp', label: 'Speech' },
  { id: 'ot', label: 'Occupational therapy' },
  { id: 'pt', label: 'Physical therapy' },
  { id: 'psych', label: 'Psychology' },
  { id: 'other', label: 'Other' },
];

function showDate(iso: string): string {
  const day = String(iso || '').slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(day)) return iso || '—';
  const parts = day.split('-').map(Number);
  return new Intl.DateTimeFormat('en-US', { timeZone: 'UTC', month: 'long', day: 'numeric', year: 'numeric' }).format(new Date(Date.UTC(parts[0], parts[1] - 1, parts[2])));
}

function kindLabel(kind: string): string {
  if (kind === 'ei') return 'Early Intervention';
  if (kind === 'private') return 'Private';
  return 'Not set yet';
}

function daysText(row: Referral): string {
  if (row.daysToPlace != null) return row.daysToPlace + (row.daysOverGoal ? ' working days, past 28' : ' working days');
  if (row.daysOverGoal) return 'Past 28 working days';
  return '—';
}

function draftHint(id: string): string {
  if (id === 'ei-first') return 'The rest of this message was not written down. Replace the sample name and finish it. Approving does not send it.';
  if (id === 'ei-last' || id === 'private-first') return 'A sample name is still in this draft. Change it if you need to. Approving does not send it.';
  if (id === 'sc-no-response') return 'Type the phone number into the message. This is not sent on a schedule. Approving does not send it.';
  return 'The wording was not written down. Write it here. Approving does not send it.';
}

function snapshot(row: Referral): string {
  return JSON.stringify({
    familyName: row.familyName,
    kind: row.kind,
    serviceCoordinator: row.serviceCoordinator,
    notes: row.notes,
    receivedOn: row.receivedOn,
    lastContactOn: row.lastContactOn,
    serviceGridOn: row.serviceGridOn,
    placedOn: row.placedOn,
    source: row.source,
    therapies: row.therapies,
  });
}

function messageOf(err: unknown): string {
  return err instanceof Error ? err.message : 'Could not save the referral';
}

function SourceField(props: { value: string; sources: SourceInfo[]; onChange: (value: string) => void }) {
  const current = props.sources.find((item) => item.id === props.value);
  const known = !!current || !props.value;
  return (
    <div>
      <label>Board or source
        <select value={props.value} onChange={(event) => props.onChange(event.target.value)}>
          <option value="">Choose</option>
          {!known ? <option value={props.value}>{props.value}</option> : null}
          {props.sources.map((item) => <option key={item.id} value={item.id}>{item.label}</option>)}
        </select>
      </label>
      {current && current.how ? <p className="muted">{current.how}. {current.note}</p> : null}
      {current && current.paperwork ? <p className="muted">When you reply to this organization, include this line: Please contact our referrals team at referrals@emmeco.org with any paperwork or questions.</p> : null}
    </div>
  );
}

type Tone = 'green' | 'pink' | 'yellow' | 'blue';

function ListTabs(props: {
  label: string;
  value: string;
  onChange: (id: string) => void;
  tabs: Array<{ id: string; label: string; tone: Tone; disabled?: boolean }>;
}) {
  return (
    <div className="ref-tabs" role="tablist" aria-label={props.label}>
      {props.tabs.map((tab) => (
        <button
          key={tab.id}
          type="button"
          role="tab"
          aria-selected={props.value === tab.id}
          disabled={tab.disabled}
          className={'ref-tab ref-tab--' + tab.tone + (props.value === tab.id ? ' is-on' : '')}
          onClick={() => props.onChange(tab.id)}
        >{tab.label}</button>
      ))}
    </div>
  );
}

function ListBand(props: { tone: Tone; title: string; children: ReactNode }) {
  return (
    <section className={'ref-band ref-band--' + props.tone}>
      <h2>{props.title}</h2>
      {props.children}
    </section>
  );
}

function FamilyTable(props: { rows: Referral[]; empty: string; onOpen: (row: Referral) => void }) {
  if (!props.rows.length) return <p className="empty">{props.empty}</p>;
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Family</th><th>Type</th><th>Status</th><th>Outreaches</th><th>Assigned</th><th>Received</th><th>Last contact</th><th>Days to place</th><th>Open to-dos</th><th></th>
          </tr>
        </thead>
        <tbody>
          {props.rows.map((row) => (
            <tr key={row.id}>
              <td>{row.familyName}</td>
              <td>{kindLabel(row.kind)}</td>
              <td>{row.status === 'Bottom' ? 'Bottom of the list' : row.status}</td>
              <td>{row.kind === 'ei' ? (row.outreachCount || 0) + ' of 2' : '—'}</td>
              <td>{row.assignedTo || '—'}</td>
              <td>{showDate(row.receivedOn)}</td>
              <td>{showDate(row.lastContactOn)}</td>
              <td>{daysText(row)}</td>
              <td>{row.todos.filter((todo) => !todo.doneOn).length}</td>
              <td><button type="button" className="btn btn--ghost" onClick={() => props.onOpen(row)}>Open</button></td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function TherapyChoices(props: { selected: string[]; onChange: (next: string[]) => void }) {
  function toggle(id: string) {
    const has = props.selected.includes(id);
    props.onChange(has ? props.selected.filter((item) => item !== id) : props.selected.concat(id));
  }
  return (
    <div className="ref-choices">
      {THERAPIES.map((item) => (
        <label key={item.id} className="ref-check">
          <input type="checkbox" checked={props.selected.includes(item.id)} onChange={() => toggle(item.id)} />
          {item.label}
        </label>
      ))}
    </div>
  );
}

function StepEditor(props: {
  step: Step;
  clock: string;
  busy: boolean;
  onDone: (stepId: string, doneOn: string, owner: string) => void;
  onUndo: (stepId: string) => void;
  onOwner: (stepId: string, owner: string) => void;
}) {
  const [date, setDate] = useState(props.step.doneOn || props.clock);
  const [owner, setOwner] = useState(props.step.owner);
  useEffect(() => {
    setDate(props.step.doneOn || props.clock);
    setOwner(props.step.owner);
  }, [props.step.doneOn, props.step.owner, props.clock]);
  return (
    <li>
      <strong>{props.step.label}</strong>
      {props.step.doneOn ? <span className="muted">Done {showDate(props.step.doneOn)}</span> : null}
      <label>Who<input value={owner} onChange={(event) => setOwner(event.target.value)} onBlur={() => { if (owner !== props.step.owner) props.onOwner(props.step.id, owner); }} /></label>
      {props.step.doneOn ? (
        <button type="button" className="btn btn--ghost" disabled={props.busy} onClick={() => props.onUndo(props.step.id)}>Undo</button>
      ) : (
        <>
          <label>Date<input type="date" value={date} onChange={(event) => setDate(event.target.value)} /></label>
          <button type="button" className="btn btn--ghost" disabled={props.busy || !date} onClick={() => props.onDone(props.step.id, date, owner)}>Mark done</button>
        </>
      )}
    </li>
  );
}

function ReferralsOverview() {
  const [rows, setRows] = useState<Referral[]>([]);
  const [clock, setClock] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const [showRemoved, setShowRemoved] = useState(false);
  const [section, setSection] = useState<'families' | 'add' | 'family' | 'openings'>('families');
  const [familyPart, setFamilyPart] = useState<'details' | 'steps' | 'messages' | 'todos'>('details');
  const [editor, setEditor] = useState<Referral | null>(null);
  const [baseline, setBaseline] = useState('');
  const [draftText, setDraftText] = useState<Record<string, string>>({});
  const [familyName, setFamilyName] = useState('');
  const [kind, setKind] = useState('');
  const [serviceCoordinator, setServiceCoordinator] = useState('');
  const [therapies, setTherapies] = useState<string[]>([]);
  const [source, setSource] = useState('');
  const [receivedOn, setReceivedOn] = useState('');
  const [serviceGridOn, setServiceGridOn] = useState('');
  const [notes, setNotes] = useState('');
  const [openings, setOpenings] = useState<Openings | null>(null);
  const [openingsError, setOpeningsError] = useState('');
  const [kindEdits, setKindEdits] = useState<Record<string, string[]>>({});
  const [slotWindow, setSlotWindow] = useState('30');
  const [therapyFilter, setTherapyFilter] = useState('');
  const [assignEmail, setAssignEmail] = useState('');
  const [todoTitle, setTodoTitle] = useState('');
  const [todoOwner, setTodoOwner] = useState('Administrator');
  const [todoDue, setTodoDue] = useState('');
  const [sources, setSources] = useState<SourceInfo[]>([]);

  function adopt(referral: Referral) {
    setSection('family');
    setFamilyPart('details');
    setEditor(referral);
    setBaseline(snapshot(referral));
    const texts: Record<string, string> = {};
    referral.drafts.forEach((draft) => { texts[draft.id] = draft.body; });
    setDraftText(texts);
    setRows((current) => {
      const exists = current.some((row) => row.id === referral.id);
      return exists ? current.map((row) => row.id === referral.id ? referral : row) : [referral].concat(current);
    });
  }

  function loadOpenings() {
    api<Openings>('/api/referrals?view=openings')
      .then((data) => { setOpenings(data); setOpeningsError(''); })
      .catch((err: Error) => setOpeningsError(err.message));
  }

  function loadList() {
    api<{ referrals: Referral[]; today: string; sources?: SourceInfo[] }>('/api/referrals')
      .then((data) => {
        setRows(data.referrals || []);
        setSources(data.sources || []);
        setClock(data.today || '');
        if (!receivedOn) setReceivedOn(data.today || '');
        if (!todoDue) setTodoDue(data.today || '');
        setError('');
      })
      .catch((err: Error) => setError(err.message));
  }

  useEffect(() => { loadList(); loadOpenings(); }, []);

  useEffect(() => {
    if (!openings) return;
    const next: Record<string, string[]> = {};
    openings.providers.forEach((provider) => { next[provider.staffEmail] = provider.therapyIds.slice(); });
    setKindEdits(next);
  }, [openings]);

  async function post(current: Referral, body: Record<string, unknown>) {
    return api<{ referral: Referral }>('/api/referrals/' + encodeURIComponent(current.id), {
      method: 'POST',
      body: Object.assign({ id: current.id, baseUpdatedAt: current.updatedAt }, body),
    });
  }

  async function commit(action: string, extra: Record<string, unknown> = {}, sourceRow?: Referral): Promise<boolean> {
    const start = sourceRow || editor;
    if (!start) return false;
    setBusy(true);
    setError('');
    try {
      let current = start;
      const fields = {
        action: 'save',
        familyName: current.familyName,
        kind: current.kind,
        serviceCoordinator: current.serviceCoordinator,
        notes: current.notes,
        receivedOn: current.receivedOn,
        lastContactOn: current.lastContactOn,
        serviceGridOn: current.serviceGridOn,
        placedOn: current.placedOn,
        source: current.source,
        therapies: current.therapies,
      };
      if (action === 'save' || snapshot(current) !== baseline) current = (await post(current, fields)).referral;
      for (let i = 0; i < current.drafts.length; i++) {
        const draft = current.drafts[i];
        if ((action === 'approve' || action === 'save-draft') && extra.draftId === draft.id) continue;
        const local = draftText[draft.id];
        if (local != null && local !== draft.body) current = (await post(current, { action: 'save-draft', draftId: draft.id, body: local })).referral;
      }
      if (action !== 'save') {
        const body = Object.assign({}, extra, { action: action });
        if ((action === 'approve' || action === 'save-draft') && typeof extra.draftId === 'string') {
          const typed = draftText[extra.draftId];
          if (typed != null) body.body = typed;
        }
        current = (await post(current, body)).referral;
      }
      adopt(current);
      return true;
    } catch (err) {
      setError(messageOf(err));
      if (err instanceof ApiError && err.status === 409) loadList();
      return false;
    } finally {
      setBusy(false);
    }
  }

  async function addReferral(event: FormEvent) {
    event.preventDefault();
    setBusy(true);
    setError('');
    try {
      const saved = await api<{ referral: Referral }>('/api/referrals', {
        method: 'POST',
        body: { familyName, kind, serviceCoordinator, therapies, source, notes, receivedOn, serviceGridOn },
      });
      setFamilyName('');
      setKind('');
      setServiceCoordinator('');
      setTherapies([]);
      setSource('');
      setServiceGridOn('');
      setNotes('');
      adopt(saved.referral);
    } catch (err) {
      setError(messageOf(err));
    } finally {
      setBusy(false);
    }
  }

  function patch(partial: Partial<Referral>) {
    setEditor((current) => current ? Object.assign({}, current, partial) : current);
  }

  async function saveKinds() {
    if (!openings || !openings.providers.length) return;
    setBusy(true);
    setOpeningsError('');
    try {
      await api('/api/referrals', {
        method: 'POST',
        body: {
          action: 'provider-kinds',
          providers: openings.providers.map((provider) => ({
            email: provider.staffEmail,
            name: provider.staffName,
            therapies: kindEdits[provider.staffEmail] || provider.therapyIds,
          })),
        },
      });
      loadOpenings();
    } catch (err) {
      setOpeningsError(messageOf(err));
    } finally {
      setBusy(false);
    }
  }

  const activeRows = rows.filter((row) => row.status !== 'Bottom' && row.status !== 'Removed');
  const bottomRows = rows.filter((row) => row.status === 'Bottom');
  const removedRows = rows.filter((row) => row.status === 'Removed');
  const slots = (openings ? openings.slots : []).filter((slot) => {
    if (therapyFilter === 'unset') return slot.therapyIds.length === 0;
    if (therapyFilter && slot.therapyIds.indexOf(therapyFilter) === -1) return false;
    if (slotWindow === 'week' && openings) return slot.date >= openings.weekStart && slot.date <= openings.weekEnd;
    return true;
  });
  const openTodos = editor ? editor.todos.filter((todo) => !todo.doneOn) : [];
  const doneTodos = editor ? editor.todos.filter((todo) => todo.doneOn) : [];

  return (
    <section>
      <h1>Referrals</h1>
      <p className="lede">
        Track each Early Intervention and private family from the first contact through placement.
        The family, status, assigned staff, and notes stay on the Referrals SharePoint list.
        Messages stay here as drafts until a person sends them.
      </p>
      <div className="ref-note">
        <strong>Early Intervention families get 15 calendar days and 2 outreaches to respond.</strong>
        <p>The 15 days start when you record the first outreach. After the second outreach, if they have not confirmed by the end of those 15 days, they move to the bottom of the list. They stay there until the 28 working days from the service grid are up, or until the family confirms.</p>
        <p>Sara Beth has been giving families about a week. This screen uses the longer window. Recording an outreach does not send a message. Approving a draft does not send it either.</p>
        <p>Private steps are still a draft. Assigning a staff member does not email the provider, the family, or the service coordinator. Medicaid and insurance go on the PAR/HIA form as a to-do. The referrals spreadsheet is a to-do too, because this screen cannot open it.</p>
      </div>
      {error ? <p className="error" role="alert">{error}</p> : null}

      <ListTabs
        label="Referral lists"
        value={section}
        onChange={(id) => setSection(id as 'families' | 'add' | 'family' | 'openings')}
        tabs={[
          { id: 'families', label: 'Families', tone: 'green' },
          { id: 'add', label: 'Add a referral', tone: 'pink' },
          { id: 'family', label: editor ? editor.familyName : 'This family', tone: 'yellow', disabled: !editor },
          { id: 'openings', label: 'Openings', tone: 'blue' },
        ]}
      />

      {section === 'add' ? (
      <ListBand tone="pink" title="Add a referral">
      <form className="panel" onSubmit={addReferral}>
        <div className="row">
          <label>Family name<input value={familyName} onChange={(event) => setFamilyName(event.target.value)} required /></label>
          <label>Early Intervention or private
            <select value={kind} onChange={(event) => setKind(event.target.value)} required>
              <option value="">Choose</option>
              <option value="ei">Early Intervention</option>
              <option value="private">Private</option>
            </select>
          </label>
          <label>Service coordinator<input value={serviceCoordinator} onChange={(event) => setServiceCoordinator(event.target.value)} /></label>
        </div>
        <p className="muted">Kind of therapy</p>
        <TherapyChoices selected={therapies} onChange={setTherapies} />
        <div className="row">
          <SourceField value={source} sources={sources} onChange={setSource} />
          <label>Received<input type="date" value={receivedOn} onChange={(event) => setReceivedOn(event.target.value)} /></label>
          <label>Service grid date<input type="date" value={serviceGridOn} onChange={(event) => setServiceGridOn(event.target.value)} /></label>
        </div>
        <p className="muted">The service grid date is how Early Intervention days to place are counted. Weekends are skipped. Holidays are not skipped yet.</p>
        <label>Notes<textarea value={notes} onChange={(event) => setNotes(event.target.value)} /></label>
        <div className="row"><button className="btn" type="submit" disabled={busy}>Add referral</button></div>
      </form>
      </ListBand>
      ) : null}

      {section === 'families' ? (
        <>
          <label className="ref-check">
            <input type="checkbox" checked={showRemoved} onChange={(event) => setShowRemoved(event.target.checked)} />
            Show families who were removed
          </label>
          <ListBand tone="green" title={'Active families (' + activeRows.length + ')'}>
            <FamilyTable rows={activeRows} empty="No active families." onOpen={adopt} />
          </ListBand>
          <ListBand tone="yellow" title={'Bottom of the list (' + bottomRows.length + ')'}>
            <FamilyTable rows={bottomRows} empty="No families are waiting at the bottom." onOpen={adopt} />
          </ListBand>
          {showRemoved ? (
            <ListBand tone="pink" title={'Removed (' + removedRows.length + ')'}>
              <FamilyTable rows={removedRows} empty="No removed families." onOpen={adopt} />
            </ListBand>
          ) : null}
        </>
      ) : null}

      {section === 'family' && editor ? (
        <div className="panel">
          <ListTabs
            label="This family"
            value={familyPart}
            onChange={(id) => setFamilyPart(id as 'details' | 'steps' | 'messages' | 'todos')}
            tabs={[
              { id: 'details', label: 'Details', tone: 'yellow' },
              { id: 'steps', label: 'Steps', tone: 'green' },
              { id: 'messages', label: 'Messages', tone: 'pink' },
              { id: 'todos', label: 'To-dos', tone: 'blue' },
            ]}
          />
          {familyPart === 'details' ? (
          <ListBand tone="yellow" title={editor.familyName}>
          <p className="muted">{kindLabel(editor.kind)} · {editor.status === 'Bottom' ? 'Bottom of the list' : editor.status}{editor.updatedBy ? ' · Last saved by ' + editor.updatedBy : ''}</p>
          {editor.stepsUnconfirmed ? <p className="ref-note">These private steps are a draft: first contact, needs discussion, placement, follow up, removal. They are not confirmed yet.</p> : null}
          {editor.kind === 'ei' ? <p className="muted">Who does each Early Intervention step was not named, so the owner starts as Administrator. You can change it. Days are working days after the service grid date, through the placed date. The draft goal is 28. Weekends are skipped. Holidays are not skipped yet.</p> : null}
          {editor.status === 'Bottom' ? <p className="ref-warn">This family is at the bottom of the list. They had 2 outreaches and 15 days to respond. They stay here until the 28 working days are up or they confirm.</p> : null}
          {editor.daysOverGoal && editor.status !== 'Answered' && editor.status !== 'Placed' ? <p className="ref-warn">This referral is past 28 working days from the service grid. If the family has not confirmed, remove them when you are finished with the referral.</p> : null}
          {editor.daysOverGoal && (editor.status === 'Answered' || editor.status === 'Placed') ? <p className="muted">The family has confirmed. This referral is still past 28 working days from the service grid.</p> : null}
          {editor.kind === 'ei' && editor.elapsedWorkingDays != null && !editor.placedOn ? <p className="muted">Working days since the service grid: {editor.elapsedWorkingDays}. {editor.remainingWorkingDays != null && editor.remainingWorkingDays >= 0 ? editor.remainingWorkingDays + ' left in the 28.' : ''}</p> : null}
          {editor.daysToPlace != null ? <p className="muted">Days to place: {editor.daysToPlace} working days.</p> : null}

          <div className="row">
            <label>Family name<input value={editor.familyName} onChange={(event) => patch({ familyName: event.target.value })} required /></label>
            <label>Early Intervention or private
              <select value={editor.kind} onChange={(event) => patch({ kind: event.target.value as Referral['kind'] })}>
                <option value="">Choose</option>
                <option value="ei">Early Intervention</option>
                <option value="private">Private</option>
              </select>
            </label>
            <label>Service coordinator<input value={editor.serviceCoordinator} onChange={(event) => patch({ serviceCoordinator: event.target.value })} /></label>
          </div>
          <p className="muted">Kind of therapy</p>
          <TherapyChoices selected={editor.therapies} onChange={(next) => patch({ therapies: next })} />
          <div className="row">
            <SourceField value={editor.source} sources={sources} onChange={(value) => patch({ source: value })} />
            <label>Received<input type="date" value={editor.receivedOn.length >= 10 ? editor.receivedOn.slice(0, 10) : ''} onChange={(event) => patch({ receivedOn: event.target.value })} /></label>
            <label>Last contact<input type="date" value={editor.lastContactOn} onChange={(event) => patch({ lastContactOn: event.target.value })} /></label>
            <label>Service grid date<input type="date" value={editor.serviceGridOn} onChange={(event) => patch({ serviceGridOn: event.target.value })} /></label>
            <label>Placed<input type="date" value={editor.placedOn} onChange={(event) => patch({ placedOn: event.target.value })} /></label>
          </div>
          <label>Notes<textarea value={editor.notes} onChange={(event) => patch({ notes: event.target.value })} /></label>
          <div className="row">
            <button type="button" className="btn" disabled={busy} onClick={() => commit('save')}>Save</button>
            <button type="button" className="btn btn--ghost" disabled={busy || !clock} onClick={() => commit('save', {}, Object.assign({}, editor, { lastContactOn: clock }))}>Record contact today</button>
            <button type="button" className="btn btn--ghost" disabled={busy || !clock} onClick={() => commit('outreach')}>Record an outreach</button>
            <button type="button" className="btn btn--ghost" disabled={busy} onClick={() => commit('answered')}>Family confirmed</button>
            <button type="button" className="btn btn--ghost" disabled={busy} onClick={() => { setEditor(null); setSection('families'); }}>Close</button>
          </div>
          <p className="muted">
            Outreaches recorded: {editor.outreachCount || 0} of 2.
            {editor.respondBy
              ? (editor.responseDaysLeft != null
                ? ' They can respond through ' + showDate(editor.respondBy) + ' (' + editor.responseDaysLeft + ' days left after today).'
                : ' The 15 days ended on ' + showDate(editor.respondBy) + '.')
              : ' The 15 days start when you record the first outreach.'}
            {' '}Record an outreach only after you have reached the family. This does not send a message.
          </p>
          {editor.dpExpiresOn ? <p className="ref-warn">Developmental Pathways: reply by {showDate(editor.dpExpiresOn)}. These referrals expire about 48 hours after the received date.</p> : null}
          {editor.paperworkLine ? <p className="muted">Include this when you reply: {editor.paperworkLine}</p> : null}
          <p className="muted">Family confirmed sends the referral to Sara Beth for onboarding as a to-do and takes them off the bottom of the list. If the family was removed, they come back onto this list. Nothing is emailed.</p>
          </ListBand>
          ) : null}

          {familyPart === 'steps' ? (
          <ListBand tone="green" title="Steps">
          {editor.steps.length === 0 ? <p className="empty">Choose Early Intervention or private, then save, and the checklist will start.</p> : (
            <ul className="ref-steps">
              {editor.steps.map((step) => (
                <StepEditor
                  key={step.id}
                  step={step}
                  clock={clock}
                  busy={busy}
                  onDone={(stepId, doneOn, owner) => commit('step', { stepId: stepId, doneOn: doneOn, owner: owner })}
                  onUndo={(stepId) => commit('step', { stepId: stepId, doneOn: '' })}
                  onOwner={(stepId, owner) => commit('step', { stepId: stepId, owner: owner })}
                />
              ))}
            </ul>
          )}

          </ListBand>
          ) : null}

          {familyPart === 'details' ? (
          <ListBand tone="blue" title="Assign a staff member">
          <p className="muted">Assigned to {editor.assignedTo || 'nobody yet'}{editor.assignedEmail ? ' · ' + editor.assignedEmail : ''}</p>
          {openings && openings.providers.length ? (
            <div className="row">
              <label>Staff member
                <select value={assignEmail} onChange={(event) => setAssignEmail(event.target.value)}>
                  <option value="">Choose</option>
                  {openings.providers.map((provider) => (
                    <option key={provider.staffEmail} value={provider.staffEmail}>{provider.staffName}</option>
                  ))}
                </select>
              </label>
              <button
                type="button"
                className="btn"
                disabled={busy || !assignEmail}
                onClick={() => {
                  const provider = (openings.providers || []).find((item) => item.staffEmail === assignEmail);
                  if (provider) commit('assign', { assignedTo: provider.staffName, assignedEmail: provider.staffEmail });
                }}
              >Assign</button>
            </div>
          ) : <p className="muted">Staff names appear here after a provider saves a schedule. You can still assign from the openings list below once it loads.</p>}
          </ListBand>
          ) : null}

          {familyPart === 'messages' ? (
          <ListBand tone="pink" title="Messages">
          <p className="muted">These are drafts. Approving one puts a to-do on this referral so a person can send it. This screen does not send it.</p>
          {editor.drafts.length === 0 ? <p className="empty">Messages appear after you choose Early Intervention or private.</p> : editor.drafts.map((draft) => (
            <div key={draft.id} className="panel">
              <strong>{draft.label}</strong>
              <p className="muted">{draftHint(draft.id)}</p>
              {draft.status === 'approved' ? <p className="muted">Approved by {draft.approvedBy} on {showDate(draft.approvedOn)}. This was not sent.</p> : <p className="muted">Draft</p>}
              <textarea value={draftText[draft.id] ?? draft.body} onChange={(event) => setDraftText((current) => Object.assign({}, current, { [draft.id]: event.target.value }))} />
              <div className="row">
                <button type="button" className="btn btn--ghost" disabled={busy} onClick={() => commit('save-draft', { draftId: draft.id })}>Save message</button>
                <button type="button" className="btn" disabled={busy} onClick={() => commit('approve', { draftId: draft.id })}>Approve</button>
              </div>
            </div>
          ))}

          </ListBand>
          ) : null}

          {familyPart === 'todos' ? (
          <ListBand tone="blue" title="To-dos">
          <p className="muted">Anything this screen cannot do is a to-do with an owner, a due date, and a reminder on that date.</p>
          {openTodos.length === 0 ? <p className="empty">No open to-dos.</p> : (
            <div className="table-wrap">
              <table>
                <thead><tr><th>Task</th><th>Owner</th><th>Due</th><th>Reminder</th><th></th></tr></thead>
                <tbody>
                  {openTodos.map((todo) => (
                    <tr key={todo.id}>
                      <td>{todo.title}{todo.overdue ? <span className="error"> Overdue</span> : null}</td>
                      <td>{todo.owner}</td>
                      <td>{showDate(todo.dueOn)}</td>
                      <td>{showDate(todo.reminderOn)}</td>
                      <td><button type="button" className="btn btn--ghost" disabled={busy} onClick={() => commit('todo-done', { todoId: todo.id })}>Mark done</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
          <div className="row">
            <label>Task<input value={todoTitle} onChange={(event) => setTodoTitle(event.target.value)} /></label>
            <label>Owner<input value={todoOwner} onChange={(event) => setTodoOwner(event.target.value)} /></label>
            <label>Due<input type="date" value={todoDue} onChange={(event) => setTodoDue(event.target.value)} /></label>
            <button
              type="button"
              className="btn btn--ghost"
              disabled={busy || !todoTitle || !todoOwner || !todoDue}
              onClick={() => {
                void commit('add-todo', { title: todoTitle, owner: todoOwner, dueOn: todoDue }).then((ok) => { if (ok) setTodoTitle(''); });
              }}
            >Add to-do</button>
          </div>
          {doneTodos.length ? (
            <details>
              <summary>Finished to-dos ({doneTodos.length})</summary>
              <ul className="setup-list">
                {doneTodos.map((todo) => <li key={todo.id}>{todo.title} · {todo.owner} · done {showDate(todo.doneOn)}</li>)}
              </ul>
            </details>
          ) : null}
          </ListBand>
          ) : null}
        </div>
      ) : null}

      {section === 'openings' ? (
      <>
      <h2 className="section-heading">Therapist openings</h2>
      <p className="lede">
        Openings are read from each provider’s schedule for this week and the next 30 days.
        They are not copied into another list. This week means openings still ahead, Monday through Sunday.
      </p>
      {openingsError ? <p className="error">{openingsError}</p> : null}
      {!openings && !openingsError ? <p className="muted">Loading openings…</p> : null}
      {openings ? (
        <>
          <div className={openings.overGoal ? 'ref-warn' : 'ref-note'}>
            <strong>{openings.total} openings in the next 30 days.</strong>
            <p>The draft goal is no more than {openings.goalMax}. This week: {openings.weekTotal}. Kinds of therapy with an opening: {openings.disciplineCount}.</p>
            {openings.unlabeled ? <p>{openings.unlabeled} openings have no kind of therapy set. Choose the kind next to the provider.</p> : null}
            {openings.overGoal ? <p>There are more than {openings.goalMax} openings.</p> : null}
          </div>
          <div className="ref-kpis">
            <p>This week<br />{showDate(openings.weekStart)} – {showDate(openings.weekEnd)}<br /><strong>{openings.weekTotal}</strong></p>
            <p>Next 30 days<br />through {showDate(openings.through)}<br /><strong>{openings.total}</strong></p>
          </div>
          <ListBand tone="green" title="By kind of therapy">
          {openings.byTherapy.length ? (
            <div className="table-wrap">
              <table>
                <thead><tr><th>Kind of therapy</th><th>Openings</th></tr></thead>
                <tbody>
                  {openings.byTherapy.map((row) => <tr key={row.label}><td>{row.label}</td><td>{row.count}</td></tr>)}
                </tbody>
              </table>
            </div>
          ) : <p className="empty">No openings have a kind of therapy yet.</p>}
          </ListBand>
          <ListBand tone="pink" title="By month">
          {openings.byMonth.length ? (
            <div className="table-wrap">
              <table>
                <thead><tr><th>Month</th><th>Openings</th></tr></thead>
                <tbody>
                  {openings.byMonth.map((row) => <tr key={row.month}><td>{row.label}</td><td>{row.count}</td></tr>)}
                </tbody>
              </table>
            </div>
          ) : <p className="empty">No openings in the next 30 days.</p>}
          </ListBand>
          <ListBand tone="yellow" title="By provider">
          {openings.providers.length === 0 ? <p className="empty">No provider schedules yet. They appear after a provider saves a schedule.</p> : (
            <>
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr><th>Provider</th><th>Kind of therapy</th><th>This week</th><th>Next 30 days</th><th></th></tr>
                  </thead>
                  <tbody>
                    {openings.providers.map((provider) => (
                      <tr key={provider.staffEmail}>
                        <td>{provider.staffName}</td>
                        <td>
                          <div className="ref-choices">
                            {THERAPIES.map((item) => {
                              const selected = (kindEdits[provider.staffEmail] || provider.therapyIds).includes(item.id);
                              return (
                                <label key={item.id} className="ref-check">
                                  <input
                                    type="checkbox"
                                    checked={selected}
                                    onChange={() => {
                                      const current = kindEdits[provider.staffEmail] || provider.therapyIds;
                                      const next = current.includes(item.id) ? current.filter((id) => id !== item.id) : current.concat(item.id);
                                      setKindEdits((prev) => Object.assign({}, prev, { [provider.staffEmail]: next }));
                                    }}
                                  />
                                  {item.label}
                                </label>
                              );
                            })}
                          </div>
                        </td>
                        <td>{provider.week}</td>
                        <td>{provider.next30}</td>
                        <td>
                          <button
                            type="button"
                            className="btn btn--ghost"
                            disabled={busy || !editor}
                            onClick={() => commit('assign', { assignedTo: provider.staffName, assignedEmail: provider.staffEmail })}
                          >Assign</button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="row">
                <button type="button" className="btn" disabled={busy} onClick={saveKinds}>Save kinds of therapy</button>
                {!editor ? <p className="muted">Open a family on the Families tab before assigning a provider.</p> : <p className="muted">Assign adds {editor.familyName} to that staff member and makes to-dos. It does not send a message or hold the opening.</p>}
              </div>
            </>
          )}
          </ListBand>
          <ListBand tone="blue" title="Open times">
          <div className="row">
            <label>Show
              <select value={slotWindow} onChange={(event) => setSlotWindow(event.target.value)}>
                <option value="week">This week</option>
                <option value="30">The next 30 days</option>
              </select>
            </label>
            <label>Kind of therapy
              <select value={therapyFilter} onChange={(event) => setTherapyFilter(event.target.value)}>
                <option value="">All</option>
                {THERAPIES.map((item) => <option key={item.id} value={item.id}>{item.label}</option>)}
                <option value="unset">Not set</option>
              </select>
            </label>
          </div>
          {slots.length === 0 ? <p className="empty">No openings in this view.</p> : (
            <div className="table-wrap">
              <table>
                <thead><tr><th>Provider</th><th>Kind of therapy</th><th>Date</th><th>Day</th><th>Time</th></tr></thead>
                <tbody>
                  {slots.map((slot, index) => (
                    <tr key={slot.staffEmail + slot.date + slot.time + index}>
                      <td>{slot.staffName}</td>
                      <td>{slot.therapies.join(', ')}</td>
                      <td>{showDate(slot.date)}</td>
                      <td>{slot.day}</td>
                      <td>{slot.time}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
          </ListBand>
        </>
      ) : null}
      </>
      ) : null}
    </section>
  );
}

export function ReferralsTile() {
  return (
    <TileWithSetup hashId="referrals" tileId="referrals" label="Referrals">
      <ReferralsOverview />
    </TileWithSetup>
  );
}
