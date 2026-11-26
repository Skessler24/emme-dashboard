import { useEffect, useRef, useState } from 'react';
import { api } from '../../api';

type Topic = { id: string; title: string; foundation: boolean; prompt: string; e: string; p: string; r: string };
type Meeting = {
  id: string;
  date: string;
  update: string;
  background: string;
  clip: string;
  questions: string;
  noticed: string;
  action: string;
  observe: string;
  revisit: string;
  growthDates: string;
};
type GrowthMap = { email: string; name: string; updatedAt: string; updatedBy: string; topics: Topic[]; meetings: Meeting[] };
type Provider = { email: string; name: string; standing?: { short: string; label: string } };
type Visit = { date: string; day: string; time: string; label: string };
type Standing = { short: string; label: string };
type OutlookEvent = { start: string; end: string; subject: string };
type Appointment = {
  id: string;
  providerEmail: string;
  adminEmail: string;
  adminName: string;
  start: string;
  end: string;
  suggestedStart: string;
  suggestedEnd: string;
  status: string;
};
type MapResponse = { role: string; me: string; map: GrowthMap };

const MEETING_FIELDS: Array<{ key: keyof Meeting; label: string; kind: 'date' | 'text' }> = [
  { key: 'date', label: 'Meeting date', kind: 'date' },
  { key: 'update', label: 'Brief update: what I tried since last time, and what got in the way', kind: 'text' },
  { key: 'background', label: 'Background that helps the meeting (family priorities, who lives in the family, what you already tried)', kind: 'text' },
  { key: 'clip', label: 'Clip timestamp (the video stays where you keep it today)', kind: 'text' },
  { key: 'questions', label: 'My questions and reflections', kind: 'text' },
  { key: 'noticed', label: 'Something I noticed or better understood', kind: 'text' },
  { key: 'action', label: 'One action I will try, and why', kind: 'text' },
  { key: 'observe', label: 'What I will observe or ask the family', kind: 'text' },
  { key: 'revisit', label: 'When will we revisit this?', kind: 'text' },
  { key: 'growthDates', label: 'Growth map dates to add', kind: 'text' },
];

function whenLabel(value: string): string {
  if (!value) return '';
  const [day, time] = value.split('T');
  return day + ' at ' + (time || '');
}

function statusLabel(status: string): string {
  if (status === 'proposed') return 'Waiting for the provider';
  if (status === 'suggested') return 'Provider suggested another time';
  if (status === 'accepted') return 'Accepted. Not on both calendars yet';
  if (status === 'confirmed') return 'On both calendars';
  if (status === 'declined') return 'Declined';
  if (status === 'withdrawn') return 'Withdrawn';
  return status;
}

export function ReflectiveMapPanel() {
  const [providers, setProviders] = useState<Provider[]>([]);
  const [provider, setProvider] = useState('');
  const [typed, setTyped] = useState('');
  const [map, setMap] = useState<GrowthMap | null>(null);
  const [me, setMe] = useState('');
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [loading, setLoading] = useState(true);
  const [visits, setVisits] = useState<Visit[]>([]);
  const [outlook, setOutlook] = useState<OutlookEvent[]>([]);
  const [outlookError, setOutlookError] = useState('');
  const [appointments, setAppointments] = useState<Appointment[]>([]);
  const [standing, setStanding] = useState<Standing | null>(null);
  const [start, setStart] = useState('');
  const [end, setEnd] = useState('');
  const [suggesting, setSuggesting] = useState('');
  const [suggestStart, setSuggestStart] = useState('');
  const [suggestEnd, setSuggestEnd] = useState('');
  const dirty = useRef(false);
  const timer = useRef<number | null>(null);
  const baseRef = useRef('');
  const latestRef = useRef<GrowthMap | null>(null);
  const saving = useRef(false);
  const emailRef = useRef('');
  const email = (typed.trim() || provider).trim().toLowerCase();
  emailRef.current = email;

  function applyMap(next: GrowthMap) {
    baseRef.current = next.updatedAt || '';
    latestRef.current = next;
    setMap(next);
    dirty.current = false;
  }

  function loadSide(who: string) {
    const query = who ? '&provider=' + encodeURIComponent(who) : '';
    api<{ visits: Visit[]; outlook: OutlookEvent[]; outlookError: string }>('/api/reflective-supervision?part=sessions' + query)
      .then(function (data) {
        setVisits(data.visits || []);
        setOutlook(data.outlook || []);
        setOutlookError(data.outlookError || '');
      })
      .catch(function () { setVisits([]); setOutlook([]); });
    api<{ appointments: Appointment[]; standing?: Standing }>('/api/reflective-supervision?part=appointments' + query)
      .then(function (data) {
        setAppointments(data.appointments || []);
        setStanding(data.standing || null);
      })
      .catch(function () { setAppointments([]); setStanding(null); });
  }

  function loadMap(who: string, quiet: boolean) {
    const query = who ? '?part=map&provider=' + encodeURIComponent(who) : '?part=map';
    if (!quiet) setLoading(true);
    api<MapResponse>('/api/reflective-supervision' + query)
      .then(function (data) {
        setMe(data.me || '');
        if (quiet && data.map.updatedAt === baseRef.current) return;
        if (dirty.current && quiet) {
          setNotice('The other person saved this map. Reload before you save, or you will be asked to start again.');
          return;
        }
        applyMap(data.map);
        setError('');
        if (!quiet) loadSide(who);
      })
      .catch(function (err: Error) {
        if (!quiet) {
          setMap(null);
          setError(err.message);
        }
      })
      .finally(function () { if (!quiet) setLoading(false); });
  }

  useEffect(function () {
    api<{ providers: Provider[] }>('/api/reflective-supervision?part=providers')
      .then(function (data) { setProviders(data.providers || []); })
      .catch(function (err: Error) { setError(err.message); setLoading(false); });
  }, []);

  useEffect(function () {
    if (timer.current) window.clearTimeout(timer.current);
    dirty.current = false;
    loadMap(email, false);
    const poll = window.setInterval(function () { if (!dirty.current) loadMap(email, true); }, 8000);
    return function () { window.clearInterval(poll); };
  }, [email]);

  function flush(who: string) {
    const next = latestRef.current;
    if (!next || saving.current) {
      if (next && dirty.current) timer.current = window.setTimeout(function () { flush(who); }, 400);
      return;
    }
    saving.current = true;
    const query = who ? '?part=map&provider=' + encodeURIComponent(who) : '?part=map';
    api<MapResponse>('/api/reflective-supervision' + query, {
      method: 'PUT',
      body: { baseUpdatedAt: baseRef.current, map: next, provider: who },
    })
      .then(function (data) {
        saving.current = false;
        baseRef.current = data.map.updatedAt || '';
        if (who !== emailRef.current) return;
        if (latestRef.current !== next) {
          timer.current = window.setTimeout(function () { flush(who); }, 200);
          return;
        }
        latestRef.current = data.map;
        setMap(data.map);
        dirty.current = false;
        setNotice('Saved. ' + (data.map.name || 'This provider') + ' and administrators are looking at the same map.');
        setError('');
      })
      .catch(function (err: Error) {
        saving.current = false;
        setError(err.message);
        setNotice('');
      });
  }

  function updateMap(next: GrowthMap) {
    latestRef.current = next;
    setMap(next);
    dirty.current = true;
    setNotice('Saving…');
    if (timer.current) window.clearTimeout(timer.current);
    const who = email;
    timer.current = window.setTimeout(function () { flush(who); }, 900);
  }

  function setTopic(id: string, key: 'e' | 'p' | 'r', value: string) {
    if (!map) return;
    updateMap(Object.assign({}, map, {
      topics: map.topics.map(function (topic) {
        if (topic.id !== id) return topic;
        const copy = Object.assign({}, topic);
        copy[key] = value;
        return copy;
      }),
    }));
  }

  function setMeeting(id: string, key: keyof Meeting, value: string) {
    if (!map) return;
    updateMap(Object.assign({}, map, {
      meetings: map.meetings.map(function (meeting) {
        if (meeting.id !== id) return meeting;
        const copy = Object.assign({}, meeting);
        copy[key] = value;
        return copy;
      }),
    }));
  }

  function addMeeting() {
    if (!map) return;
    updateMap(Object.assign({}, map, {
      meetings: map.meetings.concat([{
        id: 'm-' + Math.random().toString(16).slice(2, 10),
        date: '', update: '', background: '', clip: '', questions: '', noticed: '', action: '', observe: '', revisit: '', growthDates: '',
      }]),
    }));
  }

  async function downloadPdf() {
    setError('');
    const query = email ? '&provider=' + encodeURIComponent(email) : '';
    const res = await fetch('/api/reflective-supervision?part=pdf' + query, { credentials: 'include' });
    if (!res.ok) {
      const data = await res.json().catch(function () { return {}; });
      setError((data && data.error) || 'The PDF could not be made.');
      return;
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = (map && map.name ? map.name : 'Provider') + ' - Reflective Supervision.pdf';
    link.click();
    URL.revokeObjectURL(url);
  }

  function saveCopy() {
    setNotice('Saving the read-only copy to SharePoint…');
    const query = email ? '?part=copy&provider=' + encodeURIComponent(email) : '?part=copy';
    api<{ copy: { ok: boolean; path: string; error: string } }>('/api/reflective-supervision' + query, { method: 'POST', body: { provider: email } })
      .then(function (data) {
        if (data.copy && data.copy.ok) setNotice('Saved a read-only PDF in SharePoint: ' + data.copy.path);
        else setError((data.copy && data.copy.error) || 'The SharePoint copy was not saved.');
      })
      .catch(function (err: Error) { setError(err.message); });
  }

  function sendAppointment(body: Record<string, string>) {
    setError('');
    api<{ appointment: Appointment; mailWarning?: string; calendarWarning?: string }>('/api/reflective-supervision?part=appointments&provider=' + encodeURIComponent(email), {
      method: 'POST',
      body: Object.assign({ provider: email }, body),
    })
      .then(function (data) {
        const warning = [data.mailWarning, data.calendarWarning].filter(Boolean).join(' ');
        setNotice(warning || 'Saved.');
        setSuggesting('');
        loadSide(email);
      })
      .catch(function (err: Error) { setError(err.message); });
  }

  return (
    <section className="rs-map">
      <h2>Growth map</h2>
      <p className="lede">The same map the provider sees. It saves itself as you type. Download a PDF, or save a read-only copy into the RS Providers folder on SharePoint. What is written here is not put on the calendar.</p>
      <div className="row">
        <label>Provider
          <select value={provider} onChange={function (event) { setTyped(''); setProvider(event.target.value); }}>
            <option value="">My map</option>
            {providers.map(function (person) {
              return <option key={person.email} value={person.email}>{person.name}{person.standing ? ' — ' + person.standing.short : ''}</option>;
            })}
          </select>
        </label>
        <label>Or an Emme email that is not in the list
          <input value={typed} placeholder="name@emmeco.org" onChange={function (event) { setTyped(event.target.value); }} />
        </label>
      </div>
      {standing ? <p className="rs-note">{standing.label}</p> : null}
      {notice ? <p className="rs-note">{notice}</p> : null}
      {error ? <p className="error">{error}</p> : null}
      {loading ? <p className="muted">Loading the growth map…</p> : null}
      {map ? (
        <>
          <p className="rs-note">E is what was explored. P is what was practiced. R is what was revisited. An F marks a shared foundation.</p>
          {map.topics.map(function (topic) {
            return (
              <article className="rs-topic" key={topic.id}>
                <h3>{topic.foundation ? '[F] ' : ''}{topic.title}</h3>
                <p>{topic.prompt}</p>
                <div className="rs-epr">
                  <label>E
                    <textarea value={topic.e} onChange={function (event) { setTopic(topic.id, 'e', event.target.value); }} />
                  </label>
                  <label>P
                    <textarea value={topic.p} onChange={function (event) { setTopic(topic.id, 'p', event.target.value); }} />
                  </label>
                  <label>R
                    <textarea value={topic.r} onChange={function (event) { setTopic(topic.id, 'r', event.target.value); }} />
                  </label>
                </div>
              </article>
            );
          })}
          <h3>Preparation and closing notes</h3>
          {map.meetings.map(function (meeting, index) {
            return (
              <article className="rs-topic" key={meeting.id}>
                <h3>Meeting {index + 1}</h3>
                {MEETING_FIELDS.map(function (field) {
                  return (
                    <label key={field.key}>{field.label}
                      {field.kind === 'date'
                        ? <input type="date" value={meeting.date} onChange={function (event) { setMeeting(meeting.id, 'date', event.target.value); }} />
                        : <textarea value={String(meeting[field.key] || '')} onChange={function (event) { setMeeting(meeting.id, field.key, event.target.value); }} />}
                    </label>
                  );
                })}
              </article>
            );
          })}
          <p className="row">
            <button type="button" className="btn btn--ghost" onClick={addMeeting}>Add another meeting</button>
            <button type="button" className="btn btn--pink" onClick={downloadPdf}>Download PDF</button>
            <button type="button" className="btn btn--green" onClick={saveCopy}>Save read-only copy in SharePoint</button>
          </p>

          <h3>When this provider is with clients</h3>
          <p className="rs-note">The next 14 days, from the schedule they save in Schedule &amp; Caseload, and from Outlook. Use an open time for reflective supervision.</p>
          {visits.length === 0 ? <p className="muted">No booked client visits were found on the saved schedule.</p> : (
            <ul className="rs-list">
              {visits.map(function (visit, index) {
                return <li key={visit.date + visit.time + index}>{visit.day}, {visit.date} {visit.time} — {visit.label}</li>;
              })}
            </ul>
          )}
          {outlookError ? <p className="rs-note">Outlook: {outlookError}</p> : null}
          {outlook.length > 0 ? (
            <ul className="rs-list">
              {outlook.map(function (event, index) {
                return <li key={event.start + index}>{event.start.replace('T', ' ')} — {event.subject}</li>;
              })}
            </ul>
          ) : null}

          <h3>Set a reflective supervision time</h3>
          <p className="rs-note">Mountain Time. The provider can accept, decline, or suggest another time. After it is confirmed, it is added to both Outlook calendars. The calendar event does not include the map.</p>
          <div className="row">
            <label>Start
              <input type="datetime-local" value={start} onChange={function (event) { setStart(event.target.value); }} />
            </label>
            <label>End
              <input type="datetime-local" value={end} onChange={function (event) { setEnd(event.target.value); }} />
            </label>
            <button type="button" className="btn btn--green" onClick={function () { sendAppointment({ action: 'propose', start: start, end: end }); }}>Propose this time</button>
          </div>
          {appointments.length === 0 ? <p className="muted">No reflective supervision times yet.</p> : (
            <ul className="rs-list">
              {appointments.map(function (item) {
                const mine = item.providerEmail === me;
                return (
                  <li key={item.id}>
                    <strong>{whenLabel(item.start)}</strong> to {whenLabel(item.end)} — {statusLabel(item.status)}
                    {item.suggestedStart ? <span> Suggested: {whenLabel(item.suggestedStart)}</span> : null}
                    <span className="row">
                      {mine && item.status === 'proposed' ? <button type="button" className="btn btn--green" onClick={function () { sendAppointment({ action: 'accept', id: item.id }); }}>Accept</button> : null}
                      {mine && item.status === 'proposed' ? <button type="button" className="btn btn--ghost" onClick={function () { sendAppointment({ action: 'decline', id: item.id }); }}>Decline</button> : null}
                      {mine && (item.status === 'proposed' || item.status === 'suggested') ? <button type="button" className="btn btn--ghost" onClick={function () { setSuggesting(item.id); }}>Suggest another time</button> : null}
                      {item.status === 'suggested' ? <button type="button" className="btn btn--green" onClick={function () { sendAppointment({ action: 'confirm', id: item.id }); }}>Confirm suggested time</button> : null}
                      {item.status === 'accepted' ? <button type="button" className="btn btn--green" onClick={function () { sendAppointment({ action: 'calendar', id: item.id }); }}>Add to both calendars</button> : null}
                      {item.status !== 'confirmed' && item.status !== 'withdrawn' && item.status !== 'declined' ? <button type="button" className="btn btn--ghost" onClick={function () { sendAppointment({ action: 'withdraw', id: item.id }); }}>Withdraw</button> : null}
                    </span>
                    {suggesting === item.id ? (
                      <span className="row">
                        <input type="datetime-local" value={suggestStart} onChange={function (event) { setSuggestStart(event.target.value); }} />
                        <input type="datetime-local" value={suggestEnd} onChange={function (event) { setSuggestEnd(event.target.value); }} />
                        <button type="button" className="btn btn--pink" onClick={function () { sendAppointment({ action: 'suggest', id: item.id, start: suggestStart, end: suggestEnd }); }}>Send suggestion</button>
                      </span>
                    ) : null}
                  </li>
                );
              })}
            </ul>
          )}
        </>
      ) : null}
    </section>
  );
}
