import { useEffect, useState, type FormEvent } from 'react';
import { api, type Row } from '../api';

function isActive(alert: Row): boolean {
  return alert.Active === true || alert.Active === 'true' || alert.Active === 1;
}

export function AnnouncementPanel() {
  const [alerts, setAlerts] = useState<Row[]>([]);
  const [draft, setDraft] = useState('');
  const [error, setError] = useState('');
  const [status, setStatus] = useState('');

  function load() {
    api<{ alerts: Row[] }>('/api/admin-alerts')
      .then((data) => {
        setAlerts((data.alerts || []).filter(isActive));
        setError('');
      })
      .catch((err: Error) => setError(err.message));
  }

  useEffect(() => { load(); }, []);

  async function post(event: FormEvent) {
    event.preventDefault();
    const message = draft.trim();
    if (!message) return;
    setError('');
    setStatus('Posting…');
    try {
      await api('/api/admin-alerts', { method: 'POST', body: { message } });
      setDraft('');
      setStatus('Posted to the staff banner.');
      load();
      window.dispatchEvent(new Event('emme-alerts-changed'));
    } catch (err) {
      setStatus('');
      setError(err instanceof Error ? err.message : 'Could not post');
    }
  }

  async function remove(id: string) {
    setError('');
    setStatus('');
    try {
      await api('/api/admin-alerts/' + encodeURIComponent(id) + '/end', { method: 'POST', body: {} });
      load();
      window.dispatchEvent(new Event('emme-alerts-changed'));
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not remove');
    }
  }

  return (
    <>
      <h2 id="announcement-title">Announcement</h2>
      <p className="ask-float__hint">Scrolls across the staff dashboard until you remove it.</p>
      <form onSubmit={post}>
        <label htmlFor="announcement-message">
          Message
          <textarea
          id="announcement-message"
          maxLength={240}
          value={draft}
          placeholder="Write a short update"
          required
          onChange={(event) => setDraft(event.target.value)}
        />
        </label>
        <button type="submit" className="btn">Post</button>
      </form>
      {status ? <p className="muted">{status}</p> : null}
      {error ? <p className="error">{error}</p> : null}
      <div className="messages">
        {alerts.length === 0 ? <p className="empty">Nothing on the banner right now.</p> : null}
        {alerts.map((alert) => {
          const id = String(alert.id || '');
          const text = String(alert.Message || alert.Title || '');
          return (
            <article className="message" key={id}>
              <p>{text}</p>
              <button type="button" className="btn btn--ghost" onClick={() => remove(id)}>Remove</button>
            </article>
          );
        })}
      </div>
    </>
  );
}
