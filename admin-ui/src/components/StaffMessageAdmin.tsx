import { useEffect, useState } from 'react';
import { api, type Row } from '../api';

export function StaffMessageAdmin() {
  const [messages, setMessages] = useState<Row[]>([]);
  const [error, setError] = useState('');
  const [drafts, setDrafts] = useState<Record<string, string>>({});

  function load() {
    api<{ messages: Row[] }>('/api/staff-messages')
      .then((data) => { setMessages(data.messages || []); setError(''); })
      .catch((err: Error) => setError(err.message));
  }

  useEffect(() => { load(); }, []);

  async function reply(id: string) {
    const text = (drafts[id] || '').trim();
    if (!text) return;
    setError('');
    try {
      await api('/api/staff-messages/' + encodeURIComponent(id) + '/reply', { method: 'POST', body: { reply: text } });
      setDrafts((current) => ({ ...current, [id]: '' }));
      load();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not send reply');
    }
  }

  return (
    <section>
      <h2>Staff questions</h2>
      {error ? <p className="error">{error}</p> : null}
      <div className="messages">
        {messages.length === 0 ? <p className="empty">No messages yet.</p> : null}
        {messages.map((message) => {
          const id = String(message.id || '');
          return (
            <article className="message" key={id}>
              <header>
                <span>{String(message.FromName || message.FromEmail || 'Staff')}</span>
                <span>{String(message.createdAt || message.Status || '')}</span>
              </header>
              <p>{String(message.Message || '')}</p>
              {message.Reply ? <p className="reply"><strong>Reply: </strong>{String(message.Reply)}</p> : (
                <div className="row">
                  <label>
                    Reply
                    <input
                      value={drafts[id] || ''}
                      onChange={(event) => setDrafts((current) => ({ ...current, [id]: event.target.value }))}
                    />
                  </label>
                  <button type="button" className="btn" onClick={() => reply(id)}>Send reply</button>
                </div>
              )}
            </article>
          );
        })}
      </div>
    </section>
  );
}
