import { useEffect, useRef, useState } from 'react';
import { api } from '../api';
import { StaffMessageAdmin } from './StaffMessageAdmin';

export function StaffConnection() {
  const [open, setOpen] = useState(false);
  const [count, setCount] = useState(0);
  const openRef = useRef(false);

  useEffect(() => {
    let stop = false;
    async function tick() {
      if (openRef.current) return;
      try {
        const data = await api<{ count: number }>('/api/staff-messages/unread?box=admin');
        if (!stop && !openRef.current) setCount(data.count || 0);
      } catch {
        if (!stop) setCount(0);
      }
    }
    tick();
    const timer = window.setInterval(tick, 20000);
    return () => {
      stop = true;
      window.clearInterval(timer);
    };
  }, []);

  useEffect(() => {
    if (!open) return;
    function onKey(event: KeyboardEvent) {
      if (event.key === 'Escape') {
        openRef.current = false;
        setOpen(false);
      }
    }
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [open]);

  async function toggle() {
    const next = !open;
    openRef.current = next;
    setOpen(next);
    if (!next) return;
    try {
      await api('/api/staff-messages/seen', { method: 'POST', body: { box: 'admin' } });
      setCount(0);
    } catch {
      setCount(0);
    }
  }

  return (
    <div className="ask-float">
      <div className="ask-float__row">
        <button
          type="button"
          className={count > 0 ? 'ask-float__bubble has-unread' : 'ask-float__bubble'}
          aria-expanded={open}
          aria-controls="staff-connection-panel"
          aria-label={count > 0 ? 'Staff Communication, ' + count + ' new' : 'Staff Communication'}
          onClick={toggle}
        >
          <svg className={open ? 'ask-float__icon is-hidden' : 'ask-float__icon'} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
            <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z" />
          </svg>
          <svg className={open ? 'ask-float__icon' : 'ask-float__icon is-hidden'} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" aria-hidden="true">
            <path d="M18 6 6 18" />
            <path d="m6 6 12 12" />
          </svg>
          {count > 0 ? <span className="ask-float__count">{count > 9 ? '9+' : count}</span> : null}
        </button>
      </div>
      {open ? (
        <div className="ask-float__panel" id="staff-connection-panel" role="dialog" aria-labelledby="staff-connection-title">
          <h2 id="staff-connection-title">Staff Communication</h2>
          <p className="ask-float__hint">Questions from the staff dashboard.</p>
          <StaffMessageAdmin />
        </div>
      ) : null}
    </div>
  );
}
