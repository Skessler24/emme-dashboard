import { useEffect, useState } from 'react';
import { StaffMessageAdmin } from './StaffMessageAdmin';

export function StaffConnection() {
  const [open, setOpen] = useState(false);

  useEffect(() => {
    if (!open) return;
    function onKey(event: KeyboardEvent) {
      if (event.key === 'Escape') setOpen(false);
    }
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [open]);

  return (
    <div className="ask-float">
      <div className="ask-float__row">
        <button
          type="button"
          className="ask-float__bubble"
          aria-expanded={open}
          aria-controls="staff-connection-panel"
          aria-label="Open Staff Connection"
          onClick={() => setOpen((current) => !current)}
        >
          <svg className={open ? 'ask-float__icon is-hidden' : 'ask-float__icon'} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
            <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z" />
          </svg>
          <svg className={open ? 'ask-float__icon' : 'ask-float__icon is-hidden'} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" aria-hidden="true">
            <path d="M18 6 6 18" />
            <path d="m6 6 12 12" />
          </svg>
        </button>
      </div>
      {open ? (
        <div className="ask-float__panel" id="staff-connection-panel" role="dialog" aria-labelledby="staff-connection-title">
          <h2 id="staff-connection-title">Staff Connection</h2>
          <p className="ask-float__hint">Questions from the staff dashboard.</p>
          <StaffMessageAdmin />
        </div>
      ) : null}
    </div>
  );
}
