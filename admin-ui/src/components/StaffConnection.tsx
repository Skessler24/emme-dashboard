import { useEffect, useState } from 'react';
import { AnnouncementPanel } from './AnnouncementPanel';
import { StaffMessageAdmin } from './StaffMessageAdmin';

type Panel = 'staff' | 'announce' | null;

export function StaffConnection() {
  const [open, setOpen] = useState<Panel>(null);

  useEffect(() => {
    if (!open) return;
    function onKey(event: KeyboardEvent) {
      if (event.key === 'Escape') setOpen(null);
    }
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [open]);

  function toggle(panel: Exclude<Panel, null>) {
    setOpen((current) => (current === panel ? null : panel));
  }

  return (
    <div className="ask-float">
      <div className="ask-float__row">
        <button
          type="button"
          className="ask-float__bubble ask-float__bubble--quiet"
          aria-expanded={open === 'announce'}
          aria-controls="announcement-panel"
          aria-label="Open announcements"
          onClick={() => toggle('announce')}
        >
          <svg className={open === 'announce' ? 'ask-float__icon is-hidden' : 'ask-float__icon'} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
            <path d="m3 11 18-5v12L3 14v-3z" />
            <path d="M11.6 16.8a3 3 0 1 1-5.8-1.6" />
          </svg>
          <svg className={open === 'announce' ? 'ask-float__icon' : 'ask-float__icon is-hidden'} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" aria-hidden="true">
            <path d="M18 6 6 18" />
            <path d="m6 6 12 12" />
          </svg>
        </button>
        <button
          type="button"
          className="ask-float__bubble"
          aria-expanded={open === 'staff'}
          aria-controls="staff-connection-panel"
          aria-label="Open Staff Connection"
          onClick={() => toggle('staff')}
        >
          <svg className={open === 'staff' ? 'ask-float__icon is-hidden' : 'ask-float__icon'} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
            <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z" />
          </svg>
          <svg className={open === 'staff' ? 'ask-float__icon' : 'ask-float__icon is-hidden'} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" aria-hidden="true">
            <path d="M18 6 6 18" />
            <path d="m6 6 12 12" />
          </svg>
        </button>
      </div>
      {open === 'announce' ? (
        <div className="ask-float__panel" id="announcement-panel" role="dialog" aria-labelledby="announcement-title">
          <AnnouncementPanel />
        </div>
      ) : null}
      {open === 'staff' ? (
        <div className="ask-float__panel" id="staff-connection-panel" role="dialog" aria-labelledby="staff-connection-title">
          <h2 id="staff-connection-title">Staff Connection</h2>
          <p className="ask-float__hint">Questions from the staff dashboard.</p>
          <StaffMessageAdmin />
        </div>
      ) : null}
    </div>
  );
}
