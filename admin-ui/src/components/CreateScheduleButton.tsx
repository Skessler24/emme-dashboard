import { useState } from 'react';
import { api } from '../api';

type Result = { ok: boolean; created: boolean; fileName?: string; path?: string; message: string };

/**
 * "Create schedule from template". Adds the provider's schedule workbook to the Provider Schedules folder
 * from the standard template. It only adds a missing file and never replaces an existing one.
 * Any tile that knows a provider's email (onboarding, referrals, and so on) can use it.
 */
export function CreateScheduleButton(props: { email: string; name?: string; onDone?: (result: Result) => void }) {
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState('');
  const [failed, setFailed] = useState(false);

  async function create() {
    const email = props.email.trim().toLowerCase();
    if (!/@emmeco\.org$/.test(email)) {
      setFailed(true);
      setMessage('Enter the provider’s emmeco.org email first.');
      return;
    }
    setBusy(true);
    setMessage('');
    setFailed(false);
    try {
      const result = await api<Result>('/api/onboarding/' + encodeURIComponent(email) + '/schedule-template', {
        method: 'POST',
        body: { employeeName: (props.name || '').trim() },
      });
      setMessage(result.message);
      props.onDone?.(result);
    } catch (err) {
      setFailed(true);
      setMessage(err instanceof Error ? err.message : 'Could not create the schedule file');
    } finally {
      setBusy(false);
    }
  }

  return (
    <span>
      <button
        className="btn btn--ghost"
        type="button"
        disabled={busy}
        onClick={create}
        title="Adds this provider's schedule spreadsheet from the standard template. It never replaces an existing file."
      >
        {busy ? 'Creating…' : 'Create schedule from template'}
      </button>
      {message ? <p className={failed ? 'error' : 'muted'} role="status">{message}</p> : null}
    </span>
  );
}
