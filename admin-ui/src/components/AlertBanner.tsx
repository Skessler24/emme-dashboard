import { useEffect, useState } from 'react';
import { api, type Row } from '../api';

export function AlertBanner() {
  const [alerts, setAlerts] = useState<Row[]>([]);

  useEffect(() => {
    let cancelled = false;
    function load() {
      api<{ alerts: Row[] }>('/api/admin-alerts?active=true')
        .then((data) => { if (!cancelled) setAlerts(data.alerts || []); })
        .catch(() => { if (!cancelled) setAlerts([]); });
    }
    load();
    window.addEventListener('emme-alerts-changed', load);
    return () => {
      cancelled = true;
      window.removeEventListener('emme-alerts-changed', load);
    };
  }, []);

  if (!alerts.length) return null;
  const line = alerts.map((alert) => String(alert.Message || alert.Title || '')).filter(Boolean).join('   ·   ');
  if (!line) return null;
  return (
    <div className="banner" role="status" aria-live="polite">
      <div className="banner__track">
        <p>{line}</p>
        <p aria-hidden="true">{line}</p>
      </div>
    </div>
  );
}
