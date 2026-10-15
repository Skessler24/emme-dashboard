import { useEffect, useState } from 'react';
import { api, type Row } from '../../api';
import { DataTable } from '../DataTable';

export function BillingTile() {
  const [status, setStatus] = useState('');
  const [rows, setRows] = useState<Row[]>([]);
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    const query = status ? '?status=' + encodeURIComponent(status) : '';
    api<{ invoices: Row[] }>('/api/billing/invoices' + query)
      .then((data) => { if (!cancelled) { setRows(data.invoices || []); setError(''); } })
      .catch((err: Error) => { if (!cancelled) setError(err.message); })
      .finally(() => { if (!cancelled) setLoading(false); });
    return () => { cancelled = true; };
  }, [status]);

  return (
    <section>
      <h1>Billing</h1>
      <p className="lede">Invoice status is read live from IntakeQ and is not stored in the Emme scheduling database.</p>
      <div className="row">
        <label>
          Status
          <select value={status} onChange={(event) => setStatus(event.target.value)}>
            <option value="">All</option>
            <option>Unpaid</option>
            <option>Paid</option>
            <option>PastDue</option>
            <option>Draft</option>
            <option>Refunded</option>
            <option>Canceled</option>
          </select>
        </label>
      </div>
      {error ? <p className="error">{error}</p> : null}
      {loading ? <p className="muted">Loading invoices…</p> : (
        <DataTable
          rows={rows}
          empty="No invoices in this window."
          columns={[
            { key: 'number', label: 'Invoice' },
            { key: 'status', label: 'Status' },
            { key: 'clientName', label: 'Client' },
            { key: 'amountDue', label: 'Due' },
            { key: 'amount', label: 'Total' },
            { key: 'date', label: 'Date' },
            { key: 'practitionerName', label: 'Practitioner' },
          ]}
        />
      )}
    </section>
  );
}
