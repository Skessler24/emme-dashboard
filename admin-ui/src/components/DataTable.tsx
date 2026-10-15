function cell(value: unknown): string {
  if (value == null || value === '') return '—';
  if (typeof value === 'object') return JSON.stringify(value);
  return String(value);
}

export function DataTable(props: {
  columns: Array<{ key: string; label: string }>;
  rows: ReadonlyArray<object>;
  empty: string;
}) {
  if (!props.rows.length) return <p className="empty">{props.empty}</p>;
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>{props.columns.map((column) => <th key={column.key}>{column.label}</th>)}</tr>
        </thead>
        <tbody>
          {props.rows.map((row, index) => {
            const record = row as Record<string, unknown>;
            return (
              <tr key={String(record.id ?? index)}>
                {props.columns.map((column) => <td key={column.key}>{cell(record[column.key])}</td>)}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}
