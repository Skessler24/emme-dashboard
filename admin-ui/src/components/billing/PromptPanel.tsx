import { useState } from 'react';
import { api } from '../../api';
import { LEVEL_TEXT, type Assessment } from './types';

type PromptResult = {
  prompt: string;
  assessments?: Assessment[];
  missing: Array<{ where: string; question: string; reason: 'unanswered' | 'not-sure' }>;
};

export function PromptPanel(props: { ensureSaved: () => Promise<boolean>; fileCount: number; base?: string; fileName?: string }) {
  const [result, setResult] = useState<PromptResult | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [copied, setCopied] = useState(false);

  async function create() {
    setBusy(true);
    setError('');
    setCopied(false);
    try {
      if (!(await props.ensureSaved())) throw new Error('Your answers could not be saved, so the prompt was not created.');
      setResult(await api<PromptResult>((props.base || '/api/billing-intake') + '/prompt'));
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not create the prompt');
    } finally {
      setBusy(false);
    }
  }

  async function copy() {
    if (!result) return;
    try {
      await navigator.clipboard.writeText(result.prompt);
      setCopied(true);
    } catch {
      setError('Copying was blocked by the browser. Select the text and copy it yourself.');
    }
  }

  function download() {
    if (!result) return;
    const url = URL.createObjectURL(new Blob([result.prompt], { type: 'text/markdown' }));
    const link = document.createElement('a');
    link.href = url;
    link.download = props.fileName || 'billing-build-prompt.md';
    link.click();
    URL.revokeObjectURL(url);
  }

  const unanswered = result ? result.missing.filter((item) => item.reason === 'unanswered').length : 0;
  const unsure = result ? result.missing.length - unanswered : 0;

  return (
    <section className="setup-block">
      <h2>Turn the answers into a build prompt</h2>
      <p className="muted">
        This puts every answer into one document that tells the builder exactly what to create, what is blocked, and what is still unknown.
        You can create it at any time, even if some questions are blank. The blanks are listed so the builder asks about them instead of guessing.
      </p>
      <div className="row">
        <button type="button" className="btn" disabled={busy} onClick={create}>{busy ? 'Working…' : 'Create the build prompt'}</button>
      </div>
      {error ? <p className="error">{error}</p> : null}
      {result ? (
        <>
          <p>
            <strong>{unanswered}</strong> questions are unanswered and <strong>{unsure}</strong> are marked “not sure”.
            {props.fileCount ? ' ' + props.fileCount + ' uploaded files are listed in the prompt. Download them from the file lists above and attach them to the chat with the prompt.' : ' No files are uploaded yet.'}
          </p>
          {result.assessments && result.assessments.length ? (
            <div className="table-wrap">
              <table>
                <thead>
                  <tr><th>Funding source</th><th>Sending bills</th><th>Finding out about payment</th><th>Blocked or needs a person</th></tr>
                </thead>
                <tbody>
                  {result.assessments.map((item) => (
                    <tr key={item.sourceId}>
                      <td>{item.name}</td>
                      <td>{LEVEL_TEXT[item.sending.level]}</td>
                      <td>{LEVEL_TEXT[item.checking.level]}</td>
                      <td>
                        {[...item.blockers, ...item.cautions].length
                          ? <ul className="setup-list">{[...item.blockers, ...item.cautions].map((text) => <li key={text}>{text}</li>)}</ul>
                          : <span className="muted">Nothing flagged yet</span>}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : null}
          <div className="row">
            <button type="button" className="btn" onClick={copy}>{copied ? 'Copied' : 'Copy the prompt'}</button>
            <button type="button" className="btn btn--ghost" onClick={download}>Download as a file</button>
          </div>
          <textarea className="setup-prompt" readOnly value={result.prompt} aria-label="The build prompt" onFocus={(event) => event.currentTarget.select()} />
        </>
      ) : null}
    </section>
  );
}
