import { useCallback, useEffect, useRef, useState } from 'react';
import { ApiError, api } from '../../api';
import { FilePanel } from './FilePanel';
import { PromptPanel } from './PromptPanel';
import { QuestionField } from './QuestionField';
import {
  LEVEL_TEXT,
  hasAnswer,
  type Answer,
  type GroupEntry,
  type Question,
  type Section,
  type SetupState,
  type StoredFile,
} from './types';

type Data = { general: Record<string, Answer>; sources: GroupEntry[] };
type SaveStatus = 'saved' | 'saving' | 'unsaved' | 'error' | 'conflict';

const STATUS_TEXT: Record<SaveStatus, string> = {
  saved: 'All changes saved',
  saving: 'Saving…',
  unsaved: 'Unsaved changes. Saving in a moment…',
  error: 'Could not save. Your changes are kept here and will be retried.',
  conflict: 'Someone else saved changes. Reload the page to see them. Your latest edits are not saved.',
};

function newId(): string {
  const bytes = new Uint8Array(5);
  crypto.getRandomValues(bytes);
  return 'src-' + Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0')).join('');
}

function SectionBlock(props: {
  sections: Section[];
  questions: Question[];
  answers: Record<string, Answer>;
  onChange: (id: string, answer: Answer) => void;
}) {
  return (
    <>
      {props.sections.map((section) => {
        const inSection = props.questions.filter((q) => q.section === section.id);
        if (!inSection.length) return null;
        const done = inSection.filter((q) => hasAnswer(props.answers[q.id])).length;
        return (
          <details key={section.id} className="setup-section">
            <summary>
              <strong>{section.title}</strong>
              <span className="muted"> · {done} of {inSection.length} answered</span>
            </summary>
            {section.intro ? <p className="muted">{section.intro}</p> : null}
            {inSection.map((question) => (
              <QuestionField
                key={question.id}
                question={question}
                answer={props.answers[question.id]}
                onChange={(answer) => props.onChange(question.id, answer)}
              />
            ))}
          </details>
        );
      })}
    </>
  );
}

export function SetupPanel(props: { topicId: string }) {
  const apiBase = '/api/intake/' + encodeURIComponent(props.topicId);
  const [state, setState] = useState<SetupState | null>(null);
  const [data, setData] = useState<Data>({ general: {}, sources: [] });
  const [files, setFiles] = useState<StoredFile[]>([]);
  const [status, setStatus] = useState<SaveStatus>('saved');
  const [error, setError] = useState('');
  const [newKind, setNewKind] = useState('');
  const [newName, setNewName] = useState('');
  const [openId, setOpenId] = useState('');

  const dataRef = useRef<Data>(data);
  const dirty = useRef(false);
  const conflict = useRef(false);
  const base = useRef('');
  const timer = useRef<number | undefined>(undefined);
  const chain = useRef<Promise<boolean>>(Promise.resolve(true));

  const doSave = useCallback(async (): Promise<boolean> => {
    if (conflict.current) return false;
    if (!dirty.current) return true;
    dirty.current = false;
    setStatus('saving');
    const sent = dataRef.current;
    try {
      const res = await api<{ updatedAt: string }>(apiBase + '/state', {
        method: 'PUT',
        body: { baseUpdatedAt: base.current, general: sent.general, sources: sent.sources },
      });
      base.current = res.updatedAt;
      setStatus(dirty.current ? 'unsaved' : 'saved');
      return true;
    } catch (err) {
      dirty.current = true;
      if (err instanceof ApiError && err.status === 409) {
        conflict.current = true;
        setStatus('conflict');
      } else {
        setStatus('error');
        timer.current = window.setTimeout(() => { void save(); }, 8000);
      }
      return false;
    }
  }, [apiBase]);

  const save = useCallback((): Promise<boolean> => {
    window.clearTimeout(timer.current);
    chain.current = chain.current.then(doSave, doSave);
    return chain.current;
  }, [doSave]);

  function edit(change: (current: Data) => Data) {
    const next = change(dataRef.current);
    dataRef.current = next;
    setData(next);
    dirty.current = true;
    if (!conflict.current) setStatus('unsaved');
    window.clearTimeout(timer.current);
    timer.current = window.setTimeout(() => { void save(); }, 1500);
  }

  const reloadFiles = useCallback(() => {
    api<SetupState>(apiBase + '/state')
      .then((fresh) => setFiles(fresh.files))
      .catch((err: Error) => setError(err.message));
  }, [apiBase]);

  useEffect(() => {
    api<SetupState>(apiBase + '/state')
      .then((loaded) => {
        setState(loaded);
        if (loaded.topic.group) setNewKind(loaded.topic.group.kinds[0].id);
        const initial = { general: loaded.answers.general, sources: loaded.answers.sources };
        dataRef.current = initial;
        setData(initial);
        setFiles(loaded.files);
        base.current = loaded.answers.updatedAt;
      })
      .catch((err: Error) => setError(err.message));
    const warn = (event: BeforeUnloadEvent) => {
      if (dirty.current) { event.preventDefault(); event.returnValue = ''; }
    };
    window.addEventListener('beforeunload', warn);
    return () => {
      window.removeEventListener('beforeunload', warn);
      void save();
    };
  }, [save, apiBase]);

  if (error && !state) return <p className="error">{error}</p>;
  if (!state) return <p className="muted">Loading the setup questions…</p>;

  const generalDone = state.generalQuestions.filter((q) => hasAnswer(data.general[q.id])).length;

  const group = state.topic.group;
  const noun = group ? group.singular : '';

  function addSource() {
    const name = newName.trim();
    if (!name) return setError('Type the name of the ' + noun + ' first.');
    setError('');
    const source: GroupEntry = { id: newId(), kind: newKind, name, answers: {} };
    edit((current) => ({ ...current, sources: [...current.sources, source] }));
    setNewName('');
    setOpenId(source.id);
  }

  function removeSource(source: GroupEntry) {
    if (!window.confirm('Remove "' + source.name + '" and every answer and file for it? This cannot be undone.')) return;
    edit((current) => ({ ...current, sources: current.sources.filter((item) => item.id !== source.id) }));
    setFiles((current) => current.filter((file) => file.scope !== source.id));
  }

  function updateSource(id: string, change: (source: GroupEntry) => GroupEntry) {
    edit((current) => ({ ...current, sources: current.sources.map((item) => (item.id === id ? change(item) : item)) }));
  }

  const kindLabel = (id: string) => group?.kinds.find((item) => item.id === id)?.label || id;
  const fileProps = {
    files,
    fileKinds: state.fileKinds,
    maxFileBytes: state.limits.maxFileBytes,
    fileTypes: state.limits.fileTypes,
    apiBase,
    ensureSaved: save,
    onChanged: reloadFiles,
  };

  return (
    <div className="setup">
      <p className="lede">
        Answer what you can. Skip anything you do not know, or tick “I am not sure”. Nothing here connects to anything outside Emme, sends anything, or changes the {state.topic.title} screen.
        Your answers are saved as you type and are shared by every administrator.
      </p>
      <p className={'setup-status setup-status--' + status} role="status">
        {STATUS_TEXT[status]}
        {state.answers.updatedBy && status === 'saved' ? ' · last saved by ' + state.answers.updatedBy : ''}
      </p>
      <p className="setup-warning">
        <strong>Never type a password, a sign-in code, or a key here.</strong> Do not upload real client information. Use fake names and numbers in samples.
      </p>
      {error ? <p className="error">{error}</p> : null}

      <section className="setup-block">
        <h2>{group ? state.topic.title + ' overall' : state.topic.title + ' setup questions'} <span className="muted">· {generalDone} of {state.generalQuestions.length} answered</span></h2>
        <SectionBlock
          sections={state.generalSections}
          questions={state.generalQuestions}
          answers={data.general}
          onChange={(id, answer) => edit((current) => ({ ...current, general: { ...current.general, [id]: answer } }))}
        />
        <FilePanel scope="general" {...fileProps} />
      </section>

      {group ? (
      <section className="setup-block">
        <h2>{group.plural.charAt(0).toUpperCase() + group.plural.slice(1)}</h2>
        <p className="muted">{group.intro}</p>
        {data.sources.map((source) => {
          const relevant = group.questions.filter((q) => !q.kinds || q.kinds.includes(source.kind));
          const done = relevant.filter((q) => hasAnswer(source.answers[q.id])).length;
          return (
            <details key={source.id} className="setup-source" open={openId === source.id ? true : undefined}>
              <summary>
                <strong>{source.name}</strong>
                <span className="muted"> · {kindLabel(source.kind)} · {done} of {relevant.length} answered</span>
              </summary>
              <div className="row">
                <label>
                  Name of this {group.singular}
                  <input
                    value={source.name}
                    maxLength={120}
                    onChange={(event) => updateSource(source.id, (item) => ({ ...item, name: event.target.value }))}
                  />
                </label>
                <button type="button" className="btn btn--ghost" onClick={() => removeSource(source)}>Remove this {group.singular}</button>
              </div>
              <SectionBlock
                sections={group.sections}
                questions={relevant}
                answers={source.answers}
                onChange={(id, answer) => updateSource(source.id, (item) => ({ ...item, answers: { ...item.answers, [id]: answer } }))}
              />
              <FilePanel scope={source.id} {...fileProps} />
            </details>
          );
        })}
        <div className="setup-add">
          <h3>Add a {group.singular}</h3>
          <div className="row">
            <label>
              What kind is it?
              <select value={newKind} onChange={(event) => setNewKind(event.target.value)}>
                {group.kinds.map((kind) => <option key={kind.id} value={kind.id}>{kind.label}</option>)}
              </select>
            </label>
            <label>
              Its name
              <input value={newName} maxLength={120} onChange={(event) => setNewName(event.target.value)} placeholder={group.namePlaceholder} />
            </label>
            <button type="button" className="btn" onClick={addSource}>Add it</button>
          </div>
          <p className="muted">{group.kinds.find((kind) => kind.id === newKind)?.description}</p>
        </div>
      </section>
      ) : null}

      <section className="setup-block">
        <h2>What is possible, and what may need hand work</h2>
        <p className="muted">
          {state.topic.guideIntro}
          {state.topic.hasAssessment ? ' The table at the bottom uses your answers to say which ones can be automated.' : ''}
        </p>
        <div className="table-wrap">
          <table>
            <thead>
              <tr><th>Way to connect</th><th>What it means</th><th>What works</th><th>What can block it</th><th>What we need</th></tr>
            </thead>
            <tbody>
              {state.connectionGuide.map((option) => (
                <tr key={option.id}>
                  <td><strong>{option.title}</strong><br /><span className="muted">{LEVEL_TEXT[option.level]}</span></td>
                  <td>{option.what}</td>
                  <td>{option.works}</td>
                  <td>{option.blocked}</td>
                  <td>{option.needs}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      <PromptPanel apiBase={apiBase} topicId={props.topicId} groupLabel={group ? group.singular : ''} ensureSaved={save} fileCount={files.length} />
    </div>
  );
}
