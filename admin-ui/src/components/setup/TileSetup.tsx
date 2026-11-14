import { useCallback, useEffect, useRef, useState, type ReactNode } from 'react';
import { ApiError, api } from '../../api';
import { FilePanel } from '../billing/FilePanel';
import { PromptPanel } from '../billing/PromptPanel';
import { QuestionField } from '../billing/QuestionField';
import { hasAnswer, type Answer, type Option, type Question, type Section, type StoredFile } from '../billing/types';

type SetupState = {
  tile: { id: string; title: string };
  answers: { updatedAt: string; updatedBy: string; answers: Record<string, Answer>; firstDraft: boolean };
  files: StoredFile[];
  sections: Section[];
  questions: Question[];
  fileKinds: Option[];
  limits: { maxFileBytes: number; fileTypes: string[] };
};

type SaveStatus = 'saved' | 'saving' | 'unsaved' | 'error' | 'conflict';

const STATUS_TEXT: Record<SaveStatus, string> = {
  saved: 'All changes saved',
  saving: 'Saving…',
  unsaved: 'Unsaved changes. Saving in a moment…',
  error: 'Could not save. Your changes are kept here and will be retried.',
  conflict: 'Someone else saved changes. Reload the page to see them. Your latest edits are not saved.',
};

/**
 * The "setup questions" page for one tile. The answers start as a first draft
 * from the requirements documents, are saved as people type, and are shared by
 * every administrator.
 */
export function TileSetup(props: { tileId: string }) {
  const base = '/api/tile-setup/' + props.tileId;
  const [state, setState] = useState<SetupState | null>(null);
  const [answers, setAnswers] = useState<Record<string, Answer>>({});
  const [files, setFiles] = useState<StoredFile[]>([]);
  const [status, setStatus] = useState<SaveStatus>('saved');
  const [draft, setDraft] = useState(false);
  const [error, setError] = useState('');

  const answersRef = useRef<Record<string, Answer>>({});
  const dirty = useRef(false);
  const conflict = useRef(false);
  const baseStamp = useRef('');
  const timer = useRef<number | undefined>(undefined);
  const chain = useRef<Promise<boolean>>(Promise.resolve(true));

  const doSave = useCallback(async (): Promise<boolean> => {
    if (conflict.current) return false;
    if (!dirty.current) return true;
    dirty.current = false;
    setStatus('saving');
    try {
      const res = await api<{ updatedAt: string }>(base, {
        method: 'PUT',
        body: { baseUpdatedAt: baseStamp.current, answers: answersRef.current },
      });
      baseStamp.current = res.updatedAt;
      setDraft(false);
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
  }, [base]);

  const save = useCallback((): Promise<boolean> => {
    window.clearTimeout(timer.current);
    chain.current = chain.current.then(doSave, doSave);
    return chain.current;
  }, [doSave]);

  function edit(id: string, answer: Answer) {
    const next = { ...answersRef.current, [id]: answer };
    answersRef.current = next;
    setAnswers(next);
    dirty.current = true;
    if (!conflict.current) setStatus('unsaved');
    window.clearTimeout(timer.current);
    timer.current = window.setTimeout(() => { void save(); }, 1500);
  }

  const reloadFiles = useCallback(() => {
    api<SetupState>(base)
      .then((fresh) => setFiles(fresh.files))
      .catch((err: Error) => setError(err.message));
  }, [base]);

  useEffect(() => {
    api<SetupState>(base)
      .then((loaded) => {
        setState(loaded);
        answersRef.current = loaded.answers.answers;
        setAnswers(loaded.answers.answers);
        setFiles(loaded.files);
        setDraft(loaded.answers.firstDraft);
        baseStamp.current = loaded.answers.updatedAt;
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
  }, [base, save]);

  if (error && !state) return <p className="error">{error}</p>;
  if (!state) return <p className="muted">Loading the {props.tileId.replace(/-/g, ' ')} setup…</p>;

  const total = state.questions.length;
  const done = state.questions.filter((q) => hasAnswer(answers[q.id])).length;

  return (
    <div className="setup">
      <p className="lede">
        Check what is here, fix anything that is wrong, and fill in what is blank. Skip anything you do not know, or tick “I am not sure”.
        Nothing here sends a message, changes a file, or connects to another system. Your answers are saved as you type and are shared by every administrator.
      </p>
      {draft ? (
        <p className="setup-note" role="note">
          <strong>This is a first draft.</strong> It was filled in from the Emme Admin Dashboard document and the boss interview sheet.
          Please read it with your bosses. Anything the documents only hinted at is marked “I am not sure”. Nothing is saved until you change something.
        </p>
      ) : null}
      <p className={'setup-status setup-status--' + status} role="status">
        {STATUS_TEXT[status]}
        {state.answers.updatedBy && status === 'saved' && !draft ? ' · last saved by ' + state.answers.updatedBy : ''}
      </p>
      <p className="setup-warning">
        <strong>Never type a password, a sign-in code, or a key here.</strong> Do not upload real client or family information. Use fake names and numbers in samples.
      </p>
      {error ? <p className="error">{error}</p> : null}

      <section className="setup-block">
        <h2>{state.tile.title} <span className="muted">· {done} of {total} answered</span></h2>
        {state.sections.map((section) => {
          const inSection = state.questions.filter((q) => q.section === section.id);
          if (!inSection.length) return null;
          const sectionDone = inSection.filter((q) => hasAnswer(answers[q.id])).length;
          return (
            <details key={section.id} className="setup-section">
              <summary>
                <strong>{section.title}</strong>
                <span className="muted"> · {sectionDone} of {inSection.length} answered</span>
              </summary>
              {section.intro ? <p className="muted">{section.intro}</p> : null}
              {inSection.map((question) => (
                <QuestionField
                  key={question.id}
                  question={question}
                  answer={answers[question.id]}
                  onChange={(answer) => edit(question.id, answer)}
                />
              ))}
            </details>
          );
        })}
        <FilePanel
          scope="general"
          base={base}
          files={files}
          fileKinds={state.fileKinds}
          maxFileBytes={state.limits.maxFileBytes}
          fileTypes={state.limits.fileTypes}
          ensureSaved={save}
          onChanged={reloadFiles}
        />
      </section>

      <PromptPanel
        ensureSaved={save}
        fileCount={files.length}
        base={base}
        fileName={props.tileId + '-build-prompt.md'}
      />
    </div>
  );
}

/** A tile page with two tabs: the tile itself and its setup questions. */
export function TileWithSetup(props: { hashId: string; tileId: string; label: string; children: ReactNode }) {
  const onSetup = window.location.hash.replace(/\/+$/, '').endsWith('/' + props.hashId + '/setup');
  return (
    <section>
      <nav className="subtabs" aria-label={props.label + ' sections'}>
        <a className={onSetup ? 'subtab' : 'subtab subtab--active'} href={'#/' + props.hashId}>Overview</a>
        <a className={onSetup ? 'subtab subtab--active' : 'subtab'} href={'#/' + props.hashId + '/setup'}>Setup questions</a>
      </nav>
      {onSetup ? <TileSetup tileId={props.tileId} /> : props.children}
    </section>
  );
}
