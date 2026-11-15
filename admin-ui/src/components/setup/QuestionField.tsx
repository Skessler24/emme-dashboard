import { useId } from 'react';
import type { Answer, Question } from './types';

const EMPTY: Answer = { choice: [], text: '', unsure: false };

export function QuestionField(props: {
  question: Question;
  answer: Answer | undefined;
  onChange: (answer: Answer) => void;
}) {
  const { question } = props;
  const answer = props.answer || EMPTY;
  const name = useId();
  const set = (patch: Partial<Answer>) => props.onChange({ ...answer, ...patch });

  function toggle(id: string) {
    if (question.kind === 'one') {
      set({ choice: [id] });
      return;
    }
    set({ choice: answer.choice.includes(id) ? answer.choice.filter((c) => c !== id) : [...answer.choice, id] });
  }

  return (
    <div className="setup-q">
      <p className="setup-q__text">{question.text}</p>
      {question.help ? <p className="muted setup-q__help">{question.help}</p> : null}
      {question.kind !== 'text' ? (
        <div className="setup-q__options" role={question.kind === 'one' ? 'radiogroup' : 'group'}>
          {(question.options || []).map((option) => (
            <label key={option.id} className="setup-check">
              <input
                type={question.kind === 'one' ? 'radio' : 'checkbox'}
                name={name}
                checked={answer.choice.includes(option.id)}
                onChange={() => toggle(option.id)}
              />
              <span>{option.label}</span>
            </label>
          ))}
        </div>
      ) : null}
      <textarea
        aria-label={question.kind === 'text' ? 'Your answer' : 'Add details (optional)'}
        placeholder={question.kind === 'text' ? 'Type your answer here' : 'Add details (optional)'}
        value={answer.text}
        maxLength={6000}
        onChange={(event) => set({ text: event.target.value })}
      />
      <label className="setup-check setup-q__unsure">
        <input type="checkbox" checked={answer.unsure} onChange={(event) => set({ unsure: event.target.checked })} />
        <span>I am not sure. Someone needs to find out.</span>
      </label>
    </div>
  );
}
