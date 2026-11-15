/**
 * Shapes shared by every "setup questions" topic (Billing, Onboarding,
 * Referrals, and so on). A topic is plain data: sections, questions, an
 * optional repeating group (for example, one block per funding source), and a
 * guide that explains what the dashboard can and cannot connect to.
 */

export type Option = { id: string; label: string };

export type Question = {
  id: string;
  section: string;
  text: string;
  /** Why we ask, or an example of a good answer. */
  help?: string;
  /** text = free writing. one = pick one. many = pick all that apply. Pick questions also allow notes. */
  kind: 'text' | 'one' | 'many';
  options?: Option[];
  /** Group questions only: limit to these kinds. Empty means all of them. */
  kinds?: string[];
};

export type Section = { id: string; title: string; intro?: string };

export type Level = 'automatic' | 'partly' | 'by-hand' | 'unknown';

export type ConnectionOption = {
  id: string;
  title: string;
  what: string;
  works: string;
  blocked: string;
  needs: string;
  level: 'automatic' | 'partly' | 'by-hand';
};

/** A repeating block of questions, such as one block per funding source. */
export type GroupDef = {
  singular: string;
  plural: string;
  intro: string;
  namePlaceholder: string;
  kinds: Array<Option & { description: string }>;
  sections: Section[];
  questions: Question[];
};

export type GroupItem = {
  id: string;
  kind: string;
  name: string;
  answers: Record<string, { choice: string[]; text: string; unsure: boolean }>;
};

export type AssessmentRow = { key: string; label: string; level: Level; notes: string[] };

export type Assessment = {
  sourceId: string;
  name: string;
  kind: string;
  rows: AssessmentRow[];
  blockers: string[];
  cautions: string[];
  toConfirm: string[];
};

export type TopicDef = {
  id: string;
  /** Short name used in headings, for example "Billing". */
  title: string;
  /** Where its answers and files are kept in storage. Never reuse a prefix. */
  storagePrefix: string;
  /** What the person asking for the build wants, one paragraph each. */
  goal: string[];
  /** Facts about what already exists in this part of the dashboard (project rules bullets). */
  existing: string[];
  /** Extra project rules that belong to this topic only. */
  rules: string[];
  sections: Section[];
  questions: Question[];
  group?: GroupDef;
  /** The question id that asks which group item to do first, if any. */
  firstId?: string;
  firstLabel?: string;
  guideIntro: string;
  guide: ConnectionOption[];
  fileKinds: Option[];
  /** Rows shown in the feasibility table when the topic has a group. */
  assess?: (item: GroupItem) => Assessment;
};
