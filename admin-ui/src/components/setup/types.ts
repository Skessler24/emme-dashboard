export type Option = { id: string; label: string };

export type Question = {
  id: string;
  section: string;
  text: string;
  help?: string;
  kind: 'text' | 'one' | 'many';
  options?: Option[];
  kinds?: string[];
};

export type Section = { id: string; title: string; intro?: string };

export type Answer = { choice: string[]; text: string; unsure: boolean };

/** One block of a repeating group, such as a funding source or a referral source. */
export type GroupEntry = {
  id: string;
  kind: string;
  name: string;
  answers: Record<string, Answer>;
};

export type StoredFile = {
  id: string;
  scope: string;
  kind: string;
  note: string;
  name: string;
  size: number;
  uploadedBy: string;
  uploadedAt: string;
};

export type ConnectionOption = {
  id: string;
  title: string;
  what: string;
  works: string;
  blocked: string;
  needs: string;
  level: 'automatic' | 'partly' | 'by-hand';
};

export type Level = 'automatic' | 'partly' | 'by-hand' | 'unknown';

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

export type GroupDef = {
  singular: string;
  plural: string;
  intro: string;
  namePlaceholder: string;
  kinds: Array<Option & { description: string }>;
  sections: Section[];
  questions: Question[];
};

export type SetupState = {
  topic: { id: string; title: string; guideIntro: string; group: GroupDef | null; hasAssessment: boolean };
  answers: { updatedAt: string; updatedBy: string; general: Record<string, Answer>; sources: GroupEntry[] };
  files: StoredFile[];
  generalSections: Section[];
  generalQuestions: Question[];
  fileKinds: Option[];
  connectionGuide: ConnectionOption[];
  limits: { maxFileBytes: number; fileTypes: string[] };
};

export const LEVEL_TEXT: Record<Level, string> = {
  automatic: 'Likely automatic',
  partly: 'Partly automatic',
  'by-hand': 'By hand',
  unknown: 'Not enough information yet',
};

export function hasAnswer(answer: Answer | undefined): boolean {
  return !!answer && (answer.choice.length > 0 || answer.text.trim().length > 0);
}
