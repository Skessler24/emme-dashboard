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

export type FundingSource = {
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

export type Assessment = {
  sourceId: string;
  name: string;
  kind: string;
  sending: { level: Level; notes: string[] };
  checking: { level: Level; notes: string[] };
  blockers: string[];
  cautions: string[];
  toConfirm: string[];
};

export type SetupState = {
  answers: { updatedAt: string; updatedBy: string; general: Record<string, Answer>; sources: FundingSource[] };
  files: StoredFile[];
  generalSections: Section[];
  generalQuestions: Question[];
  sourceSections: Section[];
  sourceQuestions: Question[];
  sourceKinds: Array<Option & { description: string }>;
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
