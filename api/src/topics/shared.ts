import { ConnectionOption, Option, Question } from '../intakeTypes';

/** File kinds every topic offers. Topics add their own on top of these. */
export const COMMON_FILE_KINDS: Option[] = [
  { id: 'screenshot', label: 'A screenshot of a screen you use (client and staff details covered)' },
  { id: 'rules-document', label: 'A policy, handbook page, contract, or rules document' },
  { id: 'other', label: 'Something else that explains the process' },
];

/** The last three questions of every topic. The build prompt reads g_success and g_deadline by these ids. */
export function closingQuestions(section: string, topic: string, handExample: string): Question[] {
  return [
    { id: 'g_hand_work', section, kind: 'text', text: 'Which ' + topic + ' tasks do you believe must stay done by a person? (For example: ' + handExample + ') Why?' },
    { id: 'g_success', section, kind: 'text', text: 'How will you know this project worked? Give a goal we can measure. (For example: "We save five hours a week" or "Nothing is missed for more than one day.")' },
    { id: 'g_deadline', section, kind: 'text', text: 'Is there a date or event this needs to be ready for (a busy season, an audit, a staff change)?' },
  ];
}

export function privacyQuestions(section: string, topic: string, who: string): Question[] {
  return [
    { id: 'g_privacy_agreement', section, kind: 'one', text: 'Do you have a signed privacy agreement with Microsoft and with IntakeQ? (Under the federal Health Insurance Portability and Accountability Act this is called a business associate agreement.)', options: [
      { id: 'both', label: 'Yes, with both' },
      { id: 'one', label: 'With only one of them (say which below)' },
      { id: 'no', label: 'No' },
      { id: 'not-sure', label: 'I am not sure' },
    ] },
    { id: 'g_who_sees', section, kind: 'text', text: 'Who is allowed to see ' + who + '? Should every administrator see everything, or only certain people?' },
    { id: 'g_retention', section, kind: 'text', text: 'How long must ' + topic + ' records be kept, and who decides when they are removed?' },
  ];
}

export function reminderQuestions(section: string, what: string): Question[] {
  return [
    { id: 'g_frequency', section, kind: 'one', text: 'How often should the dashboard look for changes to ' + what + '?', options: [
      { id: 'hourly', label: 'Every hour' },
      { id: 'daily', label: 'Once a day, each morning' },
      { id: 'twice-weekly', label: 'Twice a week' },
      { id: 'weekly', label: 'Once a week' },
      { id: 'not-sure', label: 'I am not sure' },
    ] },
    { id: 'g_alert_how', section, kind: 'many', text: 'When something is late or needs attention, how should people be told? (Pick all that apply.)', options: [
      { id: 'dashboard', label: 'On the dashboard screen' },
      { id: 'email', label: 'By email' },
      { id: 'teams', label: 'By a Microsoft Teams message' },
      { id: 'none', label: 'No notice, just show it on the screen' },
    ], help: 'Say below who should be told, and when.' },
    { id: 'g_approval', section, kind: 'text', text: 'Which actions must always be approved by a named person before they happen? Who are the approvers?' },
  ];
}

/** Connection options that apply to every topic. */
export const SHARED_GUIDE: Record<string, ConnectionOption> = {
  sharepoint: {
    id: 'sharepoint', level: 'automatic',
    title: 'SharePoint lists and folders (already connected)',
    what: 'The dashboard already reads and writes the Emme team site through Microsoft.',
    works: 'Lists, document folders, and files on the team site can be read, shown, and updated. This is the easiest place to keep new information.',
    blocked: 'The fixed folder locations in the dashboard are not changed without asking. Anything stored elsewhere (a personal OneDrive, paper) is out of reach until it is moved.',
    needs: 'The name of the list or folder, and who should be allowed to see it.',
  },
  intakeq: {
    id: 'intakeq', level: 'automatic',
    title: 'IntakeQ (already connected for reading)',
    what: 'The dashboard already reads appointments, notes, evaluations, and invoices from IntakeQ.',
    works: 'Anything the dashboard already reads from IntakeQ can be shown, counted, and used to raise reminders.',
    blocked: 'Today it only reads. Whether IntakeQ allows the dashboard to change records (for example, send a form or mark something complete) must be confirmed before we promise that.',
    needs: 'Confirmation of what IntakeQ allows.',
  },
  schedule: {
    id: 'schedule', level: 'automatic',
    title: 'How "keeps checking" and reminders work',
    what: 'Checking and reminders must run on a schedule, not only when someone opens the page.',
    works: 'The dashboard hosting only answers requests and cannot run a clock of its own. The project already uses a Power Automate scheduled flow that calls the dashboard on a timer, and we would do the same here (for example every morning).',
    blocked: 'It is checking on a schedule, not a continuous live link. Very frequent checks may need a separate scheduled service.',
    needs: 'Your answer to how often the checking must run and who should be told.',
  },
  email: {
    id: 'email', level: 'partly',
    title: 'Email through a shared Emme mailbox',
    what: 'The dashboard can send reminders from, or read messages in, a shared Emme mailbox.',
    works: 'Reminders and notices can go out by email. Messages that arrive in a shared mailbox can be spotted, and attachments in spreadsheet or text form can be read.',
    blocked: 'Scanned pictures and some saved documents cannot be read reliably. Reading or sending email needs extra permission in Microsoft Entra that someone with the right access must approve.',
    needs: 'The mailbox address (an address only, never a password) and a sample of each message.',
  },
  signature: {
    id: 'signature', level: 'partly',
    title: 'Electronic signatures and forms',
    what: 'Documents or forms that people must sign or fill in online.',
    works: 'If Emme already uses a signing or forms tool, the dashboard can link to it and show who has not finished. IntakeQ forms are one option.',
    blocked: 'Sending a form on someone\'s behalf from the dashboard needs the tool to allow it. A new signing tool means a new cost and agreement.',
    needs: 'Which tool is used today and what it allows.',
  },
  paper: {
    id: 'paper-phone', level: 'by-hand',
    title: 'Paper, phone calls, and in-person steps',
    what: 'Steps that happen on paper, by phone, or face to face.',
    works: 'The dashboard can keep a to-do list with an owner, a due date, and reminders so nothing is forgotten. Staff record the result by hand.',
    blocked: 'No software can read a phone call or a paper form on its own. These steps stay by hand.',
    needs: 'The steps and the people who do them.',
  },
  other: {
    id: 'outside-system', level: 'partly',
    title: 'A website or program Emme does not use yet',
    what: 'Any outside website (a licensing board, a payroll company, a state system).',
    works: 'Sometimes it has an official connection or a download that a person can drop into the dashboard.',
    blocked: 'Many websites have no official connection. Signing in automatically is fragile, may be forbidden by the website\'s rules, and often cannot work when a second sign-in step is required.',
    needs: 'The website name and whether it offers an official connection or a download.',
  },
};
