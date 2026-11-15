import { Question, Section, TopicDef } from '../intakeTypes';
import { COMMON_FILE_KINDS, SHARED_GUIDE, closingQuestions, privacyQuestions, reminderQuestions } from './shared';

const SECTIONS: Section[] = [
  { id: 'what', title: '1. What reflective supervision is at Emme' },
  { id: 'who', title: '2. Who takes part' },
  { id: 'schedule', title: '3. How often and how it is scheduled' },
  { id: 'process', title: '4. The observation process, step by step', intro: 'The dashboard already lists the files in the "RS_Observation Process" SharePoint folder. These questions explain what those files are for.' },
  { id: 'records', title: '5. What is written down and where it is kept' },
  { id: 'hours', title: '6. Tracking completed hours' },
  { id: 'reports', title: '7. Reports and requirements from outside' },
  { id: 'live', title: '8. Reminders and who is told' },
  { id: 'privacy', title: '9. Privacy and who can see what' },
  { id: 'hand', title: '10. What may have to stay by hand, and what success looks like' },
];

const QUESTIONS: Question[] = [
  { id: 'g_what', section: 'what', kind: 'text', text: 'In your own words, what is reflective supervision at Emme and why does Emme do it? Is it different from ordinary supervision?' },
  { id: 'g_required', section: 'what', kind: 'many', text: 'Who requires it? (Pick all that apply.)', options: [
    { id: 'emme', label: 'Emme chooses to do it' },
    { id: 'funding', label: 'A funding source requires it' },
    { id: 'state', label: 'The state requires it' },
    { id: 'license', label: 'A license or certification requires it' },
    { id: 'not-sure', label: 'I am not sure' },
  ], help: 'Say below which requirement, and where it is written down.' },

  { id: 'g_who_gets', section: 'who', kind: 'text', text: 'Who receives reflective supervision, and who gives it? List by role (and first name where it helps). Is it one to one or in a group?' },
  { id: 'g_who_observed', section: 'who', kind: 'text', text: 'The folder is called an observation process. Who observes whom, and what is being observed?' },
  { id: 'g_who_new', section: 'who', kind: 'text', text: 'Is it different for new staff, students, or contractors? Explain.' },

  { id: 'g_frequency_rule', section: 'schedule', kind: 'text', text: 'How often must each person have reflective supervision, and how long is each session?' },
  { id: 'g_book', section: 'schedule', kind: 'text', text: 'How are sessions scheduled today? Who books them, with what calendar or program, and what happens when one is missed or moved?' },
  { id: 'g_busy', section: 'schedule', kind: 'text', text: 'Are there times of year when this slips? What gets in the way?' },

  { id: 'g_steps', section: 'process', kind: 'text', text: 'Walk me through one full observation process, one step at a time, from before it happens to after it is finished. Say who does each step and which form or file is used.' },
  { id: 'g_files', section: 'process', kind: 'text', text: 'What is each file in the "RS_Observation Process" folder for, and who fills it in? Upload a blank copy of each form in the files area.' },
  { id: 'g_follow', section: 'process', kind: 'text', text: 'What happens after a session? (For example: goals are set, the supervisor writes a summary, the staff member signs, a follow-up date is chosen.)' },
  { id: 'g_done', section: 'process', kind: 'text', text: 'How do you know a person\'s observation process is finished? Who says so, and where is that recorded?' },

  { id: 'g_record_what', section: 'records', kind: 'text', text: 'What is written down for each session, who writes it, and how soon afterward?' },
  { id: 'g_record_where', section: 'records', kind: 'text', text: 'Where are these records kept today? Is it always the same folder, and who can open them?' },
  { id: 'g_record_format', section: 'records', kind: 'many', text: 'In what form are the records kept? (Pick all that apply.)', options: [
    { id: 'word', label: 'Word documents' },
    { id: 'excel', label: 'Spreadsheets' },
    { id: 'pdf', label: 'Saved document files or scanned pages' },
    { id: 'forms', label: 'Forms in IntakeQ' },
    { id: 'paper', label: 'Paper' },
    { id: 'other', label: 'Something else (describe below)' },
  ] },
  { id: 'g_record_nav', section: 'records', kind: 'text', text: 'The dashboard lets administrators click through the folders. Is that how you want to find things, or would you rather see a list by person with their status?' },

  { id: 'g_hours_track', section: 'hours', kind: 'text', text: 'How are completed hours tracked today (a spreadsheet, a form, a folder)? Who keeps the count?' },
  { id: 'g_hours_goal', section: 'hours', kind: 'text', text: 'What total does each person need to reach, over what period (a month, a year), and what happens if they fall behind?' },
  { id: 'g_hours_view', section: 'hours', kind: 'text', text: 'What should the dashboard show about hours? (For example: hours completed against hours needed, who is behind, who is finished.)' },

  { id: 'g_report_who', section: 'reports', kind: 'text', text: 'Does anyone outside Emme ask to see reflective supervision records or hours (a funding source, an auditor, a licensing board)? Who, how often, and in what form?' },
  { id: 'g_report_proof', section: 'reports', kind: 'text', text: 'What proof of reflective supervision do you have to produce, and how long does it take to gather today?' },

  ...reminderQuestions('live', 'reflective supervision'),
  ...privacyQuestions('privacy', 'reflective supervision', 'reflective supervision notes, which may include very personal reflections'),
  { id: 'g_who_sees_notes', section: 'privacy', kind: 'text', text: 'Should the content of a session ever be visible to administrators, or only the fact that it happened and its hours? Who is promised confidentiality, and by whom?' },
  ...closingQuestions('hand', 'reflective supervision', 'the conversation itself, a signature in person, notes that must stay private.'),
];

export const reflectiveSupervision: TopicDef = {
  id: 'reflective-supervision',
  title: 'Reflective Supervision',
  storagePrefix: 'intake/reflective-supervision/',
  goal: [
    'Build the Reflective Supervision area of the Administration Dashboard so that leaders can see, for each person, whether their sessions and observation steps are on schedule and how many hours are complete, without anyone chasing forms. The dashboard should remind the right person ahead of time and protect the private content of sessions.',
  ],
  existing: [
    'The Reflective Supervision screen already lists the files and folders in the "RS_Observation Process" SharePoint folder and lets administrators click through them (`admin-ui/src/components/tiles/ReflectiveSupervisionTile.tsx`). Its web address carries the folder path, so any new tab must not use a folder name. Extend the screen rather than replacing it.',
  ],
  rules: [
    'Reflective supervision is a private, reflective conversation. Do not show or log the content of a session unless I clearly say who may see it in my answers below. Showing that it happened and the hours is a different, safer thing.',
    'Anything that cannot be automated must still appear as a to-do item with an owner, a due date, and a reminder, so nothing is lost.',
  ],
  sections: SECTIONS,
  questions: QUESTIONS,
  guideIntro: 'Plain explanations of what the dashboard can connect to for reflective supervision, and where each one stops working.',
  guide: [
    SHARED_GUIDE.sharepoint,
    {
      id: 'calendar', level: 'partly',
      title: 'Outlook calendar',
      what: 'Sessions booked on the Outlook calendar.',
      works: 'If sessions are booked on a shared calendar, the dashboard could show which are coming and which were missed.',
      blocked: 'Reading calendars needs extra permission in Microsoft Entra that someone with the right access must approve. It cannot tell whether a session really happened, only that it was booked.',
      needs: 'Whether sessions are on a calendar and whose.',
    },
    SHARED_GUIDE.signature,
    SHARED_GUIDE.intakeq,
    SHARED_GUIDE.schedule,
    SHARED_GUIDE.email,
    SHARED_GUIDE.paper,
  ],
  fileKinds: [
    { id: 'observation-form', label: 'A blank observation form or template from the folder' },
    { id: 'hours-sheet', label: 'The sheet used to count hours (copy with fake names)' },
    { id: 'sample-summary', label: 'A sample session summary (all private details removed)' },
  ].concat(COMMON_FILE_KINDS),
};
