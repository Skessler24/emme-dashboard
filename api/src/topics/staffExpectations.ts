import { Question, Section, TopicDef } from '../intakeTypes';
import { COMMON_FILE_KINDS, SHARED_GUIDE, closingQuestions, privacyQuestions, reminderQuestions } from './shared';

const SECTIONS: Section[] = [
  { id: 'people', title: '1. People and time' },
  { id: 'expect', title: '2. What Emme expects of each provider', intro: 'The rules as they are today, even if they are not written down anywhere.' },
  { id: 'openings', title: '3. Open times and schedule updates' },
  { id: 'mileage', title: '4. Mileage' },
  { id: 'cancel', title: '5. Cancellations' },
  { id: 'docs', title: '6. Documentation' },
  { id: 'view', title: '7. The weekly view' },
  { id: 'live', title: '8. Reminders and who is told' },
  { id: 'privacy', title: '9. Privacy and who can see what' },
  { id: 'hand', title: '10. What may have to stay by hand, and what success looks like' },
];

const QUESTIONS: Question[] = [
  { id: 'g_who_manages', section: 'people', kind: 'text', text: 'Who checks that providers are meeting expectations today? List each person by first name and job title, and say what each one checks.' },
  { id: 'g_hours', section: 'people', kind: 'text', text: 'About how many hours each week does that checking take, and how much of it is collecting information from several places?' },

  { id: 'g_expect_list', section: 'expect', kind: 'text', text: 'List every expectation Emme has of a provider, one per line. (For example: hours worked each week, number of clients, notes finished within two days, mileage submitted by the fifth of the month, updating open times.)' },
  { id: 'g_expect_differ', section: 'expect', kind: 'text', text: 'Do expectations differ by role, by employee or contractor, or by funding source? Explain each difference.' },
  { id: 'g_expect_where', section: 'expect', kind: 'text', text: 'Where are these expectations written down today (a handbook, an email, a contract), and who is allowed to change them?' },
  { id: 'g_expect_missed', section: 'expect', kind: 'text', text: 'What happens when an expectation is not met? Describe the steps, who is involved, and when it goes from a reminder to a conversation.' },

  { id: 'g_open_how', section: 'openings', kind: 'text', text: 'The dashboard already finds open times from each provider\'s schedule workbook. How do providers keep that workbook up to date, and how often does it go out of date?' },
  { id: 'g_open_use', section: 'openings', kind: 'text', text: 'Who uses the list of open times and for what? (For example: matching new referrals, filling a cancellation, planning the week.)' },
  { id: 'g_open_rules', section: 'openings', kind: 'text', text: 'What are the rules for open times? (For example: a provider must offer at least a certain number of hours, must tell the office before blocking time off, must update by a set day.)' },
  { id: 'g_open_late', section: 'openings', kind: 'text', text: 'When a provider\'s schedule is out of date or missing, what should happen and who should be told?' },

  { id: 'g_mile_rate', section: 'mileage', kind: 'text', text: 'How is mileage paid? Give the rate per mile, what starts and ends the count (home, office, or the first client), and any limits.' },
  { id: 'g_mile_submit', section: 'mileage', kind: 'text', text: 'The dashboard already collects mileage in a nightly sync. How does a provider submit mileage today, how often, and what makes a submission complete?' },
  { id: 'g_mile_check', section: 'mileage', kind: 'text', text: 'Who checks mileage, what are they checking for (for example, distances that do not match the visit addresses), and what happens when something looks wrong?' },
  { id: 'g_mile_pay', section: 'mileage', kind: 'text', text: 'How does mileage get to payroll? Which program or spreadsheet, who sends it, and by what date?' },

  { id: 'g_cancel_policy', section: 'cancel', kind: 'text', text: 'What is Emme\'s cancellation policy for families? Include how much notice is required, any fee, and who can make an exception.' },
  { id: 'g_cancel_provider', section: 'cancel', kind: 'text', text: 'What is the rule when a provider cancels? How much notice, who is told, and who finds a replacement or a make-up time?' },
  { id: 'g_cancel_source', section: 'cancel', kind: 'text', text: 'The dashboard already shows a cancellations list that Power Automate fills from IntakeQ. Is that list right? What is missing, and what do you do with it today?' },
  { id: 'g_cancel_billing', section: 'cancel', kind: 'text', text: 'How do cancellations change billing or provider pay? (For example: a late cancellation is billed, a no-show is paid at half rate.)' },
  { id: 'g_cancel_track', section: 'cancel', kind: 'text', text: 'Which cancellation patterns need a person to step in? (For example: a family who cancels three times in a row, a provider with many cancellations.)' },

  { id: 'g_docs_deadline', section: 'docs', kind: 'text', text: 'How soon after a visit must the note be finished and signed? Does that differ by funding source or kind of visit?' },
  { id: 'g_docs_count', section: 'docs', kind: 'text', text: 'The dashboard already shows weekly documentation counts. What exactly is counted as "done" and "late" today? Does that match your rule?' },
  { id: 'g_docs_review', section: 'docs', kind: 'text', text: 'Does anyone review notes for quality? Who, how many, how often, and what do they look for?' },
  { id: 'g_docs_follow', section: 'docs', kind: 'text', text: 'When notes are late, what happens next, step by step? Who reminds the provider, who is told if it continues, and what are the consequences?' },

  { id: 'g_view_items', section: 'view', kind: 'many', text: 'What should the weekly view show? (Pick all that apply.)', options: [
    { id: 'openings', label: 'Open times by provider' },
    { id: 'cancellations', label: 'Cancellations this week' },
    { id: 'documentation', label: 'Late and finished notes by provider' },
    { id: 'mileage', label: 'Mileage submitted and missing' },
    { id: 'caseload', label: 'Number of clients and visits by provider' },
    { id: 'hours', label: 'Hours worked compared with the expected hours' },
    { id: 'trend', label: 'Change compared with last week' },
    { id: 'attention', label: 'A list of providers who need a conversation' },
  ], help: 'Add anything else you want to see below.' },
  { id: 'g_view_who', section: 'view', kind: 'text', text: 'Who looks at the weekly view, on which day, and what do they do with it?' },
  { id: 'g_view_fair', section: 'view', kind: 'text', text: 'Are there fairness concerns about what is measured? (For example: a new provider, part-time staff, a provider with harder cases.) How should the view show those?' },

  ...reminderQuestions('live', 'provider expectations'),
  ...privacyQuestions('privacy', 'provider performance', 'how each provider is doing against expectations'),
  ...closingQuestions('hand', 'provider expectations', 'a hard conversation, an exception for a family emergency, signing off a mileage claim.'),
];

export const staffExpectations: TopicDef = {
  id: 'expectations',
  title: 'Staff Expectations',
  storagePrefix: 'intake/expectations/',
  goal: [
    'Build the Staff Expectations area of the Administration Dashboard so that leaders can see each week whether providers are meeting Emme\'s expectations: open times kept up to date, mileage submitted, cancellations handled, and notes finished on time. The dashboard should show clearly and fairly who needs a reminder or a conversation, and remind people before things are late.',
  ],
  existing: [
    'The Staff Expectations screen already shows open times found in provider schedule workbooks, a nightly mileage sync, a cancellations list that Power Automate fills from IntakeQ, and weekly documentation counts (`admin-ui/src/components/tiles/StaffExpectationsTile.tsx`, `api/src/schedulingService.ts`, `api/src/scheduleSync.ts`, `api/src/mileageWorkbook.ts`). Extend it rather than replacing it.',
  ],
  rules: [
    'Provider performance information is private and can affect people\'s jobs. Only show it to the administrators who need it, explain how each number is worked out, and never write it to logs.',
    'The dashboard may show and remind, but a person decides about any conversation or consequence.',
    'Anything that cannot be automated must still appear as a to-do item with an owner, a due date, and a reminder, so nothing is lost.',
  ],
  sections: SECTIONS,
  questions: QUESTIONS,
  guideIntro: 'Plain explanations of what the dashboard can connect to for staff expectations, and where each one stops working.',
  guide: [
    SHARED_GUIDE.sharepoint,
    SHARED_GUIDE.intakeq,
    {
      id: 'workbooks', level: 'automatic',
      title: 'Provider schedule and mileage workbooks (already connected)',
      what: 'The dashboard already reads open times from each provider\'s schedule workbook and collects mileage nightly.',
      works: 'Open times, mileage, and schedule changes can be shown, counted, and checked on a schedule.',
      blocked: 'It only knows what providers write in the workbooks. If a provider does not update theirs, the dashboard cannot know.',
      needs: 'The rules for how and when providers must update them.',
    },
    {
      id: 'power-automate', level: 'automatic',
      title: 'Power Automate (already connected for cancellations)',
      what: 'A Power Automate flow already fills the cancellations list from IntakeQ.',
      works: 'The same pattern can add more lists or send reminders on a schedule.',
      blocked: 'Each new flow needs someone with access to build and watch it.',
      needs: 'What should trigger each flow and who owns it.',
    },
    SHARED_GUIDE.schedule,
    SHARED_GUIDE.email,
    {
      id: 'payroll', level: 'by-hand',
      title: 'Payroll',
      what: 'Getting mileage and hours to the company that pays staff.',
      works: 'The dashboard can prepare a spreadsheet for payroll in the format they need.',
      blocked: 'Sending it directly depends on what the payroll company offers. Most have no connection for outside programs, so a person uploads or emails it.',
      needs: 'The name of the payroll company and the format they accept.',
    },
    SHARED_GUIDE.paper,
  ],
  fileKinds: [
    { id: 'handbook', label: 'The staff handbook or expectations document' },
    { id: 'mileage-form', label: 'The mileage form or spreadsheet (copy with fake details)' },
    { id: 'schedule-sample', label: 'A sample provider schedule workbook (names replaced with fake ones)' },
    { id: 'cancellation-policy', label: 'The cancellation policy' },
  ].concat(COMMON_FILE_KINDS),
};
