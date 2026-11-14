import { Option, Question, Section } from './billingQuestions';

/**
 * Setup questions for the Onboarding, Referrals, Staff Information, and Staff
 * Expectations tiles. Everything the administrators read is written here in
 * plain language, with no abbreviations. Changing a question here changes the
 * form and the generated build prompt together.
 *
 * "seed" is a first draft of the answer, taken from the Emme Admin Dashboard
 * requirements document and the boss interview sheet. It is only shown until
 * someone saves the tile's answers for the first time. Seeds may use the
 * wording the bosses used (including their abbreviations).
 */

export type Seed = { choice?: string[]; text?: string; unsure?: boolean };
export type SetupQuestion = Question & { seed?: Seed };

export type TileDef = {
  id: string;
  title: string;
  /** What this tile is for, in one or two sentences. Used in the build prompt. */
  goal: string;
  /** Facts about what already exists in the project for this tile. Used in the build prompt. */
  existing: string[];
  sections: Section[];
  questions: SetupQuestion[];
  fileKinds: Option[];
};

const OTHER_FILE: Option = { id: 'other', label: 'Something else that explains the process' };
const SCREENSHOT: Option = { id: 'screenshot', label: 'A screenshot of a screen you use (client details covered)' };

const YES_NO: Option[] = [
  { id: 'yes', label: 'Yes' },
  { id: 'no', label: 'No' },
  { id: 'not-sure', label: 'I am not sure' },
];

/** The same closing questions for every tile, so the prompt always says what success means. */
function closing(prefix: string): SetupQuestion[] {
  return [
    { id: prefix + '_success', section: 'finish', kind: 'text', text: 'What would success look like? How will you know this tile has taken the work off your plate?' },
    { id: prefix + '_deadline', section: 'finish', kind: 'text', text: 'Is there a date or a busy time of year this has to be ready for?' },
    { id: prefix + '_hands_off', section: 'finish', kind: 'text', text: 'Which parts should run completely by themselves, with no administrator checking or approving? Which parts must always have a person approve them first?', help: 'For example: reminders may send by themselves, but a message to a family needs a person to approve it.' },
    { id: prefix + '_printable', section: 'finish', kind: 'one', text: 'Do you need to print or save a report from this tile?', options: YES_NO },
    { id: prefix + '_mobile', section: 'finish', kind: 'one', text: 'Does this tile need to work well on a phone?', options: YES_NO },
  ];
}

const FINISH: Section = { id: 'finish', title: 'Last: what success looks like' };

/* ------------------------------------------------------------------ */
/* Onboarding                                                          */
/* ------------------------------------------------------------------ */

const ONBOARDING: TileDef = {
  id: 'onboarding',
  title: 'Onboarding',
  goal: 'Give each new employee a live onboarding checklist and a place for their documents (certificates, licenses, insurance, and similar), keep an eye on which documents expire, tell people a month ahead, and file replaced documents away automatically.',
  existing: [
    'The Onboarding tile already exists (`admin-ui/src/components/tiles/OnboardingTile.tsx`, endpoint `/api/onboarding` in `api/src/dispatch.ts`). It shows a checklist from the SharePoint list `OnboardingChecklist` and lists files in the two folders below. Extend it rather than replacing it.',
    'Fixed SharePoint folder paths (do not change without asking me): `Employees/<employee name>/Onboarding/CurrentDocuments` and `Employees/<employee name>/Onboarding/ArchivedDocuments`.',
    'The SharePoint list `OnboardingChecklist` has the columns Title, EmployeeEmail, EmployeeName, Item, Status, DueDate. Ask me before adding columns.',
  ],
  sections: [
    { id: 'who', title: '1. Who needs a checklist', intro: 'Different jobs need different documents.' },
    { id: 'docs', title: '2. Documents and expiration dates' },
    { id: 'folders', title: '3. Folders and filing' },
    { id: 'access', title: '4. Who can do what' },
    { id: 'screen', title: '5. What the Onboarding tile should show' },
    FINISH,
  ],
  questions: [
    { id: 'o_jobs', section: 'who', kind: 'many', text: 'Which jobs need their own onboarding checklist? (Pick all that apply.)', options: [
      { id: 'slp', label: 'Speech language pathologist' },
      { id: 'ot', label: 'Occupational therapist' },
      { id: 'pt', label: 'Physical therapist' },
      { id: 'psych', label: 'Psychologist or mental health therapist' },
      { id: 'special-ed', label: 'Early childhood special educator' },
      { id: 'admin', label: 'Administrative staff' },
      { id: 'other', label: 'Another job (name it below)' },
    ], seed: { choice: ['slp', 'ot', 'pt', 'psych', 'admin'], text: 'The interview sheet lists speech language pathology, occupational therapy, physical therapy, psychology, and administration.' } },
    { id: 'o_required', section: 'who', kind: 'text', text: 'For each job, which items must be on the checklist? List them job by job. (For example: license, certificate, insurance, background check, signed policies, training.)', seed: { text: 'Employees have different needs and requirements based on their discipline. Examples named: certificates, licensures, insurance. The exact list for each discipline has not been written down yet.', unsure: true } },
    { id: 'o_active', section: 'who', kind: 'text', text: 'Which items must be finished before a new staff member counts as active and can see clients?' },
    { id: 'o_existing_lists', section: 'who', kind: 'text', text: 'Is there a checklist today in Word, Excel, or on paper? Where is it kept? Upload a copy in the files area below if you can.' },

    { id: 'o_expire', section: 'docs', kind: 'text', text: 'Which documents expire, and about how long does each one last?', seed: { text: 'The checklist document watches for expiring documents. Which documents expire (for example licenses and insurance) has not been listed yet.', unsure: true } },
    { id: 'o_notice', section: 'docs', kind: 'one', text: 'How far ahead should people be reminded that a required document is about to expire?', options: [
      { id: 'two-weeks', label: 'Two weeks ahead' },
      { id: 'one-month', label: 'One month ahead' },
      { id: 'two-months', label: 'Two months ahead' },
      { id: 'three-months', label: 'Three months ahead' },
      { id: 'other', label: 'Something else (describe below)' },
    ], seed: { choice: ['one-month'], text: 'Notify staff members a month ahead of time of expiring documents that are required.' } },
    { id: 'o_notify', section: 'docs', kind: 'many', text: 'Who should be told about an expiring document? (Pick all that apply.)', options: [
      { id: 'staff', label: 'The staff member' },
      { id: 'supervisor', label: 'Their supervisor' },
      { id: 'leaders', label: 'Jenny and Sara Beth' },
      { id: 'other-admin', label: 'Another administrator (name them below)' },
    ], seed: { choice: ['staff'], text: 'The requirements say the system notifies staff members. Whether administrators are also told was not answered.', unsure: true } },
    { id: 'o_notify_how', section: 'docs', kind: 'many', text: 'How should the reminder reach them? (Pick all that apply.)', options: [
      { id: 'email', label: 'Email' },
      { id: 'banner', label: 'The alert banner on their dashboard' },
      { id: 'teams', label: 'A message in Microsoft Teams' },
      { id: 'text', label: 'A text message' },
    ] },
    { id: 'o_overdue', section: 'docs', kind: 'text', text: 'If the new document is still not uploaded, how often should the reminder repeat? Who is told when it is overdue? What happens to the staff member then?' },

    { id: 'o_layout', section: 'folders', kind: 'text', text: 'Describe how each employee\'s folder should look and what happens when a document is replaced.', seed: { text: 'Make a SharePoint folder for each employee, named with their employee name, with a subfolder named Onboarding. Inside the Onboarding folder lives the live checklist document, which always syncs and connects only to that employee. When a document is updated, the new document is synced and the old attached documents are moved to a new subfolder labeled "Archived documents".' } },
    { id: 'o_attach', section: 'folders', kind: 'text', text: 'Besides files, what else should an administrator be able to attach to a checklist item? (For example: a web link, an expiration date, a note.)', seed: { text: 'Administrators attach documents, links, and similar items (certificates, licensures, insurance, and so on).' } },

    { id: 'o_upload', section: 'access', kind: 'one', text: 'Who can upload onboarding documents?', options: [
      { id: 'admin', label: 'Administrators only' },
      { id: 'staff', label: 'The staff member only, for their own folder' },
      { id: 'both', label: 'Both administrators and the staff member' },
    ], seed: { choice: ['admin'], text: 'The requirements describe administrators attaching the documents. Whether staff may upload their own was asked but not answered.', unsure: true } },
    { id: 'o_see', section: 'access', kind: 'text', text: 'Who can see a person\'s onboarding folder? Can the staff member see their own?' },

    { id: 'o_show', section: 'screen', kind: 'many', text: 'What should the Onboarding tile show? (Pick all that apply.)', options: [
      { id: 'status', label: 'Each person\'s overall status' },
      { id: 'missing', label: 'Which items are missing' },
      { id: 'dates', label: 'Expiration dates' },
      { id: 'files', label: 'The uploaded files' },
      { id: 'soon', label: 'A list of everything expiring soon, across all staff' },
      { id: 'archive', label: 'The archived documents' },
    ], seed: { choice: ['status', 'missing', 'dates', 'files'], text: 'The interview sheet suggests: status, missing items, expiration dates, uploaded files.' } },
    ...closing('o'),
  ],
  fileKinds: [
    { id: 'checklist', label: 'A checklist you use now (Word, Excel, or paper scan)' },
    { id: 'required-list', label: 'A list of required documents for a job' },
    { id: 'sample-doc', label: 'A sample document, such as a license or certificate (personal details covered)' },
    { id: 'policy', label: 'A policy or form every new employee signs' },
    SCREENSHOT,
    OTHER_FILE,
  ],
};

/* ------------------------------------------------------------------ */
/* Referrals                                                           */
/* ------------------------------------------------------------------ */

const REFERRALS: TileDef = {
  id: 'referrals',
  title: 'Referrals',
  goal: 'Run the referral process for new families, both Early Intervention and private, as a tracked and mostly automatic process: record each step, make sure the reaching out happens on time, show therapist openings by discipline and provider, and let an administrator assign a staff member straight from the referral.',
  existing: [
    'The Referrals tile already exists (`admin-ui/src/components/tiles/ReferralsTile.tsx`, endpoint `/api/referrals`). It keeps a simple list in the SharePoint list `Referrals` (Title, ReferrerName, ReceivedOn, Status, AssignedTo, Notes). Extend it rather than replacing it.',
    'Therapist openings already come from the Schedule & Caseload app: `GET /api/scheduling/availability` (next 14 days, no client names) and the openings view on the Staff Expectations tile. Reuse those. Do not copy schedule data into a second database.',
    'Messages to families and service coordinators are sensitive. Unless my answers below say a message may send by itself, prepare a draft and let a person approve it.',
  ],
  sections: [
    { id: 'ei', title: '1. Early Intervention referral steps' },
    { id: 'reach', title: '2. Reaching out to families and service coordinators', intro: 'The messages and the timing.' },
    { id: 'private', title: '3. Private referrals' },
    { id: 'record', title: '4. What to record' },
    { id: 'openings', title: '5. Therapist openings' },
    { id: 'assign', title: '6. Assigning staff and what happens automatically' },
    FINISH,
  ],
  questions: [
    { id: 'r_ei_steps', section: 'ei', kind: 'text', text: 'List every step of an Early Intervention referral, in order. Say who does each step.', seed: { text: 'Steps for EI: 1. Call. 2. MID (Most Important Document), and update the PAR/HIA list. 3. Provider notification. 4. Welcome email. 5. SC (service coordinator) notified. 6. Record the days it takes to place, and put that in the tracking document. 7. Remove from the referral list. The interview sheet also tracks Placed (date) and Removed.' } },
    { id: 'r_ei_boards', section: 'ei', kind: 'text', text: 'Which Early Intervention boards send referrals to Emme? Does each board work differently?', seed: { text: 'One example in the notes: "your service coordinator with RMHS". The full list of boards has not been written down yet.', unsure: true } },
    { id: 'r_ei_lists', section: 'ei', kind: 'text', text: 'Where is each referral list kept today (spreadsheet names, the list that gets updated at the Most Important Document step)? Who updates it?', seed: { text: 'Step 2 updates the PAR/HIA list. Step 7 removes the family from the referral list. What these lists are and where they live was not described.', unsure: true } },
    { id: 'r_ei_days', section: 'ei', kind: 'text', text: 'How is the number of days to place a family counted? From which date, and on which calendar?', seed: { text: 'Record how many days it took to place each family. We are on a 28 working day calendar, counted from the date of the service grid.' } },
    { id: 'r_ei_response', section: 'ei', kind: 'text', text: 'What happens when a family answers? What happens when a family answers late, after being dropped?', seed: { text: 'When a family responds, send them to SB (Sara Beth) for onboarding. If a parent responds later, the referral goes back to the active referrals.' } },

    { id: 'r_first_email', section: 'reach', kind: 'text', text: 'Write the first message to an Early Intervention family, word for word.', seed: { text: 'Email: "Hi Sam! SB & Jenny here from EMME. I received your information from your service coordinator with RMHS" (the rest of the message was not written down).', unsure: true } },
    { id: 'r_family_schedule', section: 'reach', kind: 'text', text: 'When a family does not answer, what is the schedule of messages, day by day? Say which days are texts and which are emails.', seed: { text: 'Version A (from the interview sheet): reach out every single day. If they do not answer, they get a follow-up email, then a text asking if they are still interested. If they do not answer again, they get an email saying the end of the day is the deadline. If there is still no word, email the service coordinator and consider it a dead referral.\n\nVersion B (same sheet, two-week window, three days in a row): Day 1: text, respond within 24 hours. Day 2, if no response: text or email, a personal message, 24 hour response. Day 3: personal message saying no contact and no response means we discontinue, just let me know. Day 4: reach out to the service coordinator. Only the reaching out needs to happen, and a yes or no answer is what we are waiting for.', unsure: true } },
    { id: 'r_schedule_which', section: 'reach', kind: 'text', text: 'Two different schedules were written down (see the answer above). Which one is right, or how do they fit together?' },
    { id: 'r_last_message', section: 'reach', kind: 'text', text: 'Write the last message to a family who has not answered, word for word.', seed: { text: '"Hi hi Jenny, out of respect for you and your family today will be the last day I reach out to you. I will need to hear from you by end of today or I can let your service coordinator know there has been no response. For the record, we would love to support you and your family!"' } },
    { id: 'r_sc_message', section: 'reach', kind: 'text', text: 'When do you contact the service coordinator about a family who did not answer, how often, and what do you say?', seed: { text: 'Day 4: "Hey, haven\'t heard a response, here is the phone number, can you verify and check in if they want services?" Send this once. Contact the service coordinator once a week so there is no abandonment of care.' } },
    { id: 'r_no_response_trigger', section: 'reach', kind: 'text', text: 'What exactly makes a family count as "no response" so the last message goes out?' },

    { id: 'r_private_first', section: 'private', kind: 'text', text: 'Write the first message to a private family, word for word.', seed: { text: '"Hi Jenny! Sara Beth & Jenny from EMME Therapy Services here. It looks like you reached out to us on our website for Speech and OT services. Do you have time to connect with us this week to talk about needs and next steps?"' } },
    { id: 'r_private_steps', section: 'private', kind: 'text', text: 'List every step of a private referral, in order.', seed: { text: 'Suggested on the interview sheet: first contact, needs discussion, placement, follow up, removal. These have not been confirmed.', unsure: true } },
    { id: 'r_private_sources', section: 'private', kind: 'text', text: 'Where do private referrals come from? Do any organizations send them with their own forms or templates?', seed: { text: 'The website (the example message above). The notes also say private templates need to be provided by MSU, Colorado Academy, and similar organizations. The complete list of sources has not been written down.', unsure: true } },
    { id: 'r_private_schedule', section: 'reach', kind: 'text', text: 'Do private families get the same schedule of messages as Early Intervention families? If not, what is different?' },

    { id: 'r_data', section: 'record', kind: 'many', text: 'What should be stored for each referral? (Pick all that apply.)', options: [
      { id: 'family', label: 'Family name' },
      { id: 'coordinator', label: 'Service coordinator' },
      { id: 'staff', label: 'Assigned staff member' },
      { id: 'dates', label: 'The date of each step' },
      { id: 'notes', label: 'Notes' },
      { id: 'days', label: 'Days it took to place' },
      { id: 'last-contact', label: 'Date of last contact' },
      { id: 'discipline', label: 'Which kind of therapy the family needs' },
      { id: 'source', label: 'Early Intervention board or private source' },
    ], seed: { choice: ['family', 'coordinator', 'staff', 'dates', 'notes', 'days', 'last-contact'], text: 'From the interview sheet: family name, coordinator, assigned staff, dates, notes, days to place, last contact.' } },
    { id: 'r_dead', section: 'record', kind: 'text', text: 'When is a referral dead, and where does it go? How long is it kept, and can it come back?' },

    { id: 'r_open_view', section: 'openings', kind: 'many', text: 'How should administrators see therapist openings? (Pick all that apply.)', options: [
      { id: 'provider', label: 'By provider' },
      { id: 'discipline', label: 'By kind of therapy' },
      { id: 'week', label: 'This week' },
      { id: 'month', label: 'The next thirty days, and by month' },
      { id: 'blocks', label: 'As time blocks on a day' },
    ], seed: { choice: ['provider', 'discipline', 'week', 'month'], text: 'Per discipline and provider (staff). Openings by the week and also by the month, so they know when openings are coming up. Always show openings this week and in the next 30 days.' } },
    { id: 'r_open_goal', section: 'openings', kind: 'text', text: 'What numbers matter? Should the screen warn when there are too many or too few openings?', seed: { text: 'Show the total number of openings and the total number of disciplines. The goal is no more than 20 openings.' } },

    { id: 'r_assign', section: 'assign', kind: 'one', text: 'Should an administrator be able to assign a staff member straight from the referral?', options: YES_NO, seed: { choice: ['yes'], text: 'Yes please, with automation.' } },
    { id: 'r_assign_after', section: 'assign', kind: 'text', text: 'When a staff member is assigned, what should happen by itself? (For example: tell the staff member, hold the opening, send the welcome email, tell the service coordinator.)', seed: { text: 'For Early Intervention the steps after the Most Important Document are: provider notification, welcome email, service coordinator notified. Which of these should happen automatically has not been decided.', unsure: true } },
    { id: 'r_auto', section: 'assign', kind: 'text', text: 'Which parts of the referral process do you want fully automatic, and which should a person approve first?', seed: { text: 'We need the reaching out to happen. The referral process should be automated.' } },
    ...closing('r'),
  ],
  fileKinds: [
    { id: 'tracking', label: 'The referral list or tracking spreadsheet (families\' names replaced with fake ones)' },
    { id: 'templates', label: 'Message wording or templates for families or service coordinators' },
    { id: 'private-template', label: 'A private referral template from an organization such as a university or school' },
    { id: 'service-grid', label: 'A sample service grid or approval (client details replaced with fake ones)' },
    SCREENSHOT,
    OTHER_FILE,
  ],
};

/* ------------------------------------------------------------------ */
/* Staff Information                                                   */
/* ------------------------------------------------------------------ */

const STAFF_INFO: TileDef = {
  id: 'staff-information',
  title: 'Staff Information',
  goal: 'Hold the things staff need and the things administrators check about staff in one place: required trainings and who has not finished them, quarterly growth and productivity, the quarterly ethics and rating scale spreadsheets, a question box so staff can reach administrators, and a scrolling alert banner.',
  existing: [
    'The Staff Information tile already exists (`admin-ui/src/components/tiles/StaffInformationTile.tsx`, endpoints `/api/trainings` and `/api/staff/<email>/...`). Trainings come from the SharePoint lists `Trainings` and `StaffTrainingStatus`. Extend it rather than replacing it.',
    'Staff questions and the alert banner already exist: `StaffMessages` and `AdminAlerts` lists, `admin-ui/src/components/StaffMessageAdmin.tsx`, `AlertBanner.tsx`, and `AnnouncementPanel.tsx`. Improve those instead of building new ones.',
    'The quarterly spreadsheets are already made. Do not rebuild them; connect to them and keep them live.',
    'Who is an administrator is decided in `api/src/auth.ts` (the Emme Admin group or the EMME_ADMIN_EMAILS setting).',
  ],
  sections: [
    { id: 'training', title: '1. Trainings' },
    { id: 'quarter', title: '2. Quarterly growth, ethics, and rating scales' },
    { id: 'messages', title: '3. Staff questions and alerts' },
    { id: 'access', title: '4. Who can see what' },
    FINISH,
  ],
  questions: [
    { id: 's_train_list', section: 'training', kind: 'text', text: 'What trainings does Emme have? Which are required for which jobs, and which are optional?', seed: { text: 'Trainings available to staff, the required trainings, and who has not completed them. The list of trainings and which jobs need which has not been written down yet.', unsure: true } },
    { id: 's_train_show', section: 'training', kind: 'text', text: 'What should the tile show about trainings?', seed: { text: 'The trainings available to staff, which are required, and who has not completed them.' } },
    { id: 's_train_track', section: 'training', kind: 'one', text: 'How should a completed training be recorded?', options: [
      { id: 'admin', label: 'An administrator marks it complete' },
      { id: 'upload', label: 'The staff member uploads proof, such as a certificate' },
      { id: 'automatic', label: 'The system finds out by itself from the training website' },
      { id: 'both', label: 'Staff upload proof and an administrator confirms it' },
    ] },
    { id: 's_train_where', section: 'training', kind: 'text', text: 'Where do staff take trainings today (websites, live sessions, recordings)? Where are completion records kept now?' },
    { id: 's_train_due', section: 'training', kind: 'text', text: 'When must each required training be finished, and who is reminded when it is late?' },

    { id: 's_q_what', section: 'quarter', kind: 'many', text: 'What quarterly information should be on this tile? (Pick all that apply.)', options: [
      { id: 'productivity', label: 'Productivity' },
      { id: 'growth', label: 'Growth notes and criteria' },
      { id: 'ethics', label: 'The code of ethics spreadsheet' },
      { id: 'rating', label: 'The therapist rating scales, for each kind of therapy' },
    ], seed: { choice: ['productivity', 'growth', 'ethics', 'rating'], text: 'Each quarter (Q1, Q2, Q3, Q4) administrators talk with staff about productivity, growth criteria, and similar. Quarterly spreadsheets for the code of ethics and therapist rating scales by discipline.' } },
    { id: 's_q_growth', section: 'quarter', kind: 'text', text: 'Describe how the quarterly growth conversation works. What are the productivity and growth criteria? Where are they written down?' },
    { id: 's_q_sheets', section: 'quarter', kind: 'text', text: 'Describe each quarterly spreadsheet: what it is for, who fills it in, and where it is stored.', seed: { text: 'Quarterly spreadsheets for the code of ethics and therapist rating scales by discipline. They should be live, and they are already made.' } },
    { id: 's_q_format', section: 'quarter', kind: 'one', text: 'Should the quarterly spreadsheets stay as Excel files, or be turned into SharePoint lists?', options: [
      { id: 'excel', label: 'Keep them as Excel files and show them live in the dashboard' },
      { id: 'lists', label: 'Turn them into SharePoint lists' },
      { id: 'not-sure', label: 'I am not sure' },
    ], seed: { text: 'The notes say the spreadsheets should be live and are already made, but not whether to convert them.', unsure: true } },

    { id: 's_msg_box', section: 'messages', kind: 'text', text: 'Describe how staff should ask a question or send a comment, and how administrators answer.', seed: { text: 'Staff can ask a question or send a comment to admin directly from their dashboard. It could look like a Microsoft To-Do list or a Microsoft Teams channel, but it must show as a scrolling list so admin can see and respond quickly. Responses also go back to the employees. This is done through Zoom today.' } },
    { id: 's_msg_rules', section: 'messages', kind: 'text', text: 'Who should receive staff questions? Should a question be assigned to one administrator, and how fast should it be answered?' },
    { id: 's_banner', section: 'messages', kind: 'text', text: 'Describe the alerts administrators want to send to everyone\'s dashboard.', seed: { text: 'Bosses want to send out little alerts to everyone\'s dashboard, like a scrolling banner.' } },
    { id: 's_banner_rules', section: 'messages', kind: 'text', text: 'Who may send an alert? Should alerts end by themselves on a date? Can an alert go to only some people?' },

    { id: 's_see', section: 'access', kind: 'one', text: 'Should staff see their own information on the dashboard, or only administrators?', options: [
      { id: 'own', label: 'Staff see their own information. Administrators see everyone.' },
      { id: 'admin', label: 'Only administrators see it' },
      { id: 'mixed', label: 'It depends on the item (explain below)' },
    ] },
    { id: 's_admin_who', section: 'access', kind: 'text', text: 'Who counts as an administrator for the admin tiles? (Owners, supervisors, the billing team, human resources?)' },
    { id: 's_admin_both', section: 'access', kind: 'one', text: 'Should administrators see both the admin tiles and the regular Emme dashboard tiles?', options: YES_NO },
    ...closing('s'),
  ],
  fileKinds: [
    { id: 'training-list', label: 'A list of trainings, with which are required' },
    { id: 'quarterly-sheet', label: 'A quarterly spreadsheet, such as the code of ethics (names replaced with fake ones)' },
    { id: 'rating-scale', label: 'A therapist rating scale' },
    { id: 'growth-criteria', label: 'The productivity or growth criteria' },
    SCREENSHOT,
    OTHER_FILE,
  ],
};

/* ------------------------------------------------------------------ */
/* Staff Expectations                                                  */
/* ------------------------------------------------------------------ */

const EXPECTATIONS: TileDef = {
  id: 'staff-expectations',
  title: 'Staff Expectations',
  goal: 'Automate the weekly things administrators now check by hand: mileage, cancellations, notes that are not finished and locked, evaluations and reports that are open or missing, emails checked and answered, and reminders with due dates.',
  existing: [
    'The Staff Expectations tile already exists (`admin-ui/src/components/tiles/StaffExpectationsTile.tsx`). It shows openings, mileage, cancellations, and documentation counts from the SharePoint lists `Mileage`, `Cancellations`, and `DocumentationCompliance`. Extend it rather than replacing it.',
    'Mileage forms already log confirmed visits (`api/src/mileageWorkbook.ts`) in `EMME/PAYROLL & MILEAGE/Mileage/<provider>/`. Do not change that path.',
    'Open notes, evaluations, and cancellations are read live from IntakeQ (`api/src/intakeQService.ts`) and are not copied into the scheduling database. The documentation endpoint stores counts only. Never store note text.',
    'Reminders need a timer. The API on Static Web Apps cannot run by itself, so scheduled work should be a Power Automate flow that calls the API with the automation key.',
  ],
  sections: [
    { id: 'mileage', title: '1. Mileage' },
    { id: 'cancel', title: '2. Cancellations' },
    { id: 'notes', title: '3. Notes, evaluations, and reports' },
    { id: 'email', title: '4. Email' },
    { id: 'alerts', title: '5. Reminders and alerts' },
    FINISH,
  ],
  questions: [
    { id: 'e_mile_enter', section: 'mileage', kind: 'text', text: 'How should staff record mileage? For which clients or visits?', seed: { text: 'Staff add mileage for certain clients.' } },
    { id: 'e_mile_rule', section: 'mileage', kind: 'text', text: 'What is the mileage rule? What is allowed, what is not, and who can make an exception?', seed: { text: 'Mileage cannot be billed outside of the in-person zone (5 miles). It cannot be approved outside of this, unless approved by Jenny or SB (Sara Beth).' } },
    { id: 'e_mile_approve', section: 'mileage', kind: 'many', text: 'How should mileage be approved? (Pick all that apply.)', options: [
      { id: 'enter', label: 'Staff enter their own mileage' },
      { id: 'outside', label: 'Mileage outside the five mile zone needs approval from Jenny or Sara Beth' },
      { id: 'all', label: 'An administrator approves all mileage' },
      { id: 'monthly', label: 'An administrator receives a monthly summary' },
    ], seed: { choice: ['enter', 'outside'], text: 'Whether every trip needs approval, and whether there is a monthly summary, was not answered.', unsure: true } },
    { id: 'e_mile_zone', section: 'mileage', kind: 'text', text: 'How is the five mile zone measured? From where to where? Is it a straight line or driving distance?' },

    { id: 'e_cancel_why', section: 'cancel', kind: 'text', text: 'Why do you track cancellations, and where do they get used?', seed: { text: 'They need to keep track of cancellations for productivity documentation. They go over this quarterly.' } },
    { id: 'e_cancel_types', section: 'cancel', kind: 'many', text: 'Which of these count as a cancellation for productivity? (Pick all that apply.)', options: [
      { id: 'parent', label: 'The parent cancels' },
      { id: 'therapist', label: 'The therapist cancels' },
      { id: 'no-show', label: 'The family does not show up' },
      { id: 'late', label: 'The family cancels late (describe how late below)' },
      { id: 'illness-weather', label: 'Illness or weather' },
    ] },
    { id: 'e_cancel_how', section: 'cancel', kind: 'text', text: 'How are cancellations recorded today? Should the dashboard find them by itself from IntakeQ, or should staff enter them?' },

    { id: 'e_notes_rule', section: 'notes', kind: 'text', text: 'When must notes be finished and locked? What should happen if they are not?', seed: { text: 'Complete and lock notes by the end of the week. If they are not, a reminder needs to be sent that says which notes.' } },
    { id: 'e_notes_when', section: 'notes', kind: 'one', text: 'When must notes be locked?', options: [
      { id: 'day', label: 'By the end of each day' },
      { id: 'week', label: 'By the end of each week' },
      { id: 'other', label: 'Something else (describe below)' },
    ], seed: { choice: ['week'] } },
    { id: 'e_notes_remind', section: 'notes', kind: 'text', text: 'What should trigger a reminder about open notes: a weekly list, a reminder each day they are late, or both? Who gets it: the staff member, an administrator, or both?' },
    { id: 'e_evals', section: 'notes', kind: 'text', text: 'What should be tracked for evaluations and reports? How do you know when one is due or missing?', seed: { text: 'Evals and reports that are open and missing.' } },
    { id: 'e_evals_due', section: 'notes', kind: 'text', text: 'What are the due dates for evaluations and reports? How long does a staff member have, and who is reminded?' },

    { id: 'e_email', section: 'email', kind: 'text', text: 'What does "emails checked and answered" mean? How would you know someone did it?', seed: { text: 'Emails checked and responded to.' } },
    { id: 'e_email_how', section: 'email', kind: 'one', text: 'How should it be recorded?', options: [
      { id: 'self', label: 'Staff mark it themselves each day' },
      { id: 'time', label: 'The system looks at how long a message waited for an answer' },
      { id: 'other', label: 'Something else (describe below)' },
    ] },

    { id: 'e_remind_all', section: 'alerts', kind: 'text', text: 'Which reminders, due dates, and actions do your bosses keep track of by hand today?', seed: { text: 'We need a lot of automation around report reminders, due dates, and actions that need completed. Our bosses do all of this individually, and it is a lot.' } },
    { id: 'e_alert_which', section: 'alerts', kind: 'many', text: 'Which of these should send an alert by themselves? (Pick all that apply.)', options: [
      { id: 'notes', label: 'Notes not locked by the deadline' },
      { id: 'evals', label: 'Evaluations or reports open or missing' },
      { id: 'mileage', label: 'Mileage outside the five mile zone' },
      { id: 'cancels', label: 'Too many cancellations' },
      { id: 'email', label: 'Email not checked' },
    ] },
    { id: 'e_alert_how', section: 'alerts', kind: 'many', text: 'How should alerts reach people? (Pick all that apply.)', options: [
      { id: 'email', label: 'Email' },
      { id: 'banner', label: 'The alert banner on their dashboard' },
      { id: 'teams', label: 'A message in Microsoft Teams' },
    ] },
    { id: 'e_openings', section: 'alerts', kind: 'text', text: 'What should staff schedules show about therapy openings, and who looks at them?', seed: { text: 'Auto sync schedules for therapy openings. The Schedule & Caseload database stays the source of truth, and nightly openings come from SharePoint files that Power Automate writes.' } },
    ...closing('e'),
  ],
  fileKinds: [
    { id: 'mileage-form', label: 'A mileage form or the mileage policy' },
    { id: 'cancel-sheet', label: 'The sheet used to track cancellations (names replaced with fake ones)' },
    { id: 'report-list', label: 'A list of open notes, evaluations, or reports (names replaced with fake ones)' },
    { id: 'expectations-doc', label: 'A document that lists what staff are expected to do' },
    SCREENSHOT,
    OTHER_FILE,
  ],
};

export const TILE_SETUPS: TileDef[] = [ONBOARDING, REFERRALS, STAFF_INFO, EXPECTATIONS];

export function tileSetup(id: string): TileDef | undefined {
  return TILE_SETUPS.find(function (tile) { return tile.id === id; });
}
