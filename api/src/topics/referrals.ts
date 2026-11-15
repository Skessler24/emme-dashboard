import { Assessment, GroupItem, Level, Option, Question, Section, TopicDef } from '../intakeTypes';
import { COMMON_FILE_KINDS, SHARED_GUIDE, closingQuestions, privacyQuestions, reminderQuestions } from './shared';

const SECTIONS: Section[] = [
  { id: 'people', title: '1. People and volume' },
  { id: 'steps', title: '2. From a referral arriving to a first appointment', intro: 'The real steps, in order, the way they happen today.' },
  { id: 'status', title: '3. What each status means' },
  { id: 'match', title: '4. Choosing a provider for a referral' },
  { id: 'waiting', title: '5. Waiting lists and follow-up' },
  { id: 'live', title: '6. Reminders and who is told' },
  { id: 'privacy', title: '7. Privacy and who can see what' },
  { id: 'hand', title: '8. What may have to stay by hand, and what success looks like' },
];

const ARRIVES: Option[] = [
  { id: 'portal', label: 'We sign in to a website and look for new referrals' },
  { id: 'official', label: 'They send them electronically through an official connection or a secure file transfer site' },
  { id: 'email', label: 'They email us' },
  { id: 'form', label: 'They fill in a form on our website or in IntakeQ' },
  { id: 'fax', label: 'They fax us' },
  { id: 'phone', label: 'They phone us' },
  { id: 'paper', label: 'Paper arrives by mail or in person' },
  { id: 'other', label: 'Something else (describe below)' },
];

const REPLY: Option[] = [
  { id: 'portal', label: 'We type our answer into their website' },
  { id: 'email', label: 'We email them' },
  { id: 'phone', label: 'We phone them' },
  { id: 'fax', label: 'We fax them' },
  { id: 'official', label: 'We send it through an official connection or secure file transfer site' },
  { id: 'none', label: 'No reply is expected' },
  { id: 'other', label: 'Something else (describe below)' },
];

const QUESTIONS: Question[] = [
  { id: 'g_who_handles', section: 'people', kind: 'text', text: 'Who handles referrals today? List each person by first name and job title, and say what each one does (for example: receives, decides, phones the family, matches to a provider).' },
  { id: 'g_who_decides', section: 'people', kind: 'text', text: 'Who has the final say on whether Emme accepts a referral?' },
  { id: 'g_volume', section: 'people', kind: 'text', text: 'About how many referrals does Emme receive in a typical month? Break it down by where they come from if you can, and tell us about the busiest times of the year.' },
  { id: 'g_hours', section: 'people', kind: 'text', text: 'About how many hours each week do people spend on referrals, and how much of that is typing the same information into more than one place?' },

  { id: 'g_steps', section: 'steps', kind: 'text', text: 'Walk me through one referral, one step at a time. Start when it arrives and finish when the first appointment happens. Say who does each step and which program, list, or spreadsheet they use.' },
  { id: 'g_list', section: 'steps', kind: 'text', text: 'The dashboard already shows a referral list kept in SharePoint with the referrer name, status, who it is assigned to, the date received, and notes. Is that the list you really use? What is missing, and what is there that nobody uses?' },
  { id: 'g_before_accept', section: 'steps', kind: 'text', text: 'What must be known before Emme can accept a referral? (For example: the child\'s age, the service needed, the address, the funding source, a signed permission.)' },
  { id: 'g_decline', section: 'steps', kind: 'text', text: 'When Emme cannot accept a referral, what happens? Who tells the referrer and the family, how, and what do they say?' },
  { id: 'g_duplicates', section: 'steps', kind: 'text', text: 'How do you handle the same family being referred twice, or a referral for someone who is already a client?' },
  { id: 'g_worst', section: 'steps', kind: 'text', text: 'Describe the last referral that was lost, late, or went wrong. What happened and how was it fixed?', help: 'Real stories show us the unusual cases the system has to handle.' },

  { id: 'g_status_defs', section: 'status', kind: 'text', text: 'List every status a referral can have in your process and what each one means. (For example: new, contacted, waiting for family, scheduled, accepted, declined, closed.) Say who changes each status.' },
  { id: 'g_status_time', section: 'status', kind: 'text', text: 'How long should a referral stay in each status before someone should be warned that it is stuck?' },
  { id: 'g_response_time', section: 'status', kind: 'text', text: 'How quickly must Emme respond to a new referral? Does a funding source, a contract, or a law set that time? Is it different for each referral source?' },

  { id: 'g_match_factors', section: 'match', kind: 'many', text: 'What decides which provider gets a referral? (Pick all that apply.)', options: [
    { id: 'service', label: 'The kind of service needed' },
    { id: 'location', label: 'Where the family lives' },
    { id: 'availability', label: 'The provider\'s open times' },
    { id: 'language', label: 'The language the family speaks' },
    { id: 'funding', label: 'Which funding sources the provider can bill' },
    { id: 'caseload', label: 'How many clients the provider already has' },
    { id: 'experience', label: 'The provider\'s experience with a need' },
    { id: 'family-choice', label: 'The family\'s preference' },
  ], help: 'Say below how these are weighed, and anything else that matters.' },
  { id: 'g_match_who', section: 'match', kind: 'text', text: 'Who matches a referral to a provider today, and how do they find out who has room? The dashboard already reads open times from provider schedule workbooks. Would a list of providers with open times near the family be useful?' },
  { id: 'g_match_approval', section: 'match', kind: 'one', text: 'Should the dashboard suggest a provider, or only show the facts and let a person decide?', options: [
    { id: 'facts', label: 'Only show the facts' },
    { id: 'suggest', label: 'Suggest a provider and wait for a person to approve' },
    { id: 'not-sure', label: 'I am not sure' },
  ], help: 'We recommend that a person always decides who a family is matched with.' },

  { id: 'g_waitlist', section: 'waiting', kind: 'text', text: 'What happens when no provider has room? How is the waiting list kept, who is told, and how do families find out how long the wait is?' },
  { id: 'g_followup', section: 'waiting', kind: 'text', text: 'How many times and how often do you try to reach a family before closing a referral? What do you say each time, and how is it recorded?' },
  { id: 'g_reopen', section: 'waiting', kind: 'text', text: 'When a closed referral comes back to life (the family calls again), what do you do?' },

  ...reminderQuestions('live', 'referrals'),
  ...privacyQuestions('privacy', 'referral', 'referral details such as a child\'s name, birth date, and needs'),
  { id: 'g_who_sees_notes', section: 'privacy', kind: 'text', text: 'Should the notes on a referral be visible to every administrator, or only to the people who work on referrals?' },
  { id: 'g_first_source', section: 'hand', kind: 'text', text: 'If we could automate only one referral source first, which would save the most time or prevent the most lost referrals, and why?' },
  ...closingQuestions('hand', 'referral', 'phoning a family, reading a faxed form, a source that only accepts replies by phone.'),
];

const SOURCE_SECTIONS: Section[] = [
  { id: 'basics', title: 'The basics' },
  { id: 'arrive', title: 'How referrals arrive' },
  { id: 'info', title: 'What a referral contains' },
  { id: 'reply', title: 'Answering the referrer' },
  { id: 'connection', title: 'Connecting the dashboard to this source', intro: 'Never type a password, a code, or a key here. We only need to know how the source works.' },
];

const SOURCE_QUESTIONS: Question[] = [
  { id: 'r_what', section: 'basics', kind: 'text', text: 'In your own words, who is this referral source, what do they do, and how often do they send referrals?' },
  { id: 'r_contacts', section: 'basics', kind: 'text', text: 'Who can Emme call or email at this source when a referral has a problem? Give the role or department and how to reach them.' },
  { id: 'r_different', section: 'basics', kind: 'text', text: 'What is unusual about this source compared with the others? List special rules, forms, or deadlines that someone new would miss.' },
  { id: 'r_ei_board', section: 'basics', kind: 'text', kinds: ['ei'], text: 'Which Community Centered Board is this, and which counties or regions does it cover? How is a referral from this board different from the other boards?' },
  { id: 'r_ei_deadline', section: 'basics', kind: 'text', kinds: ['ei'], text: 'Does this board set a deadline for contacting the family or for the first visit? What happens if it is missed?' },
  { id: 'r_school_age', section: 'basics', kind: 'text', kinds: ['school'], text: 'Which schools or school districts send referrals, and what ages or needs do they usually involve?' },

  { id: 'r_arrive', section: 'arrive', kind: 'many', text: 'How does a referral arrive from this source today? (Pick all that apply.)', options: ARRIVES },
  { id: 'r_arrive_steps', section: 'arrive', kind: 'text', text: 'Walk me through how someone at Emme finds out about a new referral from this source and gets it into the referral list. Which website, inbox, or form, and what do they do first?' },
  { id: 'r_notice', section: 'arrive', kind: 'text', text: 'How quickly does Emme usually notice a new referral from this source? Has one ever sat unseen for days?' },

  { id: 'r_fields', section: 'info', kind: 'text', text: 'What information does a referral from this source carry? List every item. (For example: child\'s name, birth date, parent contact, address, reason for the referral, funding source, approval number.) Upload a blank form or a sample with fake details in the files area.' },
  { id: 'r_missing', section: 'info', kind: 'text', text: 'What is often missing or wrong in a referral from this source, and how do you get it fixed?' },
  { id: 'r_approval', section: 'info', kind: 'text', text: 'Does this source have to approve services before they start (called prior authorization)? If so, what do you receive and where is it tracked?' },

  { id: 'r_reply', section: 'reply', kind: 'many', text: 'How does Emme answer this source? (Pick all that apply.)', options: REPLY },
  { id: 'r_reply_what', section: 'reply', kind: 'text', text: 'What must Emme tell this source, and when? (For example: that the referral was received, that it was accepted, the date of the first visit, that it was declined and why.)' },
  { id: 'r_reply_who', section: 'reply', kind: 'text', text: 'Who sends the answer today, and how is it recorded that it was sent?' },

  { id: 'r_site', section: 'connection', kind: 'text', text: 'What is the website address (or the name of the system) where referrals from this source are found or answered? Addresses and names only. Never type passwords here.' },
  { id: 'r_second_step', section: 'connection', kind: 'one', text: 'Does signing in require a second step, like a code texted to a phone or an approval in a phone app?', options: [
    { id: 'none', label: 'No second step' },
    { id: 'text-code', label: 'A code sent by text message' },
    { id: 'app-approval', label: 'An approval in a phone app' },
    { id: 'email-code', label: 'A code sent by email' },
    { id: 'hardware-key', label: 'A physical security key' },
    { id: 'not-sure', label: 'I am not sure' },
  ] },
  { id: 'r_official_connection', section: 'connection', kind: 'one', text: 'Does this source offer an official automatic connection for computer programs (sometimes called an application programming interface), or an official way to send and receive files automatically (a secure file transfer site)?', options: [
    { id: 'connection', label: 'Yes, an automatic connection' },
    { id: 'files', label: 'Yes, a secure file transfer site' },
    { id: 'no', label: 'No, only the website, email, or paper' },
    { id: 'not-sure', label: 'I do not know' },
    { id: 'will-ask', label: 'I will ask them and add the answer' },
  ] },
  { id: 'r_automation_allowed', section: 'connection', kind: 'one', text: 'Do this source\'s rules (their terms of use) allow a computer program to sign in and check for referrals on your behalf?', options: [
    { id: 'yes', label: 'Yes' },
    { id: 'no', label: 'No' },
    { id: 'not-sure', label: 'I do not know' },
    { id: 'will-ask', label: 'I will ask them and add the answer' },
  ] },
];

const ARRIVE_RULES: Record<string, { level: Level; note: string }> = {
  official: { level: 'automatic', note: 'Referrals arrive through an official connection, so they can be read and added to the referral list automatically once Emme is given access.' },
  form: { level: 'automatic', note: 'Referrals arrive in a form Emme controls. The dashboard can read the form and add each one to the referral list.' },
  email: { level: 'partly', note: 'Referrals arrive by email. The dashboard can spot new messages in a shared mailbox and read spreadsheet or text attachments. Scanned pictures cannot be read reliably. This needs extra permission to read the mailbox.' },
  portal: { level: 'partly', note: 'Referrals are only on a website. A person can download a list and drop it into the dashboard. Signing in automatically is fragile and cannot run on the dashboard hosting alone.' },
  fax: { level: 'by-hand', note: 'Referrals arrive by fax. A person has to enter them, though the dashboard can make that quick with a short form and a reminder.' },
  phone: { level: 'by-hand', note: 'Referrals arrive by phone. A person has to enter them, though the dashboard can make that quick with a short form.' },
  paper: { level: 'by-hand', note: 'Referrals arrive on paper. A person has to enter them.' },
};

const REPLY_RULES: Record<string, { level: Level; note: string }> = {
  official: { level: 'automatic', note: 'Answers can go back through the official connection once Emme is given access, with a person approving.' },
  email: { level: 'partly', note: 'The dashboard can prepare an email for a person to send, or send one from a shared mailbox after approval (this needs extra permission).' },
  portal: { level: 'by-hand', note: 'Answers are typed into the source\'s website. This stays by hand unless an official connection exists. The dashboard can show a ready-to-type summary.' },
  phone: { level: 'by-hand', note: 'Answers are given by phone. This stays by hand. The dashboard can list who to call and when.' },
  fax: { level: 'by-hand', note: 'Answers are faxed. This stays by hand. The dashboard can prepare the page and remind people.' },
  none: { level: 'automatic', note: 'No reply is expected, so nothing needs sending.' },
};

const ORDER: Level[] = ['by-hand', 'partly', 'automatic'];
const SECOND = ['text-code', 'app-approval', 'email-code', 'hardware-key'];

function choices(item: GroupItem, id: string): string[] {
  const answer = item.answers[id];
  return answer && Array.isArray(answer.choice) ? answer.choice : [];
}

function best(levels: Level[]): Level {
  let top = -1;
  levels.forEach(function (level) { top = Math.max(top, ORDER.indexOf(level)); });
  return top < 0 ? 'unknown' : ORDER[top];
}

function assess(item: GroupItem): Assessment {
  const official = choices(item, 'r_official_connection')[0] || '';
  const allowed = choices(item, 'r_automation_allowed')[0] || '';
  const second = choices(item, 'r_second_step')[0] || '';
  const hasOfficial = official === 'connection' || official === 'files';
  const blockers: string[] = [];
  const cautions: string[] = [];
  const toConfirm: string[] = [];

  const arriveLevels: Level[] = [];
  const arriveNotes: string[] = [];
  const arrivePicked = choices(item, 'r_arrive');
  arrivePicked.forEach(function (id) {
    const rule = ARRIVE_RULES[id];
    if (rule) { arriveLevels.push(rule.level); arriveNotes.push(rule.note); }
    else if (id === 'other') toConfirm.push('Review the "something else" way referrals arrive and decide how it can be handled.');
  });
  const replyLevels: Level[] = [];
  const replyNotes: string[] = [];
  choices(item, 'r_reply').forEach(function (id) {
    const rule = REPLY_RULES[id];
    if (rule) { replyLevels.push(rule.level); replyNotes.push(rule.note); }
    else if (id === 'other') toConfirm.push('Review the "something else" way of answering this source and decide how it can be handled.');
  });
  if (hasOfficial) {
    const note = 'The source offers an official ' + (official === 'files' ? 'secure file transfer site' : 'automatic connection') + ', so receiving and answering can likely be automated once Emme is given access.';
    arriveLevels.push('automatic');
    replyLevels.push('automatic');
    arriveNotes.push(note);
    replyNotes.push(note);
  }
  const arrive = best(arriveLevels);
  const reply = best(replyLevels);

  if (arrive === 'unknown') toConfirm.push('How referrals arrive from this source has not been answered.');
  if (reply === 'unknown') toConfirm.push('How Emme answers this source has not been answered.');
  if (allowed === 'no') blockers.push('The source rules do not allow a computer program to sign in for Emme. Do not automate through the website. Use an official connection if one exists, otherwise keep these steps by hand.');
  if (!hasOfficial && SECOND.indexOf(second) !== -1) blockers.push('Signing in needs a second step. A computer program cannot do that on its own without a person approving each time.');
  if (!hasOfficial && arrivePicked.length === 1 && arrivePicked[0] === 'portal' && allowed !== 'no') {
    cautions.push('The only way to see new referrals is by signing in to a website. Automating that is fragile. Plan on a person downloading the list, with the dashboard adding the referrals and raising reminders.');
  }
  if (arrive === 'by-hand' || reply === 'by-hand') cautions.push('Some steps stay by hand. Give those a to-do list with due dates and reminders so no referral is forgotten.');
  if (official === 'not-sure' || official === 'will-ask') toConfirm.push('Ask whether the source offers an official automatic connection or a secure file transfer site.');
  if (allowed === 'not-sure' || allowed === 'will-ask') toConfirm.push('Ask whether the source rules allow a computer program to sign in.');
  if (second === 'not-sure') toConfirm.push('Find out whether signing in has a second step.');

  return {
    sourceId: item.id,
    name: item.name,
    kind: item.kind,
    rows: [
      { key: 'arrive', label: 'Receiving referrals', level: arrive, notes: arriveNotes },
      { key: 'reply', label: 'Answering the referrer', level: reply, notes: replyNotes },
    ],
    blockers,
    cautions,
    toConfirm,
  };
}

export const referrals: TopicDef = {
  id: 'referrals',
  title: 'Referrals',
  storagePrefix: 'intake/referrals/',
  goal: [
    'Build the Referrals area of the Administration Dashboard so that no referral is lost, late, or typed in twice. The dashboard should get new referrals into one list, show who is handling each one and how long it has waited, help match each referral to a provider who has room, and remind people before anything goes stale.',
    'Referrals come from several kinds of sources, such as Early Intervention Community Centered Boards (each with its own process), school districts, doctors and clinics, and families who call directly. Do not assume one source works like another.',
  ],
  existing: [
    'The Referrals screen already shows a referral list kept in SharePoint with the referrer name, status, who it is assigned to, the date received, and notes (`admin-ui/src/components/tiles/ReferralsTile.tsx`). Provider open times are already read from the schedule workbooks (`api/src/schedulingService.ts`). Extend these rather than replacing them.',
  ],
  rules: [
    'Referral details (a child\'s name, birth date, needs) are sensitive health related information. Do not write them to logs. Only show them to the administrators who need them.',
    'A person must decide who a family is matched with and must approve anything sent to a referral source, unless I clearly say otherwise in my answers below.',
    'Anything that cannot be automated must still appear as a to-do item with an owner, a due date, and a reminder, so no referral is lost.',
  ],
  sections: SECTIONS,
  questions: QUESTIONS,
  group: {
    singular: 'referral source',
    plural: 'referral sources',
    intro: 'Each kind of referral source works differently, so answer these separately for each one. Add every Early Intervention Community Centered Board on its own.',
    namePlaceholder: 'For example: Jefferson County Community Centered Board',
    kinds: [
      { id: 'ei', label: 'Early Intervention Community Centered Board', description: 'A regional board that sends Early Intervention referrals. Each board has its own rules, so add each board separately.' },
      { id: 'school', label: 'School district', description: 'A school or district that refers children.' },
      { id: 'clinic', label: 'Doctor\'s office or clinic', description: 'A medical practice that refers patients.' },
      { id: 'family', label: 'Families contacting Emme directly', description: 'Families who call, email, or fill in a form themselves.' },
      { id: 'other', label: 'Another organization', description: 'A county program, a community group, another therapy practice, and so on.' },
    ],
    sections: SOURCE_SECTIONS,
    questions: SOURCE_QUESTIONS,
  },
  firstId: 'g_first_source',
  firstLabel: 'First referral source to automate',
  guideIntro: 'Plain explanations of each way referrals could reach the dashboard, and where each one stops working.',
  guide: [
    SHARED_GUIDE.sharepoint,
    {
      id: 'form', level: 'automatic',
      title: 'An online referral form',
      what: 'A form on the Emme website or in IntakeQ that referrers and families fill in.',
      works: 'Each answer can land in the referral list automatically with the date received. This is the most dependable way for any source that has no official connection.',
      blocked: 'Only helps if the source is willing to use the form instead of email, fax, or phone.',
      needs: 'The list of information the form must ask for.',
    },
    {
      id: 'official-connection', level: 'automatic',
      title: 'An official automatic connection or secure file transfer site',
      what: 'Some sources, especially public boards, let approved programs receive and send referrals directly.',
      works: 'Allowed and dependable, with no one signing in.',
      blocked: 'Only exists if the source offers it, and it usually needs an application and special access for Emme.',
      needs: 'The source\'s documentation and Emme asking for access.',
    },
    SHARED_GUIDE.email,
    SHARED_GUIDE.intakeq,
    {
      id: 'schedule-match', level: 'automatic',
      title: 'Matching a referral to a provider with open times',
      what: 'Showing which providers have room near the family.',
      works: 'The dashboard already reads open times from the provider schedule workbooks. It can list providers with room for a service and a day, with a person making the final decision.',
      blocked: 'It cannot know a provider\'s experience, language, or funding approvals unless those are written down somewhere it can read.',
      needs: 'Where that information is kept today.',
    },
    SHARED_GUIDE.schedule,
    SHARED_GUIDE.paper,
  ],
  fileKinds: [
    { id: 'blank-referral', label: 'A blank referral form' },
    { id: 'sample-referral', label: 'A sample referral (client details replaced with fake ones)' },
    { id: 'status-list', label: 'The list or spreadsheet used to track referrals (copy with fake client details)' },
    { id: 'reply-template', label: 'A message or letter sent back to a referrer' },
  ].concat(COMMON_FILE_KINDS),
  assess,
};
