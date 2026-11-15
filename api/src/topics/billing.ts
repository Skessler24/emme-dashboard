/**
 * The Billing setup questions. Everything the administrators see is written
 * here in plain language: no abbreviations, and every question asks about the
 * real process. Changing a question here changes the form and the generated
 * build prompt together.
 */

import { Assessment, ConnectionOption, GroupItem, Level, Option, Question, Section, TopicDef } from '../intakeTypes';

const SOURCE_KINDS: Array<Option & { description: string }> = [
  { id: 'ei', label: 'Early Intervention Community Centered Board', description: 'A regional board that pays for Early Intervention services. Each board has its own rules, so add each board separately.' },
  { id: 'medicaid', label: 'Medicaid', description: 'Medicaid or a managed plan that administers it.' },
  { id: 'private', label: 'Private insurance company', description: 'Add each insurance company separately if their rules differ.' },
  { id: 'other', label: 'Other funding source', description: 'School district, county program, grant, military insurance, families paying on their own, and so on.' },
];

const FILE_KINDS: Option[] = [
  { id: 'spreadsheet', label: 'The spreadsheet used to confirm invoices (copy with fake client details)' },
  { id: 'blank-bill', label: 'A blank bill, claim form, or invoice template' },
  { id: 'sample-bill', label: 'A sample bill that was sent (client details replaced with fake ones)' },
  { id: 'sample-payment', label: 'A sample payment report or payment explanation (client details replaced with fake ones)' },
  { id: 'sample-rejection', label: 'A sample rejection or denial notice (client details replaced with fake ones)' },
  { id: 'rebill-form', label: 'A form or template used to correct and re-send a bill' },
  { id: 'price-list', label: 'Price list or rate sheet' },
  { id: 'billing-manual', label: 'The funding source billing manual, contract, or rules document' },
  { id: 'screenshot', label: 'A screenshot of a screen you use (client details covered)' },
  { id: 'other', label: 'Something else that explains the process' },
];

const SEND_HOW: Option[] = [
  { id: 'portal-typed', label: 'Someone types each bill into the funding source website' },
  { id: 'portal-upload', label: 'Someone uploads a file or spreadsheet to the funding source website' },
  { id: 'clearinghouse', label: 'Through a clearinghouse (a middleman company that forwards bills to insurers and returns their answers)' },
  { id: 'email', label: 'We email the bill or a form' },
  { id: 'fax', label: 'We fax it' },
  { id: 'mail', label: 'We mail paper' },
  { id: 'other', label: 'Something else (describe below)' },
];

const PAY_HOW: Option[] = [
  { id: 'portal-report', label: 'We sign in to a website and look at a payment screen or report' },
  { id: 'remittance-file', label: 'We receive or download an electronic payment explanation file (electronic remittance advice)' },
  { id: 'clearinghouse-reports', label: 'Our clearinghouse shows us the results' },
  { id: 'email-notice', label: 'They email us a notice or a report' },
  { id: 'paper-remittance', label: 'A paper letter or check arrives with an explanation of payment' },
  { id: 'phone', label: 'We have to phone them to ask' },
  { id: 'bank-only', label: 'We only see a deposit in the bank and match it by hand' },
  { id: 'other', label: 'Something else (describe below)' },
];

const GENERAL_SECTIONS: Section[] = [
  { id: 'people', title: '1. People and time', intro: 'Who does this work today and how much time it takes.' },
  { id: 'today', title: '2. How a visit becomes an invoice today' },
  { id: 'sheet', title: '3. The invoice confirmation spreadsheet', intro: 'The spreadsheet your bosses use to confirm invoices. Upload a copy in the files area with all real client names and numbers replaced by fake ones.' },
  { id: 'reconcile', title: '4. Checking whether invoices were paid, and fixing the ones that were not' },
  { id: 'live', title: '5. How live the checking needs to be' },
  { id: 'screen', title: '6. What the Billing screen should show' },
  { id: 'money', title: '7. Where payments are recorded' },
  { id: 'privacy', title: '8. Privacy and who can see what' },
  { id: 'hand', title: '9. What may have to stay by hand, and what success looks like' },
];

const GENERAL_QUESTIONS: Question[] = [
  { id: 'g_who_bills', section: 'people', kind: 'text', text: 'Who does the billing work today? List each person by first name and job title, and say what each one does (for example: enters bills, checks payments, re-bills unpaid bills, approves).' },
  { id: 'g_who_approves', section: 'people', kind: 'text', text: 'Who has the final say that an invoice is correct and ready to send? Who confirms that it was paid?' },
  { id: 'g_hours', section: 'people', kind: 'text', text: 'About how many hours each week do your bosses and billing staff spend checking whether invoices were paid and re-billing the unpaid ones?' },
  { id: 'g_volume', section: 'people', kind: 'text', text: 'About how many invoices does Emme send in a typical month, and about how many dollars is that? Break it down by funding source if you can. Tell us about the busiest times of the year.' },

  { id: 'g_where_visits', section: 'today', kind: 'many', text: 'Where is the record that a visit happened kept today? (Pick all that apply.)', options: [
    { id: 'intakeq', label: 'IntakeQ appointments and notes' },
    { id: 'calendar', label: 'Outlook calendar' },
    { id: 'spreadsheet', label: 'A spreadsheet' },
    { id: 'paper', label: 'Paper' },
    { id: 'other', label: 'Another program (name it below)' },
  ] },
  { id: 'g_ready', section: 'today', kind: 'text', text: 'What must be true before a visit is allowed to be billed? (For example: the clinical note is finished and signed, the family has given permission to bill, the approval for services is still active, the visit was not cancelled.)' },
  { id: 'g_invoice_made', section: 'today', kind: 'text', text: 'Walk me through how a finished visit becomes an invoice or claim, one step at a time. Start when the visit ends and finish when the bill is sent. Say who does each step and which program, website, or spreadsheet they use.' },
  { id: 'g_intakeq_invoices', section: 'today', kind: 'one', text: 'Does IntakeQ already hold the invoices for every funding source?', options: [
    { id: 'all', label: 'Yes, all of them' },
    { id: 'some', label: 'Only some (say which ones are kept elsewhere below)' },
    { id: 'none', label: 'No, they are kept somewhere else' },
    { id: 'not-sure', label: 'I am not sure' },
  ] },

  { id: 'g_sheet_purpose', section: 'sheet', kind: 'text', text: 'Describe the spreadsheet your bosses use to confirm invoices. What is each tab for, and what is each column for?' },
  { id: 'g_sheet_updates', section: 'sheet', kind: 'text', text: 'How does information get into the spreadsheet: typed by hand, copied from a website, pasted from a report, or something else? Who updates it, and how often?' },
  { id: 'g_sheet_rules', section: 'sheet', kind: 'text', text: 'Do any colors, formulas, or notes in the spreadsheet carry meaning? Explain each one. (For example: red means more than thirty days unpaid.)' },
  { id: 'g_sheet_where', section: 'sheet', kind: 'text', text: 'Where is the spreadsheet stored? Is there one copy or several (one per funding source, one per person, one per month)?' },

  { id: 'g_status_defs', section: 'reconcile', kind: 'text', text: 'List every status an invoice can have in your process and what each one means. (For example: sent, accepted, paid in full, paid in part, rejected, denied, waiting for more information, written off.) Say who decides each status.' },
  { id: 'g_check_steps', section: 'reconcile', kind: 'text', text: 'Walk me through how your bosses check whether an invoice was paid, one step at a time. Which websites or reports do they open, in what order, and what do they compare?' },
  { id: 'g_unpaid_steps', section: 'reconcile', kind: 'text', text: 'When an invoice was not paid, what happens next, one step at a time? Include how you find the reason, who fixes it, how it is sent again, and how it is marked in the spreadsheet.' },
  { id: 'g_when_finished', section: 'reconcile', kind: 'text', text: 'When is an invoice finished? What are all the ways one can end (fully paid, written off, paid by the family, sent to appeal)?' },
  { id: 'g_partial', section: 'reconcile', kind: 'text', text: 'How do you handle an invoice that was paid for less than the amount billed, or a payment the funding source later took back?' },
  { id: 'g_worst', section: 'reconcile', kind: 'text', text: 'Describe the most frustrating or most expensive billing problem from the last year. What went wrong and how was it finally fixed?', help: 'Real stories show us the unusual cases the system has to handle.' },

  { id: 'g_frequency', section: 'live', kind: 'one', text: 'How often should the system look for payment updates?', options: [
    { id: 'hourly', label: 'Every hour' },
    { id: 'daily', label: 'Once a day, each morning' },
    { id: 'twice-weekly', label: 'Twice a week' },
    { id: 'weekly', label: 'Once a week' },
    { id: 'as-published', label: 'Whenever each funding source publishes new information' },
    { id: 'not-sure', label: 'I am not sure' },
  ] },
  { id: 'g_alert_how', section: 'live', kind: 'many', text: 'When an invoice is late, rejected, or underpaid, how should people be told? (Pick all that apply.)', options: [
    { id: 'dashboard', label: 'On the Billing screen of the dashboard' },
    { id: 'email', label: 'By email' },
    { id: 'teams', label: 'By a Microsoft Teams message' },
    { id: 'none', label: 'No notice, just show it on the screen' },
  ], help: 'Say below who should be told, and whether it is different for each funding source.' },
  { id: 'g_late_days', section: 'live', kind: 'text', text: 'How many days after an invoice is sent should it be marked late if it is still unpaid? Is that different for each funding source?' },
  { id: 'g_auto_rebill', section: 'live', kind: 'one', text: 'When a rejected invoice has a known fix, what should the system do?', options: [
    { id: 'list-only', label: 'Only add it to a list for a person to fix' },
    { id: 'prepare', label: 'Prepare the corrected invoice and wait for a person to approve it' },
    { id: 'send', label: 'Send the corrected invoice by itself' },
    { id: 'depends', label: 'It depends on the funding source or the reason (explain below)' },
  ], help: 'We recommend that a person approves anything that is sent to a funding source or that changes money.' },
  { id: 'g_approval', section: 'live', kind: 'text', text: 'Which actions must always be approved by a named person before they happen? (For example: sending any invoice to a funding source, writing off a balance, billing a family.) Who are the approvers?' },
  { id: 'g_give_up', section: 'live', kind: 'text', text: 'The goal is for the system to keep checking until every invoice is settled. What should happen to an invoice that has been re-billed several times and is still unpaid?' },

  { id: 'g_screen_items', section: 'screen', kind: 'many', text: 'What should the Billing screen show? (Pick all that apply.)', options: [
    { id: 'unpaid-by-source', label: 'Every unpaid invoice, grouped by funding source' },
    { id: 'aging', label: 'How many days each unpaid invoice has been waiting' },
    { id: 'rejection-reasons', label: 'Why each rejected invoice was rejected, and what fixes it' },
    { id: 'rebill-list', label: 'A to-do list of invoices that are ready to re-bill' },
    { id: 'deadlines', label: 'The last day to bill or re-bill each invoice' },
    { id: 'money-totals', label: 'Money billed, paid, and still owed, by month and funding source' },
    { id: 'by-provider', label: 'Billing problems grouped by provider' },
    { id: 'history', label: 'A history of every action taken on each invoice' },
    { id: 'download', label: 'Download to a spreadsheet' },
  ], help: 'Add anything else you want to see below.' },
  { id: 'g_reports', section: 'screen', kind: 'text', text: 'What reports do you have to produce, for whom, and how often? (For example: a monthly report to leadership, a report to the accountant, county or state reports, audit requests.)' },

  { id: 'g_accounting', section: 'money', kind: 'text', text: 'Where are payments recorded for bookkeeping (accounting software, a spreadsheet, a bank website)? Should the new system write to it, only read from it, or stay separate?' },
  { id: 'g_deposits', section: 'money', kind: 'text', text: 'Who matches bank deposits to invoices, and how do they do it?' },

  { id: 'g_privacy_agreement', section: 'privacy', kind: 'one', text: 'Do you have a signed privacy agreement with Microsoft and with IntakeQ? (Under the federal Health Insurance Portability and Accountability Act this is called a business associate agreement.)', options: [
    { id: 'both', label: 'Yes, with both' },
    { id: 'one', label: 'With only one of them (say which below)' },
    { id: 'no', label: 'No' },
    { id: 'not-sure', label: 'I am not sure' },
  ] },
  { id: 'g_who_sees', section: 'privacy', kind: 'text', text: 'Who is allowed to see client names and billing details? Should every administrator see everything, or only certain people?' },
  { id: 'g_retention', section: 'privacy', kind: 'text', text: 'How long must you keep billing records? Does any funding source require special record keeping?' },

  { id: 'g_hand_work', section: 'hand', kind: 'text', text: 'Which billing tasks do you believe must stay done by a person? (For example: phone calls, mailing paper, a funding source that forbids automatic sign-in.) Why?' },
  { id: 'g_first_source', section: 'hand', kind: 'text', text: 'If we could automate only one funding source first, which would save the most time or money, and why?' },
  { id: 'g_success', section: 'hand', kind: 'text', text: 'How will you know this project worked? Give a goal we can measure. (For example: "Unpaid invoices are noticed within one day instead of thirty" or "We save ten hours a week.")' },
  { id: 'g_deadline', section: 'hand', kind: 'text', text: 'Is there a date or event this needs to be ready for (a contract renewal, an audit, a staff change)?' },
];

const SOURCE_SECTIONS: Section[] = [
  { id: 'basics', title: 'The basics' },
  { id: 'permission', title: 'Permission to bill and who is covered' },
  { id: 'sending', title: 'Sending bills' },
  { id: 'payment', title: 'Finding out what was paid' },
  { id: 'problems', title: 'When bills are rejected, denied, or paid too little' },
  { id: 'special', title: 'Rules that belong to this kind of funding source' },
  { id: 'connection', title: 'Connecting the dashboard to this funding source', intro: 'Never type a password, a code, or a key here. We only need to know how the funding source works.' },
];

const SOURCE_QUESTIONS: Question[] = [
  { id: 's_what', section: 'basics', kind: 'text', text: 'In your own words, what is this funding source, who runs it, and which area or clients does it cover?' },
  { id: 's_services', section: 'basics', kind: 'text', text: 'Which Emme services are billed to this funding source? List each service by the name you use and how it is described on the bill. (For example: "occupational therapy home visit, 45 minutes".)' },
  { id: 's_contacts', section: 'basics', kind: 'text', text: 'Who can Emme call or email at this funding source when a bill has a problem? Give the role or department and how to reach them.' },
  { id: 's_different', section: 'basics', kind: 'text', text: 'What is unusual about this funding source compared with the others? List special rules, forms, or exceptions that someone new to billing would miss.' },

  { id: 's_eligibility', section: 'permission', kind: 'text', text: 'How do you confirm that a client is covered by this funding source before the visit? Where is the result recorded?' },
  { id: 's_authorization', section: 'permission', kind: 'text', text: 'Does this funding source require approval before services start (called prior authorization)? If so, what do you receive (an approval number, a number of visits or hours, an end date) and where is it tracked?' },
  { id: 's_price', section: 'permission', kind: 'text', text: 'How is the amount to bill worked out? (A set price per visit, per fifteen minutes, per hour, or per month.) Where does the price list live, and how often does it change?' },

  { id: 's_send_how', section: 'sending', kind: 'many', text: 'How is a bill sent to this funding source today? (Pick all that apply.)', options: SEND_HOW },
  { id: 's_send_steps', section: 'sending', kind: 'text', text: 'Walk me through sending a bill one step at a time. Which website, screen, or form? What is typed or attached? What confirmation appears afterward?' },
  { id: 's_bill_fields', section: 'sending', kind: 'text', text: 'What information must be on each bill? List every item this funding source asks for. (For example: client name, birth date, member number, date of service, service code, diagnosis code, provider identification number, units, price.) Upload a blank bill in the files area.' },
  { id: 's_timing', section: 'sending', kind: 'text', text: 'How soon after a visit must it be billed, and how often do you send bills (after each visit, weekly, monthly)? What happens if a deadline is missed?' },

  { id: 's_pay_how', section: 'payment', kind: 'many', text: 'How do you find out whether a bill was paid? (Pick all that apply.)', options: PAY_HOW },
  { id: 's_pay_steps', section: 'payment', kind: 'text', text: 'Walk me through checking payment one step at a time. Where do you look, and what do you compare?' },
  { id: 's_pay_days', section: 'payment', kind: 'text', text: 'How many days does payment usually take? After how many days do you treat a bill as late?' },
  { id: 's_pay_match', section: 'payment', kind: 'text', text: 'How do you match a payment to a specific bill or client? (A bill number, a client name and date, a batch number.) What happens when one payment covers many bills?' },

  { id: 's_reasons', section: 'problems', kind: 'text', text: 'List the most common reasons this funding source rejects, denies, or pays too little on a bill, and what you do to fix each one.', help: 'A rejection usually means a mistake in the information that is sent back quickly. A denial means they decided not to pay. If your funding source uses these words differently, explain how.' },
  { id: 's_rebill', section: 'problems', kind: 'text', text: 'Walk me through re-billing one step at a time. Do you send a corrected copy of the same bill or a brand new bill? What identifies the original? Do you attach anything?' },
  { id: 's_rebill_limits', section: 'problems', kind: 'text', text: 'Is there a time limit for correcting and re-sending a bill, and a limit on how many times? How does an appeal work (who, how, by when)?' },
  { id: 's_family', section: 'problems', kind: 'text', text: 'Can the family or client be billed for any part of the amount (a co-payment, a deductible, services that are not covered)? Who handles that, and how?' },

  { id: 'ei_board', section: 'special', kind: 'text', kinds: ['ei'], text: 'Which Community Centered Board is this, and which counties or regions does it cover?' },
  { id: 'ei_plan', section: 'special', kind: 'text', kinds: ['ei'], text: 'Does each bill have to match a service listed on the child\'s Individualized Family Service Plan (the written plan for the child and family)? Who keeps the plan, and how is the match checked?' },
  { id: 'ei_order', section: 'special', kind: 'text', kinds: ['ei'], text: 'Does this board require that the child\'s private insurance or Medicaid be billed first? How do you prove it was billed, and what was the answer?' },
  { id: 'ei_system', section: 'special', kind: 'text', kinds: ['ei'], text: 'Does this board have its own billing system, a state system, or forms and spreadsheets that you send in? Name each one and who runs it.' },
  { id: 'ei_logs', section: 'special', kind: 'text', kinds: ['ei'], text: 'Does this board require visit logs, session notes, or signed attendance sheets along with each bill? What exactly?' },
  { id: 'ei_fees', section: 'special', kind: 'text', kinds: ['ei'], text: 'Does the family pay a fee or a share of the cost for Early Intervention? Does Emme collect it or does the board?' },

  { id: 'mc_program', section: 'special', kind: 'text', kinds: ['medicaid'], text: 'Which Medicaid program or plan is this: the state\'s direct program or a managed plan run by another organization? How do you tell which one a client has?' },
  { id: 'mc_numbers', section: 'special', kind: 'text', kinds: ['medicaid'], text: 'Which identification numbers does Medicaid require for Emme and for each provider? Where are they kept, and who keeps them current?' },
  { id: 'mc_enrollment', section: 'special', kind: 'text', kinds: ['medicaid'], text: 'How is each provider approved to bill Medicaid? What happens to bills when a provider\'s approval is still pending or has expired?' },
  { id: 'mc_rules', section: 'special', kind: 'text', kinds: ['medicaid'], text: 'Do any rules differ for children, home visits, or video visits? List the ones that change how a bill is made.' },

  { id: 'pi_companies', section: 'special', kind: 'text', kinds: ['private'], text: 'List each private insurance company you bill, starting with the one you bill the most. For each, say whether Emme has a contract with them (in network) or not (out of network).' },
  { id: 'pi_benefits', section: 'special', kind: 'text', kinds: ['private'], text: 'How do you check a family\'s benefits (copayment, deductible, visit limits) before services? Who does it, and where is the result recorded?' },
  { id: 'pi_copay', section: 'special', kind: 'text', kinds: ['private'], text: 'Who collects copayments and deductibles, when, and how are they recorded against the invoice?' },
  { id: 'pi_credentialing', section: 'special', kind: 'text', kinds: ['private'], text: 'Is every provider fully approved by each insurance company (called credentialing)? What happens to bills for a provider who is still waiting for approval?' },
  { id: 'pi_clearinghouse', section: 'special', kind: 'text', kinds: ['private'], text: 'Do you use a clearinghouse today? If so, which company, what do you pay, and who manages the account?' },

  { id: 'ot_type', section: 'special', kind: 'text', kinds: ['other'], text: 'What kind of funding source is this? (For example: a school district, a county program, a grant, military insurance, workers\' compensation, or families paying on their own.)' },
  { id: 'ot_contract', section: 'special', kind: 'text', kinds: ['other'], text: 'Is there a contract or agreement that sets prices, billing rules, or deadlines? Summarize the key terms. You can upload a copy in the files area if it has no private details.' },

  { id: 's_site', section: 'connection', kind: 'text', text: 'What is the website address (or the name of the system) where bills are sent and payments are checked? Addresses and names only. Never type passwords here.' },
  { id: 's_second_step', section: 'connection', kind: 'one', text: 'Does signing in require a second step, like a code texted to a phone or an approval in a phone app?', options: [
    { id: 'none', label: 'No second step' },
    { id: 'text-code', label: 'A code sent by text message' },
    { id: 'app-approval', label: 'An approval in a phone app' },
    { id: 'email-code', label: 'A code sent by email' },
    { id: 'hardware-key', label: 'A physical security key' },
    { id: 'not-sure', label: 'I am not sure' },
  ] },
  { id: 's_account_owner', section: 'connection', kind: 'text', text: 'Is the sign-in shared by several people, or does each person have their own? Who owns the account, and who can ask the funding source for a separate account meant for a computer program?' },
  { id: 's_official_connection', section: 'connection', kind: 'one', text: 'Does this funding source offer an official automatic connection for computer programs (sometimes called an application programming interface), or an official way to send and receive files automatically (a secure file transfer site)?', options: [
    { id: 'connection', label: 'Yes, an automatic connection' },
    { id: 'files', label: 'Yes, a secure file transfer site' },
    { id: 'no', label: 'No, only the website' },
    { id: 'not-sure', label: 'I do not know' },
    { id: 'will-ask', label: 'I will ask them and add the answer' },
  ] },
  { id: 's_automation_allowed', section: 'connection', kind: 'one', text: 'Do this funding source\'s rules (their terms of use) allow a computer program to sign in and check on your behalf?', options: [
    { id: 'yes', label: 'Yes' },
    { id: 'no', label: 'No' },
    { id: 'not-sure', label: 'I do not know' },
    { id: 'will-ask', label: 'I will ask them and add the answer' },
  ] },
  { id: 's_downloads', section: 'connection', kind: 'text', text: 'Does the website let you download reports or payment files? Which ones, in what form (spreadsheet, text file, electronic remittance advice, which is the payment explanation a funding source sends back), and how far back do they go?' },
  { id: 's_test_area', section: 'connection', kind: 'one', text: 'Is there a practice or test area where we could try things without real clients or real money?', options: [
    { id: 'yes', label: 'Yes' },
    { id: 'no', label: 'No' },
    { id: 'not-sure', label: 'I do not know' },
  ] },
];

/** Plain explanations of each way the dashboard could connect, so people can see what is possible. */
const CONNECTION_GUIDE: ConnectionOption[] = [
  {
    id: 'clearinghouse', level: 'automatic',
    title: 'A clearinghouse connection',
    what: 'A clearinghouse is a middleman company that sends bills to many insurers and Medicaid and returns their answers.',
    works: 'The most dependable way to send bills and to get payment explanations and rejection reasons back automatically. Usually works for private insurance and Medicaid.',
    blocked: 'Costs money, needs a signed agreement, and each insurer may need a separate sign-up before it will pay through the clearinghouse. Not every funding source accepts it (Early Intervention boards often do not).',
    needs: 'Which clearinghouse Emme uses, and a connection offered by them. We never need the password.',
  },
  {
    id: 'official-connection', level: 'automatic',
    title: 'An official automatic connection from the funding source',
    what: 'Some funding sources let approved computer programs ask them questions directly (an application programming interface).',
    works: 'Allowed and dependable. The dashboard can check payment status on a schedule without a person signing in.',
    blocked: 'Only exists if the funding source offers it. Often requires an application and special credentials issued to Emme.',
    needs: 'The funding source documentation, and Emme to request access.',
  },
  {
    id: 'secure-files', level: 'automatic',
    title: 'A secure file transfer site',
    what: 'The funding source puts payment files on a private drop-off site, or accepts bills dropped there.',
    works: 'A scheduled job can pick up new payment files and read them, and can drop off prepared bills.',
    blocked: 'Needs the funding source to set up an account and share its file format.',
    needs: 'The site address, the file format description, and an account created for Emme.',
  },
  {
    id: 'email', level: 'partly',
    title: 'Notices that arrive by email',
    what: 'Payment notices or reports arrive in a shared Emme mailbox.',
    works: 'Attachments in spreadsheet or text form can be read automatically. The system can at least spot that a notice arrived.',
    blocked: 'Scanned pictures and some saved documents cannot be read reliably. Email layouts can change without warning.',
    needs: 'Samples of each email and attachment, and a shared mailbox the system may read.',
  },
  {
    id: 'upload-by-person', level: 'partly',
    title: 'A person downloads a report and drops it into the dashboard',
    what: 'A person still signs in to the funding source, downloads the report, and uploads it to the Billing screen.',
    works: 'Cuts the checking work to minutes. The dashboard does the comparing, flagging, and re-bill preparation. Works with any funding source that offers a download.',
    blocked: 'It is only as live as the person who uploads. It will not check on its own.',
    needs: 'A sample download from each funding source.',
  },
  {
    id: 'portal-signin', level: 'partly',
    title: 'The dashboard signs in to the funding source website for you',
    what: 'A program acts like a person: opens the website, signs in, and reads the screen.',
    works: 'Sometimes the only option when there is no official connection.',
    blocked: 'Often breaks when the website changes. A second sign-in step (a texted code or phone approval) stops it. Some websites forbid it in their terms of use. It cannot run on this dashboard hosting as is, and needs its own always-on helper service, so we avoid it unless nothing else works.',
    needs: 'The funding source written permission, a separate account just for the program, and an agreement that a person handles anything it cannot.',
  },
  {
    id: 'paper-phone', level: 'by-hand',
    title: 'Paper, fax, and phone calls',
    what: 'The funding source only communicates on paper, by fax, or by phone.',
    works: 'The dashboard can prepare the paperwork and keep a to-do list with due dates and reminders, so nothing is forgotten. Staff record the result by hand.',
    blocked: 'No software can read a phone call or send a fax reliably on its own. These steps stay by hand.',
    needs: 'The steps and the people who do them.',
  },
  {
    id: 'intakeq', level: 'automatic',
    title: 'IntakeQ (already connected)',
    what: 'The dashboard already reads invoices, appointments, and notes from IntakeQ.',
    works: 'Already live for reading invoices. The Billing screen lists them by status.',
    blocked: 'Today it only reads. Whether IntakeQ allows the dashboard to change an invoice (for example, mark it paid) must be confirmed before we promise that.',
    needs: 'Confirmation of what IntakeQ allows.',
  },
  {
    id: 'schedule', level: 'automatic',
    title: 'How "keeps checking until it is settled" works',
    what: 'The checking must run on a schedule, not just when someone opens the page.',
    works: 'The dashboard hosting only answers requests and cannot run a clock of its own. The project already uses a Power Automate scheduled flow that calls the dashboard on a timer, and we would do the same here (for example every morning or every hour).',
    blocked: 'It is checking on a schedule, not a continuous live link. Very frequent checks may need a separate scheduled service.',
    needs: 'Your answer to how often the checking must run.',
  },
];

// ---------------------------------------------------------------------------
// What is possible for one funding source, worked out from its answers.
// ---------------------------------------------------------------------------

const ORDER: Level[] = ['by-hand', 'partly', 'automatic'];

function choices(item: GroupItem, id: string): string[] {
  const answer = item.answers[id];
  return answer && Array.isArray(answer.choice) ? answer.choice : [];
}

function best(levels: Level[]): Level {
  let top = -1;
  levels.forEach(function (level) {
    top = Math.max(top, ORDER.indexOf(level));
  });
  return top < 0 ? 'unknown' : ORDER[top];
}

const SENDING: Record<string, { level: Level; note: string }> = {
  clearinghouse: { level: 'automatic', note: 'Sent through a clearinghouse. The dashboard can prepare the bills and a person approves them. This needs the connection the clearinghouse offers.' },
  'portal-upload': { level: 'partly', note: 'Uploaded to a website. The dashboard can prepare the file, but the upload stays by hand unless the funding source offers an automatic connection or a secure file transfer site.' },
  'portal-typed': { level: 'by-hand', note: 'Typed into a website. This stays by hand unless an official automatic connection exists. The dashboard can show a ready-to-type worksheet.' },
  email: { level: 'partly', note: 'Sent by email. The dashboard can prepare the attachment and a draft message for a person to send.' },
  fax: { level: 'by-hand', note: 'Faxed. This stays by hand. The dashboard can prepare the pages and remind people.' },
  mail: { level: 'by-hand', note: 'Mailed on paper. This stays by hand. The dashboard can prepare the pages and remind people.' },
};

const CHECKING: Record<string, { level: Level; note: string }> = {
  'remittance-file': { level: 'automatic', note: 'Payment explanation files can be read automatically once there is a safe way to receive them.' },
  'clearinghouse-reports': { level: 'automatic', note: 'The clearinghouse returns results that can be read automatically.' },
  'portal-report': { level: 'partly', note: 'Payments are only on a website screen. A person can download the report and drop it into the dashboard. Signing in automatically is fragile.' },
  'email-notice': { level: 'partly', note: 'Notices arrive by email. Spreadsheet or text attachments can be read automatically; scanned pictures cannot be read reliably.' },
  'paper-remittance': { level: 'by-hand', note: 'Payment explanations arrive on paper. A person has to key in the result.' },
  phone: { level: 'by-hand', note: 'Payment status requires a phone call. This stays by hand. The dashboard can list who to call and when.' },
  'bank-only': { level: 'by-hand', note: 'Only a bank deposit is visible. A person must match deposits to invoices.' },
};

const SECOND_STEP = ['text-code', 'app-approval', 'email-code', 'hardware-key'];

function assess(source: GroupItem): Assessment {
  const sendPicked = choices(source, 's_send_how');
  const payPicked = choices(source, 's_pay_how');
  const official = choices(source, 's_official_connection')[0] || '';
  const allowed = choices(source, 's_automation_allowed')[0] || '';
  const second = choices(source, 's_second_step')[0] || '';
  const hasOfficial = official === 'connection' || official === 'files';

  const sendNotes: string[] = [];
  const checkNotes: string[] = [];
  const blockers: string[] = [];
  const cautions: string[] = [];
  const toConfirm: string[] = [];

  const sendLevels: Level[] = [];
  sendPicked.forEach(function (id) {
    const rule = SENDING[id];
    if (rule) { sendLevels.push(rule.level); sendNotes.push(rule.note); }
    else if (id === 'other') toConfirm.push('Review the "something else" way of sending bills and decide how it can be handled.');
  });
  const payLevels: Level[] = [];
  payPicked.forEach(function (id) {
    const rule = CHECKING[id];
    if (rule) { payLevels.push(rule.level); checkNotes.push(rule.note); }
    else if (id === 'other') toConfirm.push('Review the "something else" way of finding out about payment and decide how it can be handled.');
  });

  if (hasOfficial) {
    const note = 'The funding source offers an official ' + (official === 'files' ? 'secure file transfer site' : 'automatic connection')
      + ', so both sending and checking can likely be automated once Emme is given access.';
    sendLevels.push('automatic');
    payLevels.push('automatic');
    sendNotes.push(note);
    checkNotes.push(note);
  }

  const sending = best(sendLevels);
  const checking = best(payLevels);

  if (sending === 'unknown') toConfirm.push('How bills are sent to this funding source has not been answered.');
  if (checking === 'unknown') toConfirm.push('How payment is found out has not been answered.');

  if (allowed === 'no') {
    blockers.push('The funding source rules do not allow a computer program to sign in for Emme. Do not automate through the website. Use an official connection if one exists, otherwise keep these steps by hand.');
  }
  if (!hasOfficial && SECOND_STEP.indexOf(second) !== -1) {
    blockers.push('Signing in needs a second step. A computer program cannot do that on its own without a person approving each time, unless the funding source issues an account made for programs.');
  }
  if (!hasOfficial && payPicked.length === 1 && payPicked[0] === 'portal-report' && allowed !== 'no') {
    cautions.push('The only way to see payments is by signing in to a website. Automating that is fragile and cannot run on the dashboard hosting alone. Plan on a person downloading the report, with the dashboard doing the comparing.');
  }
  if (sending === 'by-hand' || checking === 'by-hand') {
    cautions.push('Some steps stay by hand. Give those a to-do list with due dates and reminders so nothing is forgotten.');
  }

  if (official === 'not-sure' || official === 'will-ask') toConfirm.push('Ask whether the funding source offers an official automatic connection or a secure file transfer site.');
  if (!official) toConfirm.push('Whether an official automatic connection exists has not been answered.');
  if (allowed === 'not-sure' || allowed === 'will-ask') toConfirm.push('Ask whether the funding source rules allow a computer program to sign in.');
  if (second === 'not-sure') toConfirm.push('Find out whether signing in has a second step.');

  return {
    sourceId: source.id,
    name: source.name,
    kind: source.kind,
    rows: [
      { key: 'sending', label: 'Sending bills', level: sending, notes: sendNotes },
      { key: 'checking', label: 'Finding out about payment', level: checking, notes: checkNotes },
    ],
    blockers,
    cautions,
    toConfirm,
  };
}

export const billing: TopicDef = {
  id: 'billing',
  title: 'Billing',
  // Kept as it was so answers saved before the other topics existed still load.
  storagePrefix: 'billing-intake/',
  goal: [
    'Build the Billing area of the Administration Dashboard so that my bosses no longer check invoices by hand. The system must find out which invoices were paid and which were not, show the reason when one is rejected or underpaid, prepare or perform the re-billing the way each funding source requires, and keep checking on a schedule until every invoice is settled.',
    'Emme bills several kinds of funding sources: Early Intervention Community Centered Boards (each board has its own rules and process), Medicaid, private insurance companies, and others, because Emme provides more than Early Intervention services. Do not assume one funding source works like another.',
  ],
  existing: [
    'The Billing screen already lists invoices read live from IntakeQ (`api/src/intakeQService.ts`, `admin-ui/src/components/tiles/BillingTile.tsx`). Extend it rather than replacing it.',
  ],
  rules: [
    'Client billing information is sensitive health related information. Do not write client names, birth dates, member numbers, or diagnoses to logs. Only show them to the administrators who need them.',
    'A person must approve anything that is sent to a funding source or that changes money, unless I clearly say otherwise in my answers below. Keep a history of who did what to each invoice and when.',
    'Anything that cannot be automated must still appear as a to-do item with an owner, a due date, and a reminder, so nothing is lost.',
  ],
  sections: GENERAL_SECTIONS,
  questions: GENERAL_QUESTIONS,
  group: {
    singular: 'funding source',
    plural: 'funding sources',
    intro: 'Each funding source has its own rules, so answer these separately for each one. Add every Early Intervention Community Centered Board on its own.',
    namePlaceholder: 'For example: Jefferson County Community Centered Board',
    kinds: SOURCE_KINDS,
    sections: SOURCE_SECTIONS,
    questions: SOURCE_QUESTIONS,
  },
  firstId: 'g_first_source',
  firstLabel: 'First funding source to automate',
  guideIntro: 'Plain explanations of each way the dashboard could connect to a funding source, and where each one stops working.',
  guide: CONNECTION_GUIDE,
  fileKinds: FILE_KINDS,
  assess,
};
