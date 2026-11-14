"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.assessSource = assessSource;
exports.missingItems = missingItems;
exports.buildPrompt = buildPrompt;
const billingQuestions_1 = require("./billingQuestions");
const ORDER = ['by-hand', 'partly', 'automatic'];
function choices(source, id) {
    const answer = source.answers[id];
    return answer && Array.isArray(answer.choice) ? answer.choice : [];
}
function best(levels) {
    let top = -1;
    levels.forEach(function (level) {
        top = Math.max(top, ORDER.indexOf(level));
    });
    return top < 0 ? 'unknown' : ORDER[top];
}
const SENDING = {
    clearinghouse: { level: 'automatic', note: 'Sent through a clearinghouse. The dashboard can prepare the bills and a person approves them. This needs the connection the clearinghouse offers.' },
    'portal-upload': { level: 'partly', note: 'Uploaded to a website. The dashboard can prepare the file, but the upload stays by hand unless the funding source offers an automatic connection or a secure file transfer site.' },
    'portal-typed': { level: 'by-hand', note: 'Typed into a website. This stays by hand unless an official automatic connection exists. The dashboard can show a ready-to-type worksheet.' },
    email: { level: 'partly', note: 'Sent by email. The dashboard can prepare the attachment and a draft message for a person to send.' },
    fax: { level: 'by-hand', note: 'Faxed. This stays by hand. The dashboard can prepare the pages and remind people.' },
    mail: { level: 'by-hand', note: 'Mailed on paper. This stays by hand. The dashboard can prepare the pages and remind people.' },
};
const CHECKING = {
    'remittance-file': { level: 'automatic', note: 'Payment explanation files can be read automatically once there is a safe way to receive them.' },
    'clearinghouse-reports': { level: 'automatic', note: 'The clearinghouse returns results that can be read automatically.' },
    'portal-report': { level: 'partly', note: 'Payments are only on a website screen. A person can download the report and drop it into the dashboard. Signing in automatically is fragile.' },
    'email-notice': { level: 'partly', note: 'Notices arrive by email. Spreadsheet or text attachments can be read automatically; scanned pictures cannot be read reliably.' },
    'paper-remittance': { level: 'by-hand', note: 'Payment explanations arrive on paper. A person has to key in the result.' },
    phone: { level: 'by-hand', note: 'Payment status requires a phone call. This stays by hand. The dashboard can list who to call and when.' },
    'bank-only': { level: 'by-hand', note: 'Only a bank deposit is visible. A person must match deposits to invoices.' },
};
const SECOND_STEP = ['text-code', 'app-approval', 'email-code', 'hardware-key'];
function assessSource(source) {
    const sendPicked = choices(source, 's_send_how');
    const payPicked = choices(source, 's_pay_how');
    const official = choices(source, 's_official_connection')[0] || '';
    const allowed = choices(source, 's_automation_allowed')[0] || '';
    const second = choices(source, 's_second_step')[0] || '';
    const hasOfficial = official === 'connection' || official === 'files';
    const sending = { level: 'unknown', notes: [] };
    const checking = { level: 'unknown', notes: [] };
    const blockers = [];
    const cautions = [];
    const toConfirm = [];
    const sendLevels = [];
    sendPicked.forEach(function (id) {
        const rule = SENDING[id];
        if (rule) {
            sendLevels.push(rule.level);
            sending.notes.push(rule.note);
        }
        else if (id === 'other')
            toConfirm.push('Review the "something else" way of sending bills and decide how it can be handled.');
    });
    const payLevels = [];
    payPicked.forEach(function (id) {
        const rule = CHECKING[id];
        if (rule) {
            payLevels.push(rule.level);
            checking.notes.push(rule.note);
        }
        else if (id === 'other')
            toConfirm.push('Review the "something else" way of finding out about payment and decide how it can be handled.');
    });
    if (hasOfficial) {
        const note = 'The funding source offers an official ' + (official === 'files' ? 'secure file transfer site' : 'automatic connection')
            + ', so both sending and checking can likely be automated once Emme is given access.';
        sendLevels.push('automatic');
        payLevels.push('automatic');
        sending.notes.push(note);
        checking.notes.push(note);
    }
    sending.level = best(sendLevels);
    checking.level = best(payLevels);
    if (sending.level === 'unknown')
        toConfirm.push('How bills are sent to this funding source has not been answered.');
    if (checking.level === 'unknown')
        toConfirm.push('How payment is found out has not been answered.');
    if (allowed === 'no') {
        blockers.push('The funding source rules do not allow a computer program to sign in for Emme. Do not automate through the website. Use an official connection if one exists, otherwise keep these steps by hand.');
    }
    if (!hasOfficial && SECOND_STEP.indexOf(second) !== -1) {
        blockers.push('Signing in needs a second step. A computer program cannot do that on its own without a person approving each time, unless the funding source issues an account made for programs.');
    }
    if (!hasOfficial && payPicked.length === 1 && payPicked[0] === 'portal-report' && allowed !== 'no') {
        cautions.push('The only way to see payments is by signing in to a website. Automating that is fragile and cannot run on the dashboard hosting alone. Plan on a person downloading the report, with the dashboard doing the comparing.');
    }
    if (sending.level === 'by-hand' || checking.level === 'by-hand') {
        cautions.push('Some steps stay by hand. Give those a to-do list with due dates and reminders so nothing is forgotten.');
    }
    if (official === 'not-sure' || official === 'will-ask')
        toConfirm.push('Ask whether the funding source offers an official automatic connection or a secure file transfer site.');
    if (!official)
        toConfirm.push('Whether an official automatic connection exists has not been answered.');
    if (allowed === 'not-sure' || allowed === 'will-ask')
        toConfirm.push('Ask whether the funding source rules allow a computer program to sign in.');
    if (second === 'not-sure')
        toConfirm.push('Find out whether signing in has a second step.');
    return { sourceId: source.id, name: source.name, kind: source.kind, sending, checking, blockers, cautions, toConfirm };
}
function hasContent(answer) {
    return !!answer && (answer.choice.length > 0 || answer.text.trim().length > 0);
}
function kindLabel(kind) {
    const found = billingQuestions_1.SOURCE_KINDS.find(function (item) { return item.id === kind; });
    return found ? found.label : kind;
}
function levelLabel(level) {
    if (level === 'automatic')
        return 'likely automatic';
    if (level === 'partly')
        return 'partly automatic';
    if (level === 'by-hand')
        return 'by hand';
    return 'not enough information yet';
}
function renderAnswer(question, answer) {
    if (!hasContent(answer)) {
        return answer && answer.unsure ? '_Marked "not sure yet". Needs follow-up before building._' : '_Not answered._';
    }
    const parts = [];
    if (answer.choice.length) {
        const labels = answer.choice.map(function (id) {
            const option = (question.options || []).find(function (item) { return item.id === id; });
            return option ? option.label : id;
        });
        parts.push(labels.join('; '));
    }
    if (answer.text.trim())
        parts.push(answer.text.trim());
    if (answer.unsure)
        parts.push('_(Marked "not sure yet". Treat as unconfirmed.)_');
    return parts.join('\n\n');
}
function missingItems(data) {
    const out = [];
    function scan(where, questions, answers) {
        questions.forEach(function (question) {
            const answer = answers[question.id];
            if (answer && answer.unsure)
                out.push({ where, question: question.text, reason: 'not-sure' });
            else if (!hasContent(answer))
                out.push({ where, question: question.text, reason: 'unanswered' });
        });
    }
    scan('Billing overall', billingQuestions_1.GENERAL_QUESTIONS, data.general);
    data.sources.forEach(function (source) {
        scan(source.name || 'Funding source', (0, billingQuestions_1.questionsForKind)(source.kind), source.answers);
    });
    return out;
}
function fileLine(file) {
    const kind = billingQuestions_1.FILE_KINDS.find(function (item) { return item.id === file.kind; });
    return '- `' + file.name + '`: ' + (kind ? kind.label : file.kind) + (file.note ? '. ' + file.note : '');
}
function buildPrompt(data, files, now) {
    const lines = [];
    const add = function (text) { lines.push(text); };
    const when = now.toISOString().slice(0, 10);
    const missing = missingItems(data);
    const unknownCount = missing.length;
    add('# Build request: Billing for the Emme Administration Dashboard');
    add('');
    add('_Generated ' + when + ' from the Billing setup answers. Last saved by ' + (data.updatedBy || 'nobody yet') + (data.updatedAt ? ' on ' + data.updatedAt.slice(0, 10) : '') + '._');
    add('');
    add('## What I want');
    add('');
    add('Build the Billing area of the Administration Dashboard so that my bosses no longer check invoices by hand. The system must find out which invoices were paid and which were not, show the reason when one is rejected or underpaid, prepare or perform the re-billing the way each funding source requires, and keep checking on a schedule until every invoice is settled.');
    add('');
    add('Emme bills several kinds of funding sources: Early Intervention Community Centered Boards (each board has its own rules and process), Medicaid, private insurance companies, and others, because Emme provides more than Early Intervention services. Do not assume one funding source works like another.');
    add('');
    add('I am not a developer. Explain things plainly. For anything that needs a browser step in Azure, SharePoint, Entra, IntakeQ, or a funding source website, give exact click-by-click steps.');
    add('');
    add('## How to work');
    add('');
    add('1. Read everything below first. Where an answer is missing, marked not sure, or contradicts another, **ask me before you build**. Do not guess billing rules.');
    add('2. Propose a plan of **at most five steps**. Each step must deliver something that works on its own and can be reviewed. Say what a person can do after each step that they could not do before.');
    add('3. Start with the funding source I name as the first priority (see "Priorities and success"). Build it all the way through, then repeat for the next one. Do not build a generic engine for sources I have not described.');
    add('4. Show me what is possible and what must stay by hand for each funding source, using "Feasibility" below, and tell me if my answers change that.');
    add('5. One pull request per step into `main`. Do not merge it. I will say when to merge.');
    add('');
    add('## Project rules (do not break these)');
    add('');
    add('- Hosting is Azure Static Web Apps, deployed from `main` by the GitHub Action. Sign-in is Microsoft Entra. Only @emmeco.org accounts are allowed. Admin access is the Emme Admin group or the EMME_ADMIN_EMAILS setting (see `api/src/auth.ts`).');
    add('- The admin dashboard is a React app in `admin-ui/` that builds into `admin/`. The built files are committed, so run `npm run build` in `admin-ui` before committing UI changes.');
    add('- The API is Azure Functions in `api/`. TypeScript source is in `api/src` and compiles to `api/dist`, which is also committed, so run `npm run build` in `api` before committing. Run `npm test` in `api`. All tests must pass. Add or update tests for new API behavior.');
    add('- The Billing screen already lists invoices read live from IntakeQ (`api/src/intakeQService.ts`, `admin-ui/src/components/tiles/BillingTile.tsx`). Extend it rather than replacing it.');
    add('- Data sources: SharePoint on the Emme team site through Microsoft Graph with the app\'s own credentials; Azure Blob Storage (account emmereportstorage) for schedules, messages, and similar stored data. Do not change the fixed SharePoint folder paths in the code without asking me.');
    add('- Hosting limits that shape the design: the API on Static Web Apps only answers web requests. It cannot run on a timer by itself. For "keeps checking", use a scheduled Power Automate flow that calls the API (the project already uses this pattern with the automation key), or propose something else and explain why.');
    add('- Never put secrets, keys, passwords, or connection strings in code or in chat. Tell me which Azure application setting to add and give click-by-click steps. Never ask me to type a funding source password into chat.');
    add('- Keep changes small and focused.');
    add('- Client billing information is sensitive health related information. Do not write client names, birth dates, member numbers, or diagnoses to logs. Only show them to the administrators who need them.');
    add('- A person must approve anything that is sent to a funding source or that changes money, unless I clearly say otherwise in my answers below. Keep a history of who did what to each invoice and when.');
    add('- Anything that cannot be automated must still appear as a to-do item with an owner, a due date, and a reminder, so nothing is lost.');
    add('');
    add('## Billing overall: what I told you');
    add('');
    billingQuestions_1.GENERAL_SECTIONS.forEach(function (section) {
        add('### ' + section.title);
        add('');
        billingQuestions_1.GENERAL_QUESTIONS.filter(function (q) { return q.section === section.id; }).forEach(function (question) {
            add('**' + question.text + '**');
            add('');
            add(renderAnswer(question, data.general[question.id]));
            add('');
        });
    });
    const generalFiles = files.filter(function (file) { return file.scope === 'general'; });
    if (generalFiles.length) {
        add('Files for billing overall (attach these to the chat):');
        generalFiles.forEach(function (file) { add(fileLine(file)); });
        add('');
    }
    add('## Funding sources: what I told you');
    add('');
    if (!data.sources.length) {
        add('_No funding sources have been added yet. Ask me which ones to start with._');
        add('');
    }
    const assessments = data.sources.map(assessSource);
    data.sources.forEach(function (source, index) {
        const assessment = assessments[index];
        add('### ' + (source.name || 'Unnamed funding source') + ' (' + kindLabel(source.kind) + ')');
        add('');
        add('**Feasibility from my answers:** sending bills is ' + levelLabel(assessment.sending.level) + '; finding out about payment is ' + levelLabel(assessment.checking.level) + '.');
        add('');
        assessment.sending.notes.concat(assessment.checking.notes).forEach(function (note, i, all) {
            if (all.indexOf(note) === i)
                add('- ' + note);
        });
        assessment.blockers.forEach(function (text) { add('- **Blocked:** ' + text); });
        assessment.cautions.forEach(function (text) { add('- **Caution:** ' + text); });
        assessment.toConfirm.forEach(function (text) { add('- **To confirm:** ' + text); });
        add('');
        const relevant = (0, billingQuestions_1.questionsForKind)(source.kind);
        billingQuestions_1.SOURCE_SECTIONS.forEach(function (section) {
            const inSection = relevant.filter(function (q) { return q.section === section.id; });
            if (!inSection.length)
                return;
            add('#### ' + section.title);
            add('');
            inSection.forEach(function (question) {
                add('**' + question.text + '**');
                add('');
                add(renderAnswer(question, source.answers[question.id]));
                add('');
            });
        });
        const own = files.filter(function (file) { return file.scope === source.id; });
        if (own.length) {
            add('Files for ' + (source.name || 'this funding source') + ' (attach these to the chat):');
            own.forEach(function (file) { add(fileLine(file)); });
            add('');
        }
    });
    add('## Feasibility: the connection options and their limits');
    add('');
    add('Use this when deciding what to build for each funding source.');
    add('');
    billingQuestions_1.CONNECTION_GUIDE.forEach(function (option) {
        add('- **' + option.title + '** (' + levelLabel(option.level) + '). ' + option.works + ' Limits: ' + option.blocked + ' Needs: ' + option.needs);
    });
    add('');
    add('## Priorities and success');
    add('');
    const first = data.general['g_first_source'];
    const success = data.general['g_success'];
    const deadline = data.general['g_deadline'];
    add('- First funding source to automate: ' + (hasContent(first) ? first.text.trim() || 'see answer above' : '_not answered. Ask me._'));
    add('- Success looks like: ' + (hasContent(success) ? success.text.trim() || 'see answer above' : '_not answered. Ask me._'));
    add('- Needed by: ' + (hasContent(deadline) ? deadline.text.trim() || 'see answer above' : '_not answered. Ask me._'));
    add('');
    add('## Open questions to ask me before building');
    add('');
    if (!unknownCount) {
        add('_Every question has an answer. Still confirm anything above that seems unclear._');
    }
    else {
        add('There are ' + unknownCount + ' unanswered or "not sure" items. Ask me about the ones that block your first step. Do not guess the rest; list them as assumptions for me to approve.');
        add('');
        let current = '';
        missing.forEach(function (item) {
            if (item.where !== current) {
                current = item.where;
                add('**' + current + '**');
            }
            add('- ' + (item.reason === 'not-sure' ? '(not sure) ' : '') + item.question);
        });
    }
    add('');
    add('## What to give me back');
    add('');
    add('1. A short plain-language summary of what you understood, and anything that looks wrong or risky.');
    add('2. The one-to-five step plan, with what I will be able to do after each step.');
    add('3. For each funding source: what will be automatic, what will be partly automatic, and what stays by hand.');
    add('4. The exact questions you still need me to answer, and any steps I must do in a browser, with click-by-click instructions.');
    add('5. Do not start building until I approve the plan.');
    add('');
    return lines.join('\n');
}
