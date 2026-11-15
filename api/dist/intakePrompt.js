"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.questionsForKind = questionsForKind;
exports.assessAll = assessAll;
exports.levelLabel = levelLabel;
exports.missingItems = missingItems;
exports.buildPrompt = buildPrompt;
function questionsForKind(topic, kind) {
    if (!topic.group)
        return [];
    return topic.group.questions.filter(function (q) { return !q.kinds || q.kinds.indexOf(kind) !== -1; });
}
/** What is possible for each group entry (for example, each funding source). Empty when the topic has no group. */
function assessAll(topic, data) {
    const assess = topic.assess;
    return assess ? data.sources.map(function (item) { return assess(item); }) : [];
}
function hasContent(answer) {
    return !!answer && (answer.choice.length > 0 || answer.text.trim().length > 0);
}
function kindLabel(topic, kind) {
    const found = topic.group ? topic.group.kinds.find(function (item) { return item.id === kind; }) : undefined;
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
function missingItems(topic, data) {
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
    scan(topic.group ? topic.title + ' overall' : topic.title, topic.questions, data.general);
    const group = topic.group;
    if (group) {
        data.sources.forEach(function (item) {
            scan(item.name || group.singular, questionsForKind(topic, item.kind), item.answers);
        });
    }
    return out;
}
function fileLine(topic, file) {
    const kind = topic.fileKinds.find(function (item) { return item.id === file.kind; });
    return '- `' + file.name + '`: ' + (kind ? kind.label : file.kind) + (file.note ? '. ' + file.note : '');
}
function priority(data, id) {
    const answer = id ? data.general[id] : undefined;
    return hasContent(answer) ? answer.text.trim() || 'see answer above' : '_not answered. Ask me._';
}
function buildPrompt(topic, data, files, now) {
    const lines = [];
    const add = function (text) { lines.push(text); };
    const when = now.toISOString().slice(0, 10);
    const missing = missingItems(topic, data);
    const unknownCount = missing.length;
    const group = topic.group;
    const noun = group ? group.singular : 'process';
    add('# Build request: ' + topic.title + ' for the Emme Administration Dashboard');
    add('');
    add('_Generated ' + when + ' from the ' + topic.title + ' setup answers. Last saved by ' + (data.updatedBy || 'nobody yet') + (data.updatedAt ? ' on ' + data.updatedAt.slice(0, 10) : '') + '._');
    add('');
    add('## What I want');
    add('');
    topic.goal.forEach(function (paragraph) { add(paragraph); add(''); });
    add('I am not a developer. Explain things plainly. For anything that needs a browser step in Azure, SharePoint, Entra, IntakeQ, or another website, give exact click-by-click steps.');
    add('');
    add('## How to work');
    add('');
    add('1. Read everything below first. Where an answer is missing, marked not sure, or contradicts another, **ask me before you build**. Do not guess Emme\'s rules.');
    add('2. Propose a plan of **at most five steps**. Each step must deliver something that works on its own and can be reviewed. Say what a person can do after each step that they could not do before.');
    if (group) {
        add('3. Start with the ' + group.singular + ' I name as the first priority (see "Priorities and success"). Build it all the way through, then repeat for the next one. Do not build a generic engine for ' + group.plural + ' I have not described.');
    }
    else {
        add('3. Start with the part that saves the most time or prevents the most mistakes (see "Priorities and success"). Build it all the way through before starting the next. Do not build things I have not described.');
    }
    add('4. Show me what is possible and what must stay by hand for each ' + noun + ', using "Feasibility" below, and tell me if my answers change that.');
    add('5. One pull request per step into `main`. Do not merge it. I will say when to merge.');
    add('');
    add('## Project rules (do not break these)');
    add('');
    add('- Hosting is Azure Static Web Apps, deployed from `main` by the GitHub Action. Sign-in is Microsoft Entra. Only @emmeco.org accounts are allowed. Admin access is the Emme Admin group or the EMME_ADMIN_EMAILS setting (see `api/src/auth.ts`).');
    add('- The admin dashboard is a React app in `admin-ui/` that builds into `admin/`. The built files are committed, so run `npm run build` in `admin-ui` before committing UI changes.');
    add('- The API is Azure Functions in `api/`. TypeScript source is in `api/src` and compiles to `api/dist`, which is also committed, so run `npm run build` in `api` before committing. Run `npm test` in `api`. All tests must pass. Add or update tests for new API behavior.');
    topic.existing.forEach(function (text) { add('- ' + text); });
    add('- Data sources: SharePoint on the Emme team site through Microsoft Graph with the app\'s own credentials; Azure Blob Storage (account emmereportstorage) for schedules, messages, and similar stored data. Do not change the fixed SharePoint folder paths in the code without asking me. Do not change the Resource Hub page or the guidebook links.');
    add('- Hosting limits that shape the design: the API on Static Web Apps only answers web requests. It cannot run on a timer by itself. For anything that must keep checking or send reminders, use a scheduled Power Automate flow that calls the API (the project already uses this pattern with the automation key), or propose something else and explain why.');
    add('- Never put secrets, keys, passwords, or connection strings in code or in chat. Tell me which Azure application setting to add and give click-by-click steps. Never ask me to type a password into chat.');
    add('- Keep changes small and focused.');
    topic.rules.forEach(function (text) { add('- ' + text); });
    add('');
    add('## ' + (group ? topic.title + ' overall' : topic.title) + ': what I told you');
    add('');
    topic.sections.forEach(function (section) {
        add('### ' + section.title);
        add('');
        topic.questions.filter(function (q) { return q.section === section.id; }).forEach(function (question) {
            add('**' + question.text + '**');
            add('');
            add(renderAnswer(question, data.general[question.id]));
            add('');
        });
    });
    const generalFiles = files.filter(function (file) { return file.scope === 'general'; });
    if (generalFiles.length) {
        add('Files for ' + topic.title.toLowerCase() + ' overall (attach these to the chat):');
        generalFiles.forEach(function (file) { add(fileLine(topic, file)); });
        add('');
    }
    if (group) {
        add('## ' + group.plural.charAt(0).toUpperCase() + group.plural.slice(1) + ': what I told you');
        add('');
        if (!data.sources.length) {
            add('_No ' + group.plural + ' have been added yet. Ask me which ones to start with._');
            add('');
        }
        const assessments = assessAll(topic, data);
        data.sources.forEach(function (item, index) {
            const assessment = assessments[index];
            add('### ' + (item.name || 'Unnamed ' + group.singular) + ' (' + kindLabel(topic, item.kind) + ')');
            add('');
            if (assessment) {
                add('**Feasibility from my answers:** ' + assessment.rows.map(function (row) { return row.label.toLowerCase() + ' is ' + levelLabel(row.level); }).join('; ') + '.');
                add('');
                const notes = [];
                assessment.rows.forEach(function (row) { row.notes.forEach(function (note) { if (notes.indexOf(note) === -1)
                    notes.push(note); }); });
                notes.forEach(function (note) { add('- ' + note); });
                assessment.blockers.forEach(function (text) { add('- **Blocked:** ' + text); });
                assessment.cautions.forEach(function (text) { add('- **Caution:** ' + text); });
                assessment.toConfirm.forEach(function (text) { add('- **To confirm:** ' + text); });
                add('');
            }
            const relevant = questionsForKind(topic, item.kind);
            group.sections.forEach(function (section) {
                const inSection = relevant.filter(function (q) { return q.section === section.id; });
                if (!inSection.length)
                    return;
                add('#### ' + section.title);
                add('');
                inSection.forEach(function (question) {
                    add('**' + question.text + '**');
                    add('');
                    add(renderAnswer(question, item.answers[question.id]));
                    add('');
                });
            });
            const own = files.filter(function (file) { return file.scope === item.id; });
            if (own.length) {
                add('Files for ' + (item.name || 'this ' + group.singular) + ' (attach these to the chat):');
                own.forEach(function (file) { add(fileLine(topic, file)); });
                add('');
            }
        });
    }
    add('## Feasibility: the connection options and their limits');
    add('');
    add('Use this when deciding what to build.');
    add('');
    topic.guide.forEach(function (option) {
        add('- **' + option.title + '** (' + levelLabel(option.level) + '). ' + option.works + ' Limits: ' + option.blocked + ' Needs: ' + option.needs);
    });
    add('');
    add('## Priorities and success');
    add('');
    if (topic.firstId)
        add('- ' + (topic.firstLabel || 'Do first') + ': ' + priority(data, topic.firstId));
    add('- Success looks like: ' + priority(data, 'g_success'));
    add('- Needed by: ' + priority(data, 'g_deadline'));
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
    add('3. For each ' + noun + ': what will be automatic, what will be partly automatic, and what stays by hand.');
    add('4. The exact questions you still need me to answer, and any steps I must do in a browser, with click-by-click instructions.');
    add('5. Do not start building until I approve the plan.');
    add('');
    return lines.join('\n');
}
