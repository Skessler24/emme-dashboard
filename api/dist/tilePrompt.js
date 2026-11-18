"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.seededAnswers = seededAnswers;
exports.missingItems = missingItems;
exports.buildTilePrompt = buildTilePrompt;
function hasContent(answer) {
    return !!answer && (answer.choice.length > 0 || answer.text.trim().length > 0);
}
/** The first draft of every answer, taken from the requirements document and the interview sheet. */
function seededAnswers(tile) {
    const out = {};
    tile.questions.forEach(function (question) {
        const seed = question.seed;
        if (!seed)
            return;
        const valid = (question.options || []).map(function (option) { return option.id; });
        const choice = (seed.choice || []).filter(function (id) { return valid.indexOf(id) !== -1; });
        out[question.id] = { choice, text: seed.text || '', unsure: seed.unsure === true };
    });
    return out;
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
function missingItems(tile, answers) {
    const out = [];
    tile.questions.forEach(function (question) {
        const answer = answers[question.id];
        if (answer && answer.unsure)
            out.push({ question: question.text, reason: 'not-sure' });
        else if (!hasContent(answer))
            out.push({ question: question.text, reason: 'unanswered' });
    });
    return out;
}
function fileLine(tile, file) {
    const kind = tile.fileKinds.find(function (item) { return item.id === file.kind; });
    return '- `' + file.name + '`: ' + (kind ? kind.label : file.kind) + (file.note ? '. ' + file.note : '');
}
function buildTilePrompt(tile, saved, files, now) {
    const answers = saved ? saved.answers : seededAnswers(tile);
    const lines = [];
    const add = function (text) { lines.push(text); };
    const missing = missingItems(tile, answers);
    add('# Build request: ' + tile.title + ' for the Emme Administration Dashboard');
    add('');
    add('_Generated ' + now.toISOString().slice(0, 10) + ' from the ' + tile.title + ' setup answers. '
        + (saved ? 'Last saved by ' + (saved.updatedBy || 'nobody') + ' on ' + saved.updatedAt.slice(0, 10) + '.' : 'Nobody has reviewed or saved these answers yet. They are the first draft taken from the requirements document and the boss interview sheet.') + '_');
    add('');
    add('## What I want');
    add('');
    add(tile.goal);
    add('');
    add('Many answers began as notes from my bosses, written quickly. If an answer contradicts another one, or looks like a rough note rather than a decision, **ask me before you build**.');
    add('');
    add('I am not a developer. Explain things plainly. For anything that needs a browser step in Azure, SharePoint, Entra, IntakeQ, Zoom, or Microsoft Teams, give exact click-by-click steps.');
    add('');
    add('## How to work');
    add('');
    add('1. Read everything below first. Where an answer is missing, marked not sure, or contradicts another, ask me before you build. Do not guess Emme\'s rules.');
    add('2. Propose a plan of **at most five steps**. Each step must deliver something that works on its own and can be reviewed. Say what a person can do after each step that they could not do before.');
    add('3. One pull request per step into `main`. Do not merge it. I will say when to merge.');
    add('4. Do not start building until I approve the plan.');
    add('');
    add('## Project rules (do not break these)');
    add('');
    add('- Hosting is Azure Static Web Apps, deployed from `main` by the GitHub Action. Sign-in is Microsoft Entra. Only @emmeco.org accounts are allowed. Admin access is the Emme Admin group or the EMME_ADMIN_EMAILS setting (see `api/src/auth.ts`).');
    add('- The admin dashboard is a React app in `admin-ui/` that builds into `admin/`. The built files are committed, so run `npm run build` in `admin-ui` before committing UI changes.');
    add('- The API is Azure Functions in `api/`. TypeScript source is in `api/src` and compiles to `api/dist`, which is also committed, so run `npm run build` in `api` before committing. Run `npm test` in `api`. All tests must pass. Add or update tests for new API behavior.');
    add('- Data sources: SharePoint on the Emme team site through Microsoft Graph with the app\'s own credentials; Azure Blob Storage (account emmereportstorage) for schedules, messages, and similar stored data. Do not change the fixed SharePoint folder paths in the code without asking me.');
    add('- Never put secrets, keys, passwords, or connection strings in code or in chat. Tell me which Azure application setting to add and give click-by-click steps.');
    add('- Keep changes small and focused.');
    add('- Information about children and families is sensitive. Do not write names, birth dates, or notes to logs. Only show them to the administrators who need them.');
    add('- A person must approve anything sent to a family, a service coordinator, or an outside company, unless my answers below clearly say it may send by itself.');
    add('- Anything that cannot be automated must still appear as a to-do item with an owner, a due date, and a reminder, so nothing is lost.');
    add('');
    add('## What already exists for this tile');
    add('');
    tile.existing.forEach(function (text) { add('- ' + text); });
    add('');
    add('## What I told you');
    add('');
    tile.sections.forEach(function (section) {
        const inSection = tile.questions.filter(function (question) { return question.section === section.id; });
        if (!inSection.length)
            return;
        add('### ' + section.title);
        add('');
        inSection.forEach(function (question) {
            add('**' + question.text + '**');
            add('');
            add(renderAnswer(question, answers[question.id]));
            add('');
        });
    });
    if (files.length) {
        add('Files (attach these to the chat):');
        files.forEach(function (file) { add(fileLine(tile, file)); });
        add('');
    }
    add('## Open questions to ask me before building');
    add('');
    if (!missing.length) {
        add('_Every question has an answer. Still confirm anything above that seems unclear._');
    }
    else {
        add('There are ' + missing.length + ' unanswered or "not sure" items. Ask me about the ones that block your first step. Do not guess the rest; list them as assumptions for me to approve.');
        add('');
        missing.forEach(function (item) {
            add('- ' + (item.reason === 'not-sure' ? '(not sure) ' : '') + item.question);
        });
    }
    add('');
    add('## What to give me back');
    add('');
    add('1. A short plain-language summary of what you understood, and anything that looks wrong or risky.');
    add('2. The one-to-five step plan, with what I will be able to do after each step.');
    add('3. What will be automatic, what will be partly automatic, and what stays by hand.');
    add('4. The exact questions you still need me to answer, and any steps I must do in a browser, with click-by-click instructions.');
    add('5. Do not start building until I approve the plan.');
    add('');
    return lines.join('\n');
}
