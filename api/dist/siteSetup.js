"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.FOLDER_CHECKS = exports.LIST_SCHEMAS = void 0;
exports.checkSetup = checkSetup;
exports.createMissingLists = createMissingLists;
const graphClient_1 = require("./graphClient");
const providerSetup_1 = require("./providerSetup");
const sharepointService_1 = require("./sharepointService");
const config_1 = require("./config");
const text = (...names) => names.map(function (name) { return { name, kind: 'text' }; });
const multiline = (...names) => names.map(function (name) { return { name, kind: 'multiline' }; });
const numbers = (...names) => names.map(function (name) { return { name, kind: 'number' }; });
const yesNo = (...names) => names.map(function (name) { return { name, kind: 'boolean' }; });
/** Title already exists on every new list. Dates are plain text so any format the dashboard or Power Automate writes is accepted. */
exports.LIST_SCHEMAS = [
    { list: sharepointService_1.LISTS.onboarding, columns: text('EmployeeEmail', 'EmployeeName', 'Item', 'Status', 'DueDate') },
    { list: sharepointService_1.LISTS.referrals, columns: text('ReferrerName', 'ReceivedOn', 'Status', 'AssignedTo').concat(multiline('Notes')) },
    { list: sharepointService_1.LISTS.trainings, columns: text('TrainingName', 'DueDate').concat(yesNo('Required'), multiline('Description')) },
    { list: sharepointService_1.LISTS.staffTraining, columns: text('StaffEmail', 'StaffName', 'TrainingName', 'Status', 'CompletedOn') },
    { list: sharepointService_1.LISTS.alerts, columns: multiline('Message').concat(yesNo('Active'), text('StartsOn', 'EndsOn')) },
    { list: sharepointService_1.LISTS.messages, columns: multiline('Message', 'Reply').concat(text('FromEmail', 'FromName', 'RepliedBy', 'RepliedAt', 'Status')) },
    { list: sharepointService_1.LISTS.mileage, columns: text('StaffEmail', 'StaffName', 'TripDate').concat(numbers('Miles'), text('Purpose', 'Status')) },
    { list: sharepointService_1.LISTS.cancellations, columns: text('StaffEmail', 'StaffName', 'AppointmentDate', 'Status', 'ServiceName', 'ClientName', 'ExternalId', 'Source') },
    { list: sharepointService_1.LISTS.compliance, columns: text('StaffEmail', 'StaffName', 'WeekOf').concat(numbers('OpenNotes', 'OpenEvals', 'OpenReports')) },
];
exports.FOLDER_CHECKS = [
    { label: 'Provider Schedules', path: providerSetup_1.SCHEDULE_ROOT, usedFor: 'Schedule & Caseload sync' },
    { label: 'Mileage', path: providerSetup_1.MILEAGE_ROOT, usedFor: 'Mileage logs and the nightly sync' },
    { label: 'RS_Observation Process', path: providerSetup_1.RS_OBSERVATION_ROOT, usedFor: 'Reflective Supervision tile' },
    { label: 'Employees', path: ['Employees'], usedFor: 'Onboarding documents' },
];
function encodePath(segments) {
    return segments.map(encodeURIComponent).join('/');
}
function notFound(err) {
    return err instanceof graphClient_1.GraphError && (err.status === 404 || err.status === 403 || err.status === 401);
}
async function folderExists(siteId, segments) {
    if (!siteId)
        return false;
    try {
        const item = await (0, graphClient_1.graphGet)('/sites/' + siteId + '/drive/root:/' + encodePath(segments) + '?$select=id,folder');
        return !!(item && item.id);
    }
    catch (err) {
        if (notFound(err))
            return false;
        throw err;
    }
}
function columnBody(spec) {
    const body = { name: spec.name, displayName: spec.name };
    if (spec.kind === 'multiline')
        body.text = { allowMultipleLines: true };
    else if (spec.kind === 'number')
        body.number = {};
    else if (spec.kind === 'boolean')
        body.boolean = {};
    else
        body.text = {};
    return body;
}
async function readList(siteId, listName) {
    try {
        const data = await (0, graphClient_1.graphGet)('/sites/' + siteId + '/lists/' + encodeURIComponent(listName) + '?$select=id,displayName&$expand=columns($select=name)');
        const columns = ((data && data.columns) || []).map(function (column) { return String(column.name || '').toLowerCase(); });
        return { id: String(data.id), columns };
    }
    catch (err) {
        if (err instanceof graphClient_1.GraphError && err.status === 404)
            return null;
        throw err;
    }
}
async function checkSetup() {
    const cfg = (0, config_1.getConfig)();
    const status = {
        graphConfigured: (0, config_1.graphConfigured)(),
        sharePointSettings: { hostname: cfg.sharePointHostname, sitePath: cfg.sharePointSitePath },
        site: null,
        siteError: '',
        lists: [],
        folders: [],
        usingPersonalOneDrive: false,
        nextSteps: [],
        ready: false,
    };
    if (!status.graphConfigured) {
        status.nextSteps.push('Set GRAPH_CLIENT_SECRET (or AZURE_CLIENT_SECRET) in the Azure app settings.');
        return status;
    }
    const sites = await (0, graphClient_1.knownSites)();
    let siteId = '';
    if (!cfg.sharePointHostname || !cfg.sharePointSitePath) {
        status.siteError = 'SHAREPOINT_HOSTNAME and SHAREPOINT_SITE_PATH are not set.';
        status.nextSteps.push('In Azure, set SHAREPOINT_HOSTNAME (for example emmeco031.sharepoint.com) and SHAREPOINT_SITE_PATH (for example /sites/Emme).');
    }
    else {
        try {
            siteId = await (0, graphClient_1.getSharePointSiteId)();
            const info = await (0, graphClient_1.graphGet)('/sites/' + siteId + '?$select=displayName,webUrl');
            status.site = { name: String(info.displayName || ''), webUrl: String(info.webUrl || '') };
        }
        catch (err) {
            status.siteError = err instanceof Error ? err.message : 'The SharePoint site could not be reached.';
            status.nextSteps.push('The dashboard could not open ' + cfg.sharePointHostname + cfg.sharePointSitePath + '. Check the spelling, and that the app has Sites.ReadWrite.All with admin consent.');
        }
    }
    if (siteId) {
        for (let i = 0; i < exports.LIST_SCHEMAS.length; i++) {
            const schema = exports.LIST_SCHEMAS[i];
            const found = await readList(siteId, schema.list);
            if (!found) {
                status.lists.push({ name: schema.list, exists: false, missingColumns: schema.columns.map(function (c) { return c.name; }) });
            }
            else {
                status.lists.push({
                    name: schema.list,
                    exists: true,
                    missingColumns: schema.columns.filter(function (c) { return found.columns.indexOf(c.name.toLowerCase()) === -1; }).map(function (c) { return c.name; }),
                });
            }
        }
        const lacking = status.lists.filter(function (l) { return !l.exists || l.missingColumns.length; });
        if (lacking.length)
            status.nextSteps.push('Press "Create missing lists" below to add ' + lacking.length + ' list' + (lacking.length === 1 ? '' : 's') + ' or column set' + (lacking.length === 1 ? '' : 's') + '.');
    }
    for (let i = 0; i < exports.FOLDER_CHECKS.length; i++) {
        const check = exports.FOLDER_CHECKS[i];
        const onShared = await folderExists(siteId, check.path);
        const onPersonal = await folderExists(sites.personal, check.path);
        status.folders.push({
            label: check.label,
            path: '/' + check.path.join('/'),
            usedFor: check.usedFor,
            onSharedSite: onShared,
            onPersonalOneDrive: onPersonal,
        });
    }
    status.usingPersonalOneDrive = status.folders.some(function (f) { return !f.onSharedSite && f.onPersonalOneDrive; });
    const moved = status.folders.filter(function (f) { return !f.onSharedSite && f.onPersonalOneDrive; });
    if (moved.length) {
        status.nextSteps.push('Move these folders from the personal OneDrive into the Emme site Documents library: ' + moved.map(function (f) { return f.path; }).join(', ') + '.');
    }
    const missing = status.folders.filter(function (f) { return !f.onSharedSite && !f.onPersonalOneDrive; });
    if (missing.length) {
        status.nextSteps.push('Not found anywhere yet: ' + missing.map(function (f) { return f.path; }).join(', ') + '. Add them to the Emme site Documents library when you have them.');
    }
    status.ready = !!status.site && !status.lists.some(function (l) { return !l.exists || l.missingColumns.length; }) &&
        status.folders.every(function (f) { return f.onSharedSite; });
    return status;
}
async function createMissingLists() {
    const siteId = await (0, graphClient_1.getSharePointSiteId)();
    const result = { createdLists: [], addedColumns: [] };
    for (let i = 0; i < exports.LIST_SCHEMAS.length; i++) {
        const schema = exports.LIST_SCHEMAS[i];
        const found = await readList(siteId, schema.list);
        if (!found) {
            await (0, graphClient_1.graphSend)('/sites/' + siteId + '/lists', {
                method: 'POST',
                jsonBody: {
                    displayName: schema.list,
                    list: { template: 'genericList' },
                    columns: schema.columns.map(columnBody),
                },
            });
            result.createdLists.push(schema.list);
            continue;
        }
        const missing = schema.columns.filter(function (c) { return found.columns.indexOf(c.name.toLowerCase()) === -1; });
        const added = [];
        for (let j = 0; j < missing.length; j++) {
            await (0, graphClient_1.graphSend)('/sites/' + siteId + '/lists/' + found.id + '/columns', { method: 'POST', jsonBody: columnBody(missing[j]) });
            added.push(missing[j].name);
        }
        if (added.length)
            result.addedColumns.push({ list: schema.list, columns: added });
    }
    return result;
}
