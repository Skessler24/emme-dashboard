"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.LISTS = void 0;
exports.employeeFolderName = employeeFolderName;
exports.listItems = listItems;
exports.createItem = createItem;
exports.updateItem = updateItem;
exports.listFolder = listFolder;
exports.uploadEmployeeFile = uploadEmployeeFile;
exports.readJsonFile = readJsonFile;
exports.readLatestScheduleSummaries = readLatestScheduleSummaries;
exports.onboardingFolders = onboardingFolders;
const graphClient_1 = require("./graphClient");
/**
 * SharePoint lists used by the admin dashboard. Create these on the Emme
 * SharePoint site (display names must match). Title is required by SharePoint.
 *
 * OnboardingChecklist: Title, EmployeeEmail, EmployeeName, Item, Status, DueDate
 * Referrals: Title, ReferrerName, ReceivedOn, Status, AssignedTo, Notes
 * Trainings: Title, TrainingName, Required, DueDate, Description
 * StaffTrainingStatus: Title, StaffEmail, StaffName, TrainingName, Status, CompletedOn
 * AdminAlerts: Title, Message, Active (Yes/No), StartsOn, EndsOn
 * StaffMessages: Title, Message (multiline), FromEmail, FromName, Reply (multiline), RepliedBy, RepliedAt, Status
 * Mileage: Title, StaffEmail, StaffName, TripDate, Miles, Purpose, Status
 * Cancellations: Title, StaffEmail, StaffName, AppointmentDate, Status, ServiceName, ClientName, ExternalId, Source
 * DocumentationCompliance: Title, StaffEmail, StaffName, WeekOf, OpenNotes, OpenEvals, OpenReports
 *
 * Employee files live in the site document library:
 *   /Employees/<EmployeeName>/Onboarding/CurrentDocuments/
 *   /Employees/<EmployeeName>/Onboarding/ArchivedDocuments/
 *   /Employees/<EmployeeName>/Current Availability/Schedules/<Name> Schedule.xlsx
 *   /Employees/<EmployeeName>/Mileage/<Name> Mileage.xlsx
 *
 * The workbook is overwritten when that provider saves Schedule & Caseload.
 * This service reads those workbooks for the admin openings view.
 */
exports.LISTS = {
    onboarding: 'OnboardingChecklist',
    referrals: 'Referrals',
    trainings: 'Trainings',
    staffTraining: 'StaffTrainingStatus',
    alerts: 'AdminAlerts',
    messages: 'StaffMessages',
    mileage: 'Mileage',
    cancellations: 'Cancellations',
    compliance: 'DocumentationCompliance',
};
function cleanFields(fields) {
    const out = {};
    if (!fields)
        return out;
    Object.keys(fields).forEach(function (key) {
        if (key.charAt(0) === '@')
            return;
        out[key] = fields[key];
    });
    return out;
}
function employeeFolderName(name) {
    const cleaned = String(name || '')
        .replace(/[\\/:*?"<>|#%]/g, ' ')
        .replace(/\s+/g, ' ')
        .replace(/^\.+|\.+$/g, '')
        .trim();
    if (!cleaned || cleaned === '.' || cleaned === '..') {
        throw new graphClient_1.GraphError(400, 'Employee name is required');
    }
    return cleaned.slice(0, 80);
}
function encodeDrivePath(segments) {
    return segments.map(function (segment) { return encodeURIComponent(segment); }).join('/');
}
async function listPath(listName) {
    const siteId = await (0, graphClient_1.getSharePointSiteId)();
    return '/sites/' + siteId + '/lists/' + encodeURIComponent(listName);
}
async function listItems(listName) {
    let next = (await listPath(listName)) + '/items?$expand=fields&$top=200';
    const items = [];
    for (let page = 0; page < 10 && next; page++) {
        const data = await (0, graphClient_1.graphGet)(next);
        const value = (data && data.value) || [];
        value.forEach(function (item) {
            items.push({
                id: String(item.id || ''),
                createdAt: String(item.createdDateTime || ''),
                fields: cleanFields(item.fields),
            });
        });
        next = data && data['@odata.nextLink'] ? String(data['@odata.nextLink']) : null;
    }
    return items;
}
async function createItem(listName, fields) {
    const path = await listPath(listName);
    const created = await (0, graphClient_1.graphSend)(path + '/items', { method: 'POST', jsonBody: { fields } });
    return {
        id: String(created && created.id || ''),
        createdAt: String(created && created.createdDateTime || ''),
        fields: cleanFields(created && created.fields),
    };
}
async function updateItem(listName, itemId, fields) {
    const path = await listPath(listName);
    await (0, graphClient_1.graphSend)(path + '/items/' + encodeURIComponent(itemId) + '/fields', {
        method: 'PATCH',
        jsonBody: fields,
    });
}
async function listFolder(segments) {
    const siteId = await (0, graphClient_1.getSharePointSiteId)();
    const path = '/sites/' + siteId + '/drive/root:/' + encodeDrivePath(segments) + ':/children?$select=name,lastModifiedDateTime,folder,file';
    try {
        const data = await (0, graphClient_1.graphGet)(path);
        return ((data && data.value) || []).map(function (item) {
            return {
                name: String(item.name || ''),
                lastModified: String(item.lastModifiedDateTime || ''),
                isFolder: !!item.folder,
            };
        });
    }
    catch (err) {
        if (err instanceof graphClient_1.GraphError && (err.status === 404 || /itemNotFound/i.test(err.message)))
            return [];
        throw err;
    }
}
async function uploadEmployeeFile(segments, fileName, bytes, contentType) {
    const safeName = fileName.replace(/[\\/]/g, '').replace(/\.\./g, '').trim();
    if (!safeName)
        throw new graphClient_1.GraphError(400, 'File name is required');
    const siteId = await (0, graphClient_1.getSharePointSiteId)();
    const full = segments.concat([safeName]);
    const path = '/sites/' + siteId + '/drive/root:/' + encodeDrivePath(full) + ':/content';
    await (0, graphClient_1.graphSend)(path, { method: 'PUT', rawBody: bytes, contentType });
    return { name: safeName, path: '/' + full.join('/') };
}
async function readJsonFile(segments) {
    const siteId = await (0, graphClient_1.getSharePointSiteId)();
    const path = '/sites/' + siteId + '/drive/root:/' + encodeDrivePath(segments) + ':/content';
    const text = await (0, graphClient_1.graphGetText)(path);
    return JSON.parse(text);
}
const SUMMARY_FILE = /^schedule-summary-(\d{4}-\d{2}-\d{2})\.json$/i;
function unwrapSummary(value) {
    if (!value || typeof value !== 'object' || Array.isArray(value))
        return value;
    const record = value;
    if (!record.openings && record.summary && typeof record.summary === 'object')
        return record.summary;
    return value;
}
const WORKBOOK_FILE = /schedule\.xlsx$/i;
async function readFileBytes(segments) {
    const siteId = await (0, graphClient_1.getSharePointSiteId)();
    const path = '/sites/' + siteId + '/drive/root:/' + encodeDrivePath(segments) + ':/content';
    return (0, graphClient_1.graphGetBytes)(path);
}
/**
 * Reads each provider's living schedule workbook. Older JSON summaries are
 * used only when the workbook is not there yet.
 */
async function readLatestScheduleSummaries() {
    const employees = await listFolder(['Employees']);
    const summaries = [];
    for (let i = 0; i < employees.length; i++) {
        const employee = employees[i];
        if (!employee.isFolder || !employee.name)
            continue;
        const folder = ['Employees', employee.name, 'Current Availability', 'Schedules'];
        let files = [];
        try {
            files = await listFolder(folder);
        }
        catch (_) {
            continue;
        }
        const workbook = files.find(function (file) { return WORKBOOK_FILE.test(file.name); });
        if (workbook) {
            try {
                const bytes = await readFileBytes(folder.concat([workbook.name]));
                const parsed = await require('./scheduleWorkbook').openingsFromWorkbook(bytes, employee.name);
                summaries.push({
                    employee: employee.name,
                    fileName: workbook.name,
                    path: '/' + folder.concat([workbook.name]).join('/'),
                    summary: parsed,
                });
                continue;
            }
            catch (_) {
                continue;
            }
        }
        const summariesForEmployee = files
            .filter(function (file) { return SUMMARY_FILE.test(file.name); })
            .sort(function (a, b) { return a.name < b.name ? 1 : -1; });
        const latest = summariesForEmployee[0];
        if (!latest)
            continue;
        try {
            const summary = unwrapSummary(await readJsonFile(folder.concat([latest.name])));
            summaries.push({
                employee: employee.name,
                fileName: latest.name,
                path: '/' + folder.concat([latest.name]).join('/'),
                summary,
            });
        }
        catch (_) {
            continue;
        }
    }
    return summaries;
}
function onboardingFolders(employeeName) {
    const folder = employeeFolderName(employeeName);
    return {
        current: ['Employees', folder, 'Onboarding', 'CurrentDocuments'],
        archived: ['Employees', folder, 'Onboarding', 'ArchivedDocuments'],
    };
}
