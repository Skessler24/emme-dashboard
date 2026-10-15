"use strict";
/**
 * Runtime configuration for Microsoft Graph and IntakeQ.
 *
 * Set these as Azure Static Web App application settings.
 * They can be Key Vault references, for example:
 *   @Microsoft.KeyVault(SecretUri=https://<vault>.vault.azure.net/secrets/<name>/)
 *
 * Never commit secret values. Client id and tenant id may fall back to the
 * existing Emme app registration; the client secret must come from the environment.
 *
 * Graph application permissions (admin consent) required in addition to the
 * calendar and to-do permissions already used by the dashboard:
 *   GroupMember.Read.All, User.Read.All, Sites.ReadWrite.All
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.getConfig = getConfig;
exports.graphConfigured = graphConfigured;
exports.intakeqConfigured = intakeqConfigured;
exports.sharePointConfigured = sharePointConfigured;
const DEFAULT_CLIENT_ID = 'b1e3bd77-c545-4716-afdc-8b0ec14ddcdf';
const DEFAULT_TENANT_ID = '4da01280-a61e-44cc-b537-9b21b7f17c68';
function env(name) {
    return String(process.env[name] || '').trim();
}
function getConfig() {
    return {
        graphClientId: env('GRAPH_CLIENT_ID') || env('AZURE_CLIENT_ID') || DEFAULT_CLIENT_ID,
        graphClientSecret: env('GRAPH_CLIENT_SECRET') || env('AZURE_CLIENT_SECRET'),
        graphTenantId: env('GRAPH_TENANT_ID') || env('AZURE_TENANT_ID') || DEFAULT_TENANT_ID,
        intakeqApiKey: env('INTAKEQ_API_KEY'),
        sharePointHostname: env('SHAREPOINT_HOSTNAME'),
        sharePointSitePath: env('SHAREPOINT_SITE_PATH'),
        adminGroupName: env('EMME_ADMIN_GROUP_NAME') || 'Emme Admin',
        adminGroupId: env('EMME_ADMIN_GROUP_ID'),
        adminEmails: env('EMME_ADMIN_EMAILS').toLowerCase().split(/[,;\s]+/).filter(Boolean),
        automationKey: env('AUTOMATION_API_KEY'),
        storageConnection: env('AZURE_STORAGE_CONNECTION_STRING'),
        storageContainer: env('AZURE_STORAGE_CONTAINER') || 'emme-reports',
    };
}
function graphConfigured() {
    return !!getConfig().graphClientSecret;
}
function intakeqConfigured() {
    return !!getConfig().intakeqApiKey;
}
function sharePointConfigured() {
    const cfg = getConfig();
    return !!(cfg.sharePointHostname && cfg.sharePointSitePath && cfg.graphClientSecret);
}
