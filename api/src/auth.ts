import { timingSafeEqual } from 'crypto';
import { getConfig, graphConfigured } from './config';
import { GraphError, userIsAdmin } from './graphClient';
import { FnContext, HttpRequest, header, json } from './http';

export type Principal = {
  userId: string;
  name: string;
  email: string;
};

const adminCache = new Map<string, { value: boolean; expires: number }>();

export function readPrincipal(req: HttpRequest): Principal | null {
  const directId = header(req, 'x-ms-client-principal-id');
  const directName = header(req, 'x-ms-client-principal-name');
  const encoded = header(req, 'x-ms-client-principal');
  let parsed: any = null;
  if (encoded) {
    try { parsed = JSON.parse(Buffer.from(encoded, 'base64').toString('utf8')); } catch (_) { parsed = null; }
  }
  const claims = (parsed && parsed.claims) || [];
  let email = '';
  let name = '';
  function consider(value: string, prefer: boolean) {
    const text = String(value || '').trim();
    if (!text) return;
    if (!name && text.indexOf('@') === -1) name = text;
    if (text.indexOf('@') === -1) return;
    if (prefer || !email) email = text.toLowerCase();
  }
  claims.forEach(function (claim: { typ?: string; val?: string }) {
    const typ = String((claim && claim.typ) || '').toLowerCase();
    const val = String((claim && claim.val) || '').trim();
    if (typ.indexOf('name') !== -1 && typ.indexOf('upn') === -1 && val.indexOf('@') === -1) name = val;
    const prefer = typ.indexOf('email') !== -1 || typ.indexOf('preferred_username') !== -1 || typ.endsWith('/upn');
    if (prefer || val.indexOf('@') !== -1) consider(val, prefer);
  });
  consider(directName, directName.indexOf('@') !== -1);
  const userId = directId || (parsed && parsed.userId) || '';
  if (!email && !userId) return null;
  return {
    userId: String(userId || email),
    name: name || email,
    email: email,
  };
}

export function automationAuthorized(req: HttpRequest): boolean {
  const expected = getConfig().automationKey;
  if (!expected) return false;
  const got = header(req, 'x-emme-automation-key');
  const a = Buffer.from(got);
  const b = Buffer.from(expected);
  if (a.length !== b.length || a.length === 0) return false;
  return timingSafeEqual(a, b);
}

export async function resolveIsAdmin(email: string): Promise<{ isAdmin: boolean; check: string }> {
  if (!email) return { isAdmin: false, check: 'signed-out' };
  // Temporary access until Entra admin consent lets the Emme Admin group check succeed.
  // Remove EMME_ADMIN_EMAILS after that group lookup works.
  if (getConfig().adminEmails.indexOf(email.trim().toLowerCase()) !== -1) {
    return { isAdmin: true, check: 'allowlist' };
  }
  if (!graphConfigured()) return { isAdmin: false, check: 'not-configured' };
  const cached = adminCache.get(email);
  if (cached && cached.expires > Date.now()) {
    return { isAdmin: cached.value, check: 'cache' };
  }
  try {
    const isAdmin = await userIsAdmin(email);
    adminCache.set(email, { value: isAdmin, expires: Date.now() + 5 * 60 * 1000 });
    return { isAdmin, check: 'graph' };
  } catch (err) {
    const message = err instanceof GraphError ? err.message : 'Admin group lookup failed';
    return { isAdmin: false, check: message };
  }
}

export async function requireUser(context: FnContext, req: HttpRequest): Promise<Principal | null> {
  const principal = readPrincipal(req);
  if (!principal || !principal.email) {
    json(context, 401, { error: 'Sign in required' });
    return null;
  }
  return principal;
}

export async function requireAdmin(context: FnContext, req: HttpRequest): Promise<Principal | null> {
  if (automationAuthorized(req)) {
    return { userId: 'automation', name: 'Power Automate', email: '' };
  }
  const principal = await requireUser(context, req);
  if (!principal) return null;
  const result = await resolveIsAdmin(principal.email);
  if (!result.isAdmin) {
    json(context, 403, { error: 'Emme Admin access is required' });
    return null;
  }
  return principal;
}
