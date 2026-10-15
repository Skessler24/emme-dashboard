export type HttpRequest = {
  method?: string;
  headers?: Record<string, string | undefined>;
  query?: Record<string, string | undefined>;
  params?: Record<string, string | undefined>;
  body?: unknown;
};

export type FnContext = {
  res?: {
    status?: number;
    headers?: Record<string, string>;
    body?: unknown;
  };
};

export function json(context: FnContext, status: number, body: unknown): void {
  context.res = {
    status,
    headers: {
      'Content-Type': 'application/json',
      'Cache-Control': 'no-store',
    },
    body,
  };
}

export function header(req: HttpRequest, name: string): string {
  const headers = (req && req.headers) || {};
  const want = name.toLowerCase();
  if (headers[want]) return String(headers[want]);
  const key = Object.keys(headers).find(function (k) { return k.toLowerCase() === want; });
  return key ? String(headers[key]) : '';
}

export function methodOf(req: HttpRequest): string {
  return String(req.method || 'GET').toUpperCase();
}

export function queryOf(req: HttpRequest, name: string): string {
  const query = (req && req.query) || {};
  const want = name.toLowerCase();
  if (query[name] != null) return String(query[name]);
  const key = Object.keys(query).find(function (k) { return k.toLowerCase() === want; });
  return key && query[key] != null ? String(query[key]) : '';
}

export function restParts(req: HttpRequest): string[] {
  const raw = String((req.params && (req.params.rest || req.params.id)) || '').replace(/^\/+|\/+$/g, '');
  if (!raw) return [];
  return raw.split('/').filter(Boolean).map(function (part) {
    try { return decodeURIComponent(part); } catch (_) { return part; }
  });
}

export function readBody(req: HttpRequest): Record<string, unknown> {
  let body = req.body;
  if (typeof body === 'string') {
    try { body = JSON.parse(body); } catch (_) { body = {}; }
  }
  if (!body || typeof body !== 'object' || Array.isArray(body)) return {};
  return body as Record<string, unknown>;
}

export function asText(value: unknown, max = 4000): string {
  return String(value == null ? '' : value).trim().slice(0, max);
}
