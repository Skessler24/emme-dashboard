export async function api<T>(path: string, options?: { method?: string; body?: unknown }): Promise<T> {
  const res = await fetch(path, {
    method: options?.method || 'GET',
    credentials: 'include',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
    body: options?.body === undefined ? undefined : JSON.stringify(options.body),
  });
  const text = await res.text();
  let data: { error?: string } = {};
  if (text) {
    try {
      data = JSON.parse(text) as { error?: string };
    } catch {
      throw new Error(res.status === 401 || res.status === 302 ? 'Sign in required' : 'The server did not return data');
    }
  }
  if (!res.ok) throw new Error(data.error || 'Request failed');
  return data as T;
}

export type Me = { isAdmin: boolean; email?: string; name?: string; error?: string };

export type Row = Record<string, unknown> & { id?: string };
