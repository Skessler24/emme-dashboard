/**
 * Loads Outlook calendar and To Do for the signed-in Emme email.
 * Uses the site sign-in, so the same person is recognized in any browser.
 * Requires AZURE_CLIENT_SECRET plus application permissions granted by an admin.
 */
const DEFAULT_CLIENT_ID = 'b1e3bd77-c545-4716-afdc-8b0ec14ddcdf';
const DEFAULT_TENANT_ID = '4da01280-a61e-44cc-b537-9b21b7f17c68';

let tokenCache = { value: '', expires: 0 };

function json(context, status, body) {
  context.res = {
    status,
    headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
    body
  };
}

function header(req, name) {
  const headers = (req && req.headers) || {};
  const want = name.toLowerCase();
  if (headers[want]) return String(headers[want]);
  const key = Object.keys(headers).find(function (k) { return k.toLowerCase() === want; });
  return key ? String(headers[key]) : '';
}

function readEmail(req) {
  const directName = header(req, 'x-ms-client-principal-name');
  const encoded = header(req, 'x-ms-client-principal');
  let parsed = null;
  if (encoded) {
    try { parsed = JSON.parse(Buffer.from(encoded, 'base64').toString('utf8')); } catch (_) { parsed = null; }
  }
  const details = String(directName || (parsed && parsed.userDetails) || '').trim().toLowerCase();
  if (details.indexOf('@') !== -1) return details;
  const claims = (parsed && parsed.claims) || [];
  for (let i = 0; i < claims.length; i++) {
    const typ = String(claims[i].typ || '').toLowerCase();
    const val = String(claims[i].val || '').trim().toLowerCase();
    if (val.indexOf('@') === -1) continue;
    if (typ.indexOf('email') !== -1 || typ.indexOf('preferred_username') !== -1 || typ.endsWith('/upn')) return val;
  }
  return '';
}

function configured() {
  return !!process.env.AZURE_CLIENT_SECRET;
}

async function appToken() {
  if (!configured()) return '';
  if (tokenCache.value && tokenCache.expires > Date.now() + 60000) return tokenCache.value;
  const tenant = process.env.AZURE_TENANT_ID || DEFAULT_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID || DEFAULT_CLIENT_ID;
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: process.env.AZURE_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });
  const res = await fetch('https://login.microsoftonline.com/' + tenant + '/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString()
  });
  const data = await res.json().catch(function () { return {}; });
  if (!res.ok || !data.access_token) {
    throw new Error('Microsoft app sign-in failed. Check the client secret and admin consent.');
  }
  tokenCache = {
    value: data.access_token,
    expires: Date.now() + (Number(data.expires_in) || 3600) * 1000
  };
  return tokenCache.value;
}

async function graph(token, path, options) {
  options = options || {};
  const headers = {
    Authorization: 'Bearer ' + token,
    'Content-Type': 'application/json'
  };
  if (options.prefer) headers.Prefer = options.prefer;
  const res = await fetch('https://graph.microsoft.com/v1.0' + path, {
    method: options.method || 'GET',
    headers: headers,
    body: options.body ? JSON.stringify(options.body) : undefined
  });
  if (res.status === 204) return null;
  const data = await res.json().catch(function () { return {}; });
  if (!res.ok) {
    const message = (data.error && (data.error.message || data.error.code)) || ('Graph ' + res.status);
    throw new Error(message);
  }
  return data;
}

function userPath(email) {
  return '/users/' + encodeURIComponent(email);
}

async function calendarFor(token, email, start, end) {
  const calendars = await graph(token, userPath(email) + '/calendars?$select=id,owner,isDefaultCalendar');
  const owned = (calendars.value || []).filter(function (cal) {
    const addr = String((cal.owner && cal.owner.address) || '').toLowerCase();
    if (!addr) return !!cal.isDefaultCalendar;
    return addr === email;
  });
  const sources = owned.length ? owned : [{ id: '' }];
  const events = [];
  for (let i = 0; i < sources.length; i++) {
    const base = sources[i].id
      ? userPath(email) + '/calendars/' + encodeURIComponent(sources[i].id) + '/calendarView'
      : userPath(email) + '/calendarView';
    const data = await graph(
      token,
      base + '?startDateTime=' + encodeURIComponent(start) +
        '&endDateTime=' + encodeURIComponent(end) +
        '&$select=subject,start,end,isAllDay,webLink&$orderby=start/dateTime&$top=100',
      { prefer: 'outlook.timezone="America/Denver"' }
    );
    (data.value || []).forEach(function (ev) { events.push(ev); });
  }
  return events;
}

async function tasksFor(token, email) {
  const lists = await graph(token, userPath(email) + '/todo/lists');
  const listsArr = lists.value || [];
  const items = [];
  for (let i = 0; i < listsArr.length; i++) {
    const data = await graph(
      token,
      userPath(email) + '/todo/lists/' + encodeURIComponent(listsArr[i].id) +
        '/tasks?$top=50&$orderby=createdDateTime desc'
    );
    (data.value || []).forEach(function (task) {
      items.push({
        id: task.id,
        listId: listsArr[i].id,
        text: task.title || '(Untitled)',
        completed: task.status === 'completed'
      });
    });
  }
  return { lists: listsArr, items: items };
}

function defaultListId(lists) {
  const preferred = (lists || []).find(function (list) { return list.wellknownListName === 'defaultList'; })
    || (lists || []).find(function (list) { return /tasks/i.test(list.displayName || ''); })
    || (lists || [])[0];
  return preferred && preferred.id;
}

module.exports = async function (context, req) {
  if (req.method === 'OPTIONS') {
    context.res = { status: 204, headers: { 'Content-Type': 'application/json' } };
    return;
  }

  const email = readEmail(req);
  if (!email) {
    json(context, 401, { configured: configured(), error: 'Sign in required' });
    return;
  }

  const rest = String((req.params && (req.params.rest || req.params.id)) || '').replace(/^\/+|\/+$/g, '');
  const method = (req.method || 'GET').toUpperCase();

  if (rest === 'status' && method === 'GET') {
    json(context, 200, { configured: configured(), email: email });
    return;
  }

  if (!configured()) {
    json(context, 200, {
      configured: false,
      email: email,
      error: 'Add AZURE_CLIENT_SECRET on the Static Web App so calendar and to-do follow the signed-in email.'
    });
    return;
  }

  try {
    const token = await appToken();
    if (rest === 'calendar' && method === 'GET') {
      const start = String((req.query && req.query.start) || '');
      const end = String((req.query && req.query.end) || '');
      if (!start || !end) {
        json(context, 400, { error: 'Start and end are required' });
        return;
      }
      const events = await calendarFor(token, email, start, end);
      json(context, 200, { configured: true, email: email, events: events });
      return;
    }

    if (rest === 'tasks' && method === 'GET') {
      const result = await tasksFor(token, email);
      json(context, 200, { configured: true, email: email, items: result.items });
      return;
    }

    if (rest === 'tasks' && method === 'POST') {
      let body = req.body;
      if (typeof body === 'string') {
        try { body = JSON.parse(body); } catch (_) { body = {}; }
      }
      const title = String((body && body.title) || '').trim();
      if (!title) {
        json(context, 400, { error: 'A task title is required' });
        return;
      }
      const lists = await graph(token, userPath(email) + '/todo/lists');
      const listId = defaultListId(lists.value || []);
      if (!listId) throw new Error('No Microsoft To Do list found');
      const created = await graph(
        token,
        userPath(email) + '/todo/lists/' + encodeURIComponent(listId) + '/tasks',
        { method: 'POST', body: { title: title } }
      );
      json(context, 201, {
        configured: true,
        item: { id: created.id, listId: listId, text: created.title || title, completed: false }
      });
      return;
    }

    if (rest === 'tasks' && method === 'PATCH') {
      let body = req.body;
      if (typeof body === 'string') {
        try { body = JSON.parse(body); } catch (_) { body = {}; }
      }
      const id = String((body && body.id) || '');
      const listId = String((body && body.listId) || '');
      if (!id || !listId) {
        json(context, 400, { error: 'Task id is required' });
        return;
      }
      await graph(
        token,
        userPath(email) + '/todo/lists/' + encodeURIComponent(listId) + '/tasks/' + encodeURIComponent(id),
        { method: 'PATCH', body: { status: body.completed ? 'completed' : 'notStarted' } }
      );
      json(context, 200, { configured: true });
      return;
    }

    if (rest === 'tasks' && method === 'DELETE') {
      const id = String((req.query && req.query.id) || '');
      const listId = String((req.query && req.query.listId) || '');
      if (!id || !listId) {
        json(context, 400, { error: 'Task id is required' });
        return;
      }
      await graph(
        token,
        userPath(email) + '/todo/lists/' + encodeURIComponent(listId) + '/tasks/' + encodeURIComponent(id),
        { method: 'DELETE' }
      );
      json(context, 200, { configured: true });
      return;
    }

    json(context, 404, { error: 'Not found' });
  } catch (err) {
    json(context, 502, { configured: true, error: (err && err.message) || 'Could not reach Microsoft' });
  }
};
