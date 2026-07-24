/**
 * Org-wide OpenAI proxy for Emme Dashboard.
 * Key lives in Azure App Setting OPENAI_API_KEY (never in the browser).
 * Protected by Static Web Apps auth (Emme Entra sign-in).
 */
module.exports = async function (context, req) {
  const headers = {
    'Content-Type': 'application/json',
    'Cache-Control': 'no-store'
  };

  if (req.method === 'OPTIONS') {
    context.res = { status: 204, headers };
    return;
  }

  // SWA injects this when the user is signed in via Easy Auth
  const principal = req.headers['x-ms-client-principal'] || req.headers['X-MS-CLIENT-PRINCIPAL'];
  if (!principal) {
    context.res = {
      status: 401,
      headers,
      body: { error: 'Sign in to the Emme Dashboard (Emme Microsoft account) to use AI.' }
    };
    return;
  }

  const apiKey = process.env.OPENAI_API_KEY;
  const defaultModel = process.env.OPENAI_MODEL || 'gpt-4.1-mini';

  if (req.method === 'GET') {
    context.res = {
      status: 200,
      headers,
      body: {
        configured: Boolean(apiKey),
        model: defaultModel,
        org: true
      }
    };
    return;
  }

  if (!apiKey) {
    context.res = {
      status: 503,
      headers,
      body: {
        error: 'Org OpenAI is not configured yet. Add OPENAI_API_KEY in Azure Static Web App → Configuration → Application settings.'
      }
    };
    return;
  }

  let body = req.body;
  if (typeof body === 'string') {
    try { body = JSON.parse(body); } catch (_) { body = {}; }
  }
  body = body || {};

  const messages = body.messages;
  if (!Array.isArray(messages) || messages.length === 0) {
    context.res = { status: 400, headers, body: { error: 'Request must include a messages array.' } };
    return;
  }

  const model = (body.model || defaultModel).toString().trim() || defaultModel;

  try {
    const openaiRes = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: 'Bearer ' + apiKey
      },
      body: JSON.stringify({
        model,
        messages,
        temperature: 0.3,
        max_tokens: 4000
      })
    });

    const data = await openaiRes.json().catch(() => ({}));
    if (!openaiRes.ok) {
      const errMsg = (data && data.error && data.error.message) || ('OpenAI error ' + openaiRes.status);
      context.res = { status: openaiRes.status, headers, body: { error: errMsg } };
      return;
    }

    const content = data && data.choices && data.choices[0] && data.choices[0].message
      ? data.choices[0].message.content
      : null;

    if (!content) {
      context.res = { status: 502, headers, body: { error: 'OpenAI returned an empty response.' } };
      return;
    }

    context.res = { status: 200, headers, body: { content, model } };
  } catch (err) {
    context.res = {
      status: 502,
      headers,
      body: { error: (err && err.message) || 'Failed to reach OpenAI.' }
    };
  }
};
