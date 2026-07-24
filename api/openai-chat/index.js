/**
 * Org-wide AI proxy for Emme Dashboard.
 * Prefers Azure OpenAI resource (emme-openai), falls back to OpenAI.com.
 * Secrets live in Static Web App Application settings — never in the browser.
 * Protected by Static Web Apps Easy Auth (Emme Entra sign-in).
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

  const principal = req.headers['x-ms-client-principal'] || req.headers['X-MS-CLIENT-PRINCIPAL'];
  if (!principal) {
    context.res = {
      status: 401,
      headers,
      body: { error: 'Sign in to the Emme Dashboard (Emme Microsoft account) to use AI.' }
    };
    return;
  }

  const azureEndpoint = (process.env.AZURE_OPENAI_ENDPOINT || '').replace(/\/+$/, '');
  const azureKey = process.env.AZURE_OPENAI_API_KEY || '';
  const azureDeployment = process.env.AZURE_OPENAI_DEPLOYMENT || '';
  const azureApiVersion = process.env.AZURE_OPENAI_API_VERSION || '2024-10-21';

  const openAiKey = process.env.OPENAI_API_KEY || '';
  const defaultModel = process.env.OPENAI_MODEL || azureDeployment || 'gpt-4.1-mini';

  const useAzure = Boolean(azureEndpoint && azureKey && azureDeployment);
  const useOpenAI = Boolean(openAiKey);
  const configured = useAzure || useOpenAI;
  const provider = useAzure ? 'azure' : (useOpenAI ? 'openai' : 'none');

  if (req.method === 'GET') {
    context.res = {
      status: 200,
      headers,
      body: {
        configured,
        provider,
        model: useAzure ? azureDeployment : defaultModel,
        org: true
      }
    };
    return;
  }

  if (!configured) {
    context.res = {
      status: 503,
      headers,
      body: {
        error: 'Org AI is not configured. On Static Web App → Configuration, add AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_API_KEY, and AZURE_OPENAI_DEPLOYMENT from your emme-openai resource (or OPENAI_API_KEY as a fallback).'
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

  try {
    let openaiRes;
    let modelLabel;

    if (useAzure) {
      modelLabel = azureDeployment;
      const url = azureEndpoint +
        '/openai/deployments/' + encodeURIComponent(azureDeployment) +
        '/chat/completions?api-version=' + encodeURIComponent(azureApiVersion);
      openaiRes = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'api-key': azureKey
        },
        body: JSON.stringify({
          messages,
          temperature: 0.3,
          max_tokens: 4000
        })
      });
    } else {
      modelLabel = (body.model || defaultModel).toString().trim() || defaultModel;
      openaiRes = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: 'Bearer ' + openAiKey
        },
        body: JSON.stringify({
          model: modelLabel,
          messages,
          temperature: 0.3,
          max_tokens: 4000
        })
      });
    }

    const data = await openaiRes.json().catch(() => ({}));
    if (!openaiRes.ok) {
      const errMsg = (data && data.error && data.error.message) || ('AI error ' + openaiRes.status);
      context.res = { status: openaiRes.status, headers, body: { error: errMsg } };
      return;
    }

    const content = data && data.choices && data.choices[0] && data.choices[0].message
      ? data.choices[0].message.content
      : null;

    if (!content) {
      context.res = { status: 502, headers, body: { error: 'AI returned an empty response.' } };
      return;
    }

    context.res = { status: 200, headers, body: { content, model: modelLabel, provider } };
  } catch (err) {
    context.res = {
      status: 502,
      headers,
      body: { error: (err && err.message) || 'Failed to reach AI service.' }
    };
  }
};
