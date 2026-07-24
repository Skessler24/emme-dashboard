/**
 * Org-wide AI proxy for Emme Dashboard.
 * Supports:
 *   - Anthropic (PRISM) via ANTHROPIC_API_KEY or OPENAI_API_KEY starting with sk-ant-
 *   - Azure OpenAI via AZURE_OPENAI_*
 *   - OpenAI.com via OPENAI_API_KEY (sk- / sk-proj-)
 * Secrets live in Static Web App Environment variables — never in the browser.
 * Route protection: staticwebapp.config.json requires authenticated users for /api/*.
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

  const azureEndpoint = (process.env.AZURE_OPENAI_ENDPOINT || '').replace(/\/+$/, '');
  const azureKey = process.env.AZURE_OPENAI_API_KEY || '';
  const azureDeployment = process.env.AZURE_OPENAI_DEPLOYMENT || '';
  const azureApiVersion = process.env.AZURE_OPENAI_API_VERSION || '2024-10-21';

  const openAiKey = process.env.OPENAI_API_KEY || '';
  const anthropicKey = process.env.ANTHROPIC_API_KEY ||
    (openAiKey.indexOf('sk-ant-') === 0 ? openAiKey : '');
  const openAiOnlyKey = (openAiKey && openAiKey.indexOf('sk-ant-') !== 0) ? openAiKey : '';

  const defaultAnthropicModel = process.env.ANTHROPIC_MODEL || process.env.OPENAI_MODEL || 'claude-sonnet-4-20250514';
  const defaultOpenAiModel = process.env.OPENAI_MODEL || azureDeployment || 'gpt-4.1-mini';

  const useAzure = Boolean(azureEndpoint && azureKey && azureDeployment);
  const useAnthropic = Boolean(anthropicKey);
  const useOpenAI = Boolean(openAiOnlyKey);
  const configured = useAzure || useAnthropic || useOpenAI;
  const provider = useAzure ? 'azure' : (useAnthropic ? 'anthropic' : (useOpenAI ? 'openai' : 'none'));

  if (req.method === 'GET') {
    context.res = {
      status: 200,
      headers,
      body: {
        configured,
        provider,
        model: useAzure ? azureDeployment : (useAnthropic ? defaultAnthropicModel : defaultOpenAiModel),
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
        error: 'Org AI is not configured. Add ANTHROPIC_API_KEY (or OPENAI_API_KEY with a sk-ant- / sk- key) in Static Web App → Environment variables.'
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
    let content;
    let modelLabel;

    if (useAzure) {
      modelLabel = azureDeployment;
      const url = azureEndpoint +
        '/openai/deployments/' + encodeURIComponent(azureDeployment) +
        '/chat/completions?api-version=' + encodeURIComponent(azureApiVersion);
      const openaiRes = await fetch(url, {
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
      const data = await openaiRes.json().catch(() => ({}));
      if (!openaiRes.ok) {
        const errMsg = (data && data.error && data.error.message) || ('Azure OpenAI error ' + openaiRes.status);
        context.res = { status: openaiRes.status, headers, body: { error: errMsg } };
        return;
      }
      content = data && data.choices && data.choices[0] && data.choices[0].message
        ? data.choices[0].message.content
        : null;
    } else if (useAnthropic) {
      modelLabel = (body.model || defaultAnthropicModel).toString().trim() || defaultAnthropicModel;
      // Avoid sending OpenAI model names to Anthropic
      if (/^gpt-|^o[0-9]/i.test(modelLabel)) {
        modelLabel = defaultAnthropicModel;
      }

      const systemParts = [];
      const anthMessages = [];
      messages.forEach(function (m) {
        if (!m || !m.role) return;
        if (m.role === 'system') {
          systemParts.push(String(m.content || ''));
          return;
        }
        if (m.role === 'user' || m.role === 'assistant') {
          anthMessages.push({
            role: m.role,
            content: String(m.content || '')
          });
        }
      });

      // Anthropic requires alternating roles starting with user
      if (!anthMessages.length || anthMessages[0].role !== 'user') {
        anthMessages.unshift({ role: 'user', content: 'Please continue.' });
      }

      const anthRes = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': anthropicKey,
          'anthropic-version': '2023-06-01'
        },
        body: JSON.stringify({
          model: modelLabel,
          max_tokens: 4000,
          temperature: 0.3,
          system: systemParts.filter(Boolean).join('\n\n') || undefined,
          messages: anthMessages
        })
      });
      const data = await anthRes.json().catch(() => ({}));
      if (!anthRes.ok) {
        const errMsg = (data && data.error && data.error.message) || ('Anthropic error ' + anthRes.status);
        context.res = { status: anthRes.status, headers, body: { error: errMsg } };
        return;
      }
      if (data && Array.isArray(data.content)) {
        content = data.content
          .filter(function (part) { return part && part.type === 'text'; })
          .map(function (part) { return part.text; })
          .join('\n');
      } else {
        content = null;
      }
    } else {
      modelLabel = (body.model || defaultOpenAiModel).toString().trim() || defaultOpenAiModel;
      const openaiRes = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: 'Bearer ' + openAiOnlyKey
        },
        body: JSON.stringify({
          model: modelLabel,
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
      content = data && data.choices && data.choices[0] && data.choices[0].message
        ? data.choices[0].message.content
        : null;
    }

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
