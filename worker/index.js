/**
 * Excel Builder AI Proxy — Cloudflare Worker
 *
 * Uses Cloudflare Workers AI (built-in) — NO external API key needed.
 * Free forever: 10,000 AI requests/day on the free plan.
 *
 * Setup:
 *   1. npx wrangler login     (free Cloudflare account)
 *   2. npx wrangler deploy    (that's it — done)
 */

// ── Config ──
const ALLOWED_ORIGINS = [
  'https://shinrising3-debug.github.io',
  'http://localhost:3000',
  'http://127.0.0.1:5500',
];

const RATE_LIMIT_MAX = 10;        // requests per window per IP
const RATE_LIMIT_WINDOW = 3600;   // 1 hour in seconds

const SYSTEM_PROMPT = `You convert spreadsheet descriptions into a JSON specification.

RESPOND WITH ONLY A JSON OBJECT. No markdown, no backticks, no explanation.

JSON format:
{
  "title": "Sheet Title Here",
  "columns": "Name, Date (date), Amount (currency), Status (dropdown: Option1|Option2|Option3), Notes (long)",
  "sample_data": "John Smith|2024-01-15|1500.00|Option1|First entry\\nJane Doe|2024-01-16|2300.00|Option2|Second entry"
}

Column type syntax in the "columns" string:
- Plain text (default): just the name, e.g. "Employee Name"
- Date formatted: add (date), e.g. "Hire Date (date)"
- Number: add (number), e.g. "Quantity (number)"
- Currency: add (currency), e.g. "Total Cost (currency)"
- Dropdown: add (dropdown: Opt1|Opt2|Opt3), e.g. "Status (dropdown: Pending|Fixed|Escalated)"
- Wide text column: add (long), e.g. "Notes (long)"

Sample data rules:
- Columns separated by |
- Rows separated by \\n (newline)
- Match the column order (excluding the # column which is auto-added)
- If the user didn't mention sample data, include 3-5 realistic example rows anyway

RESPOND WITH ONLY THE JSON OBJECT. Nothing else.`;

// ── CORS headers ──
function corsHeaders(origin) {
  return {
    'Access-Control-Allow-Origin': origin || ALLOWED_ORIGINS[0],
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Max-Age': '86400',
  };
}

function jsonResponse(body, status, origin) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json', ...corsHeaders(origin) },
  });
}

// ── Rate limiting via KV ──
async function checkRateLimit(ip, env) {
  if (!env.RATE_LIMITER) return { allowed: true, remaining: RATE_LIMIT_MAX };

  const key = `rl:${ip}`;
  const record = await env.RATE_LIMITER.get(key, { type: 'json' });
  const now = Math.floor(Date.now() / 1000);

  if (!record || now - record.start > RATE_LIMIT_WINDOW) {
    await env.RATE_LIMITER.put(key, JSON.stringify({ count: 1, start: now }), {
      expirationTtl: RATE_LIMIT_WINDOW,
    });
    return { allowed: true, remaining: RATE_LIMIT_MAX - 1 };
  }

  if (record.count >= RATE_LIMIT_MAX) {
    const retryAfter = RATE_LIMIT_WINDOW - (now - record.start);
    return { allowed: false, remaining: 0, retryAfter };
  }

  record.count++;
  await env.RATE_LIMITER.put(key, JSON.stringify(record), {
    expirationTtl: RATE_LIMIT_WINDOW - (now - record.start),
  });
  return { allowed: true, remaining: RATE_LIMIT_MAX - record.count };
}

// ── Main handler ──
export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const origin = request.headers.get('Origin') || '';

    // CORS preflight
    if (request.method === 'OPTIONS') {
      const allowedOrigin = ALLOWED_ORIGINS.find(o => origin.startsWith(o)) || '';
      return new Response(null, { status: 204, headers: corsHeaders(allowedOrigin) });
    }

    // Only allow POST to /generate
    if (url.pathname !== '/generate' || request.method !== 'POST') {
      return jsonResponse({ error: 'Not found' }, 404, origin);
    }

    // Origin check
    const allowedOrigin = ALLOWED_ORIGINS.find(o => origin.startsWith(o));
    if (!allowedOrigin && origin) {
      return jsonResponse({ error: 'Origin not allowed' }, 403, origin);
    }
    const respOrigin = allowedOrigin || ALLOWED_ORIGINS[0];

    // Rate limit
    const ip = request.headers.get('CF-Connecting-IP') || 'unknown';
    const rl = await checkRateLimit(ip, env);
    if (!rl.allowed) {
      return jsonResponse(
        { error: `Rate limit exceeded. Try again in ${Math.ceil(rl.retryAfter / 60)} minutes.` },
        429,
        respOrigin,
      );
    }

    // Parse request
    let body;
    try {
      body = await request.json();
    } catch {
      return jsonResponse({ error: 'Invalid JSON body' }, 400, respOrigin);
    }

    const description = (body.description || '').trim();
    if (!description) {
      return jsonResponse({ error: 'Missing "description" field' }, 400, respOrigin);
    }
    if (description.length > 2000) {
      return jsonResponse({ error: 'Description too long (max 2000 chars)' }, 400, respOrigin);
    }

    // Call Cloudflare Workers AI — built-in, no API key needed
    try {
      const aiResp = await env.AI.run('@cf/meta/llama-3.1-70b-instruct', {
        messages: [
          { role: 'system', content: SYSTEM_PROMPT },
          { role: 'user', content: description },
        ],
        temperature: 0.1,
        max_tokens: 4000,
      });

      let raw = (aiResp.response || '').trim();

      // Strip markdown fences
      raw = raw.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();

      // Validate it's JSON
      const spec = JSON.parse(raw);
      if (!spec.title || !spec.columns) {
        throw new Error('Incomplete spec');
      }

      return jsonResponse(
        {
          spec,
          remaining: rl.remaining,
        },
        200,
        respOrigin,
      );
    } catch (err) {
      console.error('Generate error:', err);
      return jsonResponse(
        { error: 'AI returned an invalid response — try rephrasing your description' },
        422,
        respOrigin,
      );
    }
  },
};
