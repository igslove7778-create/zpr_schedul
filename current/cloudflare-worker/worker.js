/**
 * Zephyrus Telegram immediate-response proxy
 *
 * Cloudflare Worker secrets/variables required:
 * - APPS_SCRIPT_URL  (plain text variable: existing Apps Script /exec URL)
 * - WEBHOOK_KEY      (encrypted secret: same value as Apps Script WEBHOOK_KEY)
 *
 * This Worker stores no schedules, people, or Telegram messages. It follows
 * Google Apps Script's one-time response redirect, then returns the final
 * reply directly to Telegram as a normal HTTPS 200 response.
 */
export default {
  async fetch(request, env) {
    if (request.method === 'GET') {
      return new Response('Zephyrus Telegram proxy is running.', { status: 200 });
    }
    if (request.method !== 'POST') {
      return new Response('Method not allowed.', { status: 405 });
    }
    if (!env.APPS_SCRIPT_URL || !env.WEBHOOK_KEY) {
      return new Response('Proxy is not configured.', { status: 500 });
    }

    // Telegram receives a URL containing this key. Reject casual requests
    // before they can reach the Apps Script schedule service.
    const incomingUrl = new URL(request.url);
    if (incomingUrl.searchParams.get('key') !== env.WEBHOOK_KEY) {
      return new Response('Forbidden.', { status: 403 });
    }

    const targetUrl = new URL(env.APPS_SCRIPT_URL);
    targetUrl.searchParams.set('key', env.WEBHOOK_KEY);
    const upstream = await fetch(targetUrl.toString(), {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: await request.text(),
      redirect: 'follow'
    });
    const reply = await upstream.text();

    if (!upstream.ok) {
      console.log('Apps Script returned HTTP ' + upstream.status);
      return new Response('Upstream error.', { status: 502 });
    }

    return new Response(reply, {
      status: 200,
      headers: {
        'content-type': upstream.headers.get('content-type') || 'text/plain; charset=utf-8',
        'cache-control': 'no-store'
      }
    });
  }
};
