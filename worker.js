/**
 * Engineering Dashboard — Cloudflare Worker
 * Serves the static dashboard (public/) AND a tiny key/value API backed by D1.
 *
 * The embedded Downtime Log Book and Work Order form talk to this API through a
 * small `window.storage` bridge injected into each of them, so every device
 * reads and writes the SAME shared database instead of its own browser.
 *
 * API:
 *   GET  /api/kv?key=<k>   -> { "value": <stored string or null> }
 *   POST /api/kv           body { "key": <k>, "value": <string> } -> { "ok": true }
 *   GET  /api/health       -> { "ok": true }
 */

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === "/api/health") {
      return json({ ok: true });
    }

    if (url.pathname === "/api/kv") {
      try {
        if (request.method === "GET") {
          const key = url.searchParams.get("key");
          if (!key) return json({ error: "missing key" }, 400);
          const row = await env.DB.prepare("SELECT v FROM store WHERE k = ?").bind(key).first();
          return json({ value: row ? row.v : null });
        }
        if (request.method === "POST") {
          let body;
          try { body = await request.json(); } catch (e) { return json({ error: "bad json" }, 400); }
          const key = body && body.key;
          const value = body && body.value;
          if (!key || typeof value !== "string") return json({ error: "missing key/value" }, 400);
          await env.DB
            .prepare("INSERT INTO store (k, v, updated) VALUES (?, ?, ?) ON CONFLICT(k) DO UPDATE SET v = excluded.v, updated = excluded.updated")
            .bind(key, value, Date.now())
            .run();
          return json({ ok: true });
        }
        return json({ error: "method not allowed" }, 405);
      } catch (e) {
        return json({ error: String(e && e.message || e) }, 500);
      }
    }

    // Everything else -> static assets (the dashboard)
    return env.ASSETS.fetch(request);
  }
};

function json(obj, status) {
  return new Response(JSON.stringify(obj), {
    status: status || 200,
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store" }
  });
}
