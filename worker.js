/**
 * Engineering Dashboard — Cloudflare Worker
 * Serves the static dashboard (public/) AND a tiny key/value API backed by D1.
 *
 * The embedded Downtime Log Book and Work Order form talk to this API through a
 * small `window.storage` bridge injected into each of them, so every device
 * reads and writes the SAME shared database instead of its own browser.
 *
 * API:
 *   GET  /api/kv?key=<k>        -> { "value": <stored string or null> }
 *   POST /api/kv               body { "key": <k>, "value": <string> } -> { "ok": true }
 *   GET  /api/health           -> { "ok": true }
 *
 *   --- OEE auto-update (fed by a Power Automate flow on SharePoint) ---
 *   POST /api/oee-ingest        body = raw .xlsx bytes (octet-stream)
 *                               header  X-Ingest-Key: <INGEST_KEY>
 *                               -> stores the file in KV, returns { ok, ts, size }
 *   GET  /api/oee-ingest?meta=1 -> { ts, size, name } | null   (cheap poll)
 *   GET  /api/oee-ingest        -> the raw .xlsx bytes (dashboard downloads & parses)
 *
 * The raw workbook is kept in a Workers KV namespace (bound as OEE_FILES) because
 * it is a few MB — too large for a single D1 row. The dashboard polls ?meta=1 and,
 * when the timestamp is newer than what it has, downloads and parses the file in the
 * browser (same parser as the manual "Upload Excel" button), then renders automatically.
 */

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    // Shared secret the Power Automate flow must send. Override it by setting an
    // INGEST_KEY environment variable/secret in the Cloudflare dashboard.
    const INGEST_KEY = (env && env.INGEST_KEY) || "ALCO-oee-7Qx2026";

    if (url.pathname === "/api/health") {
      return json({ ok: true, db: !!env.DB, kv: !!env.OEE_FILES });
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

    // ---- OEE workbook ingest / download (Power Automate -> KV -> dashboard) ----
    if (url.pathname === "/api/oee-ingest") {
      try {
        const KV = env.OEE_FILES;
        if (!KV) return json({ error: "OEE_FILES KV namespace not bound on this Worker" }, 500);

        if (request.method === "GET") {
          if (url.searchParams.get("meta") != null) {
            const meta = await KV.get("meta");
            return new Response(meta || "null", {
              headers: { "Content-Type": "application/json", "Cache-Control": "no-store" }
            });
          }
          const buf = await KV.get("file", "arrayBuffer");
          if (!buf) return json({ error: "no file yet" }, 404);
          return new Response(buf, {
            headers: {
              "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              "Cache-Control": "no-store"
            }
          });
        }

        if (request.method === "POST") {
          const given = request.headers.get("X-Ingest-Key") || url.searchParams.get("key") || "";
          if (given !== INGEST_KEY) return json({ error: "unauthorized" }, 401);
          const ab = await request.arrayBuffer();
          if (!ab || ab.byteLength < 100) return json({ error: "empty body" }, 400);
          if (ab.byteLength > 24 * 1024 * 1024) return json({ error: "file too large (>24MB)" }, 413);
          const ts = Date.now();
          const name = url.searchParams.get("name") || "OEE Machine SMD 2026.xlsx";
          await KV.put("file", ab);
          await KV.put("meta", JSON.stringify({ ts: ts, size: ab.byteLength, name: name }));
          return json({ ok: true, ts: ts, size: ab.byteLength });
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
