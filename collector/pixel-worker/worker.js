// Read-tracking pixel collector for the outlook-web fork.
//
// NOT DEPLOYED. Deploying it is the owner's call: see README.md.
// It answers GET /o/<32 hex>.gif with a 1x1 transparent GIF and records one row per fetch.
// It never stores a raw IP: only a truncated hash, and only time + user agent otherwise.

const GIF = Uint8Array.from(
  atob("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw=="),
  (c) => c.charCodeAt(0),
);

const TOKEN = /^\/o\/([0-9a-f]{32})\.gif$/;

async function ipHash(request) {
  const ip =
    request.headers.get("cf-connecting-ip") ||
    request.headers.get("x-forwarded-for") ||
    "";
  if (!ip) return "";
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(ip));
  return [...new Uint8Array(digest)]
    .slice(0, 4)
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === "/events") {
      const auth = request.headers.get("authorization") || "";
      if (!env.SYNC_TOKEN || auth !== `Bearer ${env.SYNC_TOKEN}`) {
        return new Response("unauthorized", { status: 401 });
      }
      const since = url.searchParams.get("since") || "1970-01-01T00:00:00Z";
      const { results } = await env.DB.prepare(
        "SELECT token, at, ua FROM events WHERE at > ? ORDER BY at LIMIT 1000",
      )
        .bind(since)
        .all();
      return Response.json(results ?? []);
    }

    const match = TOKEN.exec(url.pathname);
    if (!match) return new Response("not found", { status: 404 });

    try {
      await env.DB.prepare("INSERT INTO events (token, at, ua, ip4) VALUES (?, ?, ?, ?)")
        .bind(
          match[1],
          new Date().toISOString(),
          (request.headers.get("user-agent") || "").slice(0, 120),
          await ipHash(request),
        )
        .run();
    } catch {
      // A tracking failure must never break image delivery.
    }

    return new Response(GIF, {
      headers: {
        "content-type": "image/gif",
        // Gmail's proxy deduplicates cached images; without this a second open is invisible.
        "cache-control": "no-store, max-age=0",
        "content-length": String(GIF.byteLength),
      },
    });
  },
};
