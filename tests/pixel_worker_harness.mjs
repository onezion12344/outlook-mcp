// Runs the real collector/pixel-worker/worker.js over node:sqlite.
//
// The only thing stubbed is Cloudflare's D1 API (prepare/bind/all/run), backed by a
// real SQLite file. That keeps the worker's own SQL, routes, headers and GIF bytes
// under test without decking out a wrangler install.
//
// usage: node harness.mjs <sqlite-path>
// prints one JSON line: {"port": <n>} then serves until stdin closes.

import { createServer } from "node:http";
import { DatabaseSync } from "node:sqlite";

const dbPath = process.argv[2];
if (!dbPath) {
  console.error("usage: node pixel_worker_harness.mjs <sqlite-path>");
  process.exit(2);
}

const db = new DatabaseSync(dbPath);
// Exactly the statements from README.md's deploy section.
db.exec(
  "CREATE TABLE IF NOT EXISTS events (" +
    "id INTEGER PRIMARY KEY AUTOINCREMENT, token TEXT NOT NULL, at TEXT NOT NULL, " +
    "ua TEXT NOT NULL DEFAULT '', ip4 TEXT NOT NULL DEFAULT '')",
);
db.exec("CREATE INDEX IF NOT EXISTS events_token_at ON events(token, at)");

const D1 = {
  prepare(sql) {
    return {
      bind(...values) {
        return {
          async all() {
            return { results: db.prepare(sql).all(...values) };
          },
          async run() {
            db.prepare(sql).run(...values);
            return { meta: {} };
          },
        };
      },
    };
  },
};

const env = { DB: D1, SYNC_TOKEN: process.env.SYNC_TOKEN || "" };
const worker = (await import("../collector/pixel-worker/worker.js")).default;

const server = createServer(async (req, res) => {
  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  const request = new Request(`http://127.0.0.1${req.url}`, {
    method: req.method,
    headers: Object.entries(req.headers).map(([k, v]) => [k, Array.isArray(v) ? v.join(",") : v]),
    body: chunks.length ? Buffer.concat(chunks) : undefined,
  });
  const response = await worker.fetch(request, env);
  res.writeHead(response.status, Object.fromEntries(response.headers));
  res.end(Buffer.from(await response.arrayBuffer()));
});

server.listen(0, "127.0.0.1", () => {
  process.stdout.write(JSON.stringify({ port: server.address().port }) + "\n");
});

const stop = () => server.close(() => process.exit(0));
process.on("SIGTERM", stop);
process.stdin.on("end", stop);
