import { createClient } from "@libsql/client";

let _client = null;

export function getClient() {
  if (_client) return _client;
  const url   = process.env.TURSO_DB_URL;
  const token = process.env.TURSO_AUTH_TOKEN;
  if (!url || !token) throw new Error("Missing TURSO_DB_URL or TURSO_AUTH_TOKEN env vars");
  _client = createClient({ url, authToken: token });
  return _client;
}

export async function initSchema() {
  const db = getClient();
  await db.execute(`
    CREATE TABLE IF NOT EXISTS responses (
      id      INTEGER PRIMARY KEY AUTOINCREMENT,
      ext_id  TEXT,
      ts      TEXT,
      created TEXT DEFAULT (datetime('now')),
      answers TEXT NOT NULL
    )
  `);
}
