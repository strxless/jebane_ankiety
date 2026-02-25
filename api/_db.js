// Direct Turso HTTP API — no @libsql/client, no migrations nonsense

function getConfig() {
  const url   = process.env.TURSO_DB_URL;
  const token = process.env.TURSO_AUTH_TOKEN;
  if (!url || !token) throw new Error("Missing TURSO_DB_URL or TURSO_AUTH_TOKEN env vars");
  // Convert libsql:// to https://
  const httpUrl = url.replace(/^libsql:\/\//, 'https://');
  return { httpUrl, token };
}

export async function sql(query, args = []) {
  const { httpUrl, token } = getConfig();
  const res = await fetch(`${httpUrl}/v2/pipeline`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      requests: [
        { type: 'execute', stmt: { sql: query, args: args.map(toArg) } },
        { type: 'close' },
      ]
    }),
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Turso error ${res.status}: ${text}`);
  }
  const data = await res.json();
  const result = data.results?.[0];
  if (result?.type === 'error') throw new Error(result.error?.message || 'SQL error');
  const { cols = [], rows = [] } = result?.response?.result ?? {};
  return rows.map(row =>
    Object.fromEntries(cols.map((c, i) => [c.name, row[i]?.value ?? null]))
  );
}

function toArg(v) {
  if (v === null || v === undefined) return { type: 'null' };
  if (typeof v === 'number' && Number.isInteger(v)) return { type: 'integer', value: String(v) };
  if (typeof v === 'number') return { type: 'float', value: v };
  return { type: 'text', value: String(v) };
}

export async function initSchema() {
  await sql(`
    CREATE TABLE IF NOT EXISTS responses (
      id      INTEGER PRIMARY KEY AUTOINCREMENT,
      ext_id  TEXT,
      ts      TEXT,
      created TEXT DEFAULT (datetime('now')),
      answers TEXT NOT NULL
    )
  `);
}
