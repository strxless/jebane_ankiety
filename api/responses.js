// api/responses.js
// POST  → save new response
// GET   → list all responses (for admin panel)

import { getClient, initSchema } from "./_db.js";

export const config = { runtime: "nodejs20.x" };

export default async function handler(req, res) {
  // CORS — allow the frontend (same domain or localhost dev)
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();

  try {
    await initSchema();
    const db = getClient();

    // ── GET /api/responses ──────────────────────────────────
    if (req.method === "GET") {
      const result = await db.execute(
        "SELECT id, ext_id, ts, created FROM responses ORDER BY id DESC"
      );
      const total = result.rows.length;
      return res.status(200).json({
        total,
        items: result.rows.map(r => ({
          id:      r[0],
          ext_id:  r[1],
          ts:      r[2],
          created: r[3],
        }))
      });
    }

    // ── POST /api/responses ─────────────────────────────────
    if (req.method === "POST") {
      const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
      const { id: ext_id, ts, answers } = body;
      if (!answers) return res.status(400).json({ error: "answers required" });

      const result = await db.execute({
        sql: "INSERT INTO responses (ext_id, ts, answers) VALUES (?, ?, ?)",
        args: [String(ext_id ?? ""), ts ?? new Date().toISOString(), JSON.stringify(answers)]
      });

      return res.status(201).json({ ok: true, db_id: Number(result.lastInsertRowid) });
    }

    return res.status(405).json({ error: "Method not allowed" });

  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message });
  }
}
