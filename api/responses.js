// api/responses.js
// POST  → save new response
// GET   → list all responses (for admin panel)

import { sql, initSchema } from "./_db.js";

export const config = { runtime: "nodejs" };

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();

  try {
    await initSchema();

    // ── GET /api/responses ──────────────────────────────────
    if (req.method === "GET") {
      const rows = await sql(
        "SELECT id, ext_id, ts, created FROM responses ORDER BY id DESC"
      );
      return res.status(200).json({
        total: rows.length,
        items: rows.map(r => ({
          id:      r.id,
          ext_id:  r.ext_id,
          ts:      r.ts,
          created: r.created,
        }))
      });
    }

    // ── POST /api/responses ─────────────────────────────────
    if (req.method === "POST") {
      const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
      const { id: ext_id, ts, answers } = body;
      if (!answers) return res.status(400).json({ error: "answers required" });

      await sql(
        "INSERT INTO responses (ext_id, ts, answers) VALUES (?, ?, ?)",
        [String(ext_id ?? ""), ts ?? new Date().toISOString(), JSON.stringify(answers)]
      );

      return res.status(201).json({ ok: true });
    }

    return res.status(405).json({ error: "Method not allowed" });

  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message });
  }
}
