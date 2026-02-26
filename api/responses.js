// api/responses.js
// GET /api/responses          → { total, items: [{id, ts, created}] }  (no answers)
// GET /api/responses?full=1   → { total, items: [{id, ts, created, answers: {...}}] }
// POST /api/responses         → saves new response, returns { ok, db_id }

import { sql, initSchema } from "./_db.js";

export const config = { runtime: "nodejs", maxDuration: 15 };

function safeParseAnswers(raw) {
  if (!raw) return {};
  if (typeof raw === "object") return raw;
  try { return JSON.parse(raw); } catch { return {}; }
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();

  try {
    await initSchema();

    // ── GET ─────────────────────────────────────────────
    if (req.method === "GET") {
      const full = req.query.full === "1" || req.query.full === "true";
      const rows = await sql("SELECT id, ext_id, ts, created, answers FROM responses ORDER BY id");

      const items = rows.map(r => {
        const base = {
          id: r.id,
          ext_id: r.ext_id,
          ts: r.ts || r.created || null,
          created: r.created || null,
        };
        if (full) {
          // Always return answers as a parsed object so the frontend
          // doesn't need to double-parse or guess the type
          base.answers = safeParseAnswers(r.answers);
        }
        return base;
      });

      return res.status(200).json({ total: items.length, items });
    }

    // ── POST ────────────────────────────────────────────
    if (req.method === "POST") {
      let body = req.body;
      // In some Vercel configs body may arrive as a string
      if (typeof body === "string") {
        try { body = JSON.parse(body); } catch { return res.status(400).json({ error: "Invalid JSON" }); }
      }

      const answers = body.answers || {};
      const ts      = body.ts || new Date().toISOString();
      const ext_id  = body.id ? String(body.id) : null;

      // Always store answers as a JSON string
      const answersStr = typeof answers === "string" ? answers : JSON.stringify(answers);

      await sql(
        "INSERT INTO responses (ext_id, ts, answers) VALUES (?, ?, ?)",
        [ext_id, ts, answersStr]
      );

      // Get the inserted id
      const lastRow = await sql("SELECT last_insert_rowid() as id");
      const db_id = lastRow[0]?.id ?? null;

      return res.status(200).json({ ok: true, db_id });
    }

    return res.status(405).json({ error: "Method not allowed" });

  } catch (err) {
    console.error("[responses] error:", err);
    return res.status(500).json({ error: err.message });
  }
}
