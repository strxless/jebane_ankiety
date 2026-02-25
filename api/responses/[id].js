// api/responses/[id].js
// GET    /api/responses/42   → single response with full answers
// DELETE /api/responses/42   → delete it

import { getClient, initSchema } from "../_db.js";

export const config = { runtime: "nodejs20.x" };

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, DELETE, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();

  const { id } = req.query;
  if (!id || isNaN(Number(id))) return res.status(400).json({ error: "Invalid id" });

  try {
    await initSchema();
    const db = getClient();

    if (req.method === "GET") {
      const result = await db.execute({
        sql: "SELECT * FROM responses WHERE id = ?",
        args: [Number(id)]
      });
      if (!result.rows.length) return res.status(404).json({ error: "Not found" });
      const r = result.rows[0];
      return res.status(200).json({
        id:      r[0],
        ext_id:  r[1],
        ts:      r[2],
        created: r[3],
        answers: JSON.parse(r[4])
      });
    }

    if (req.method === "DELETE") {
      await db.execute({ sql: "DELETE FROM responses WHERE id = ?", args: [Number(id)] });
      return res.status(204).end();
    }

    return res.status(405).json({ error: "Method not allowed" });

  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message });
  }
}
