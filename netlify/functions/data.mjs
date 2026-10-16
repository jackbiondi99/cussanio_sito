// Dati pubblici per il sito: GET /api/data
import { loadDB, loadMeta, json, fail } from "../lib/common.mjs";

export default async () => {
  try {
    const db = await loadDB();
    if (!db) return json({ error: "Dati non ancora caricati." }, 404);
    const meta = await loadMeta();
    return json({ sheets: db.sheets, updatedAt: db.updatedAt, sync: meta ? { at: meta.at, ok: meta.ok } : null });
  } catch (e) { return fail(e); }
};
export const config = { path: "/api/data" };
