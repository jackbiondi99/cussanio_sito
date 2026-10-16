// Dati pubblici per il sito: GET /api/data
import { loadDB, json, fail } from "../lib/common.mjs";

export default async () => {
  try {
    const db = await loadDB();
    if (!db) return json({ error: "Dati non ancora caricati." }, 404);
    return json({ sheets: db.sheets, updatedAt: db.updatedAt });
  } catch (e) { return fail(e); }
};
export const config = { path: "/api/data" };
