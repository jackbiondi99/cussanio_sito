// Funzioni riservate all'amministratore: /api/admin/...
//   GET    /api/admin/status                 stato dati e ultimo aggiornamento da Tuttocampo
//   POST   /api/admin/import    {sheets}     sostituisce tutti i dati con quelli del file Excel
//   POST   /api/admin/partitella {data, righe:[{giocatore, risultato}]}
//   DELETE /api/admin/partitella?giornata=N
//   POST   /api/admin/sync                   aggiorna subito da Tuttocampo
import { loadDB, saveDB, backupDB, loadMeta, normalizeDB, syncTuttocampo, requireAdmin,
         ESITI, json, fail, HttpError, toISO } from "../lib/common.mjs";

export default async req => {
  try {
    requireAdmin(req);
    const url = new URL(req.url);
    const route = url.pathname.replace(/^\/api\/admin\/?/, "").replace(/\/$/, "");
    const body = () => req.json().catch(() => { throw new HttpError(400, "Richiesta non valida."); });

    if (route === "status" && req.method === "GET") {
      const db = await loadDB();
      return json({ hasData: !!db, updatedAt: db?.updatedAt || null, sync: await loadMeta(),
        counts: db ? Object.fromEntries(Object.entries(db.sheets).map(([k, v]) => [k, v.length])) : null });
    }

    if (route === "import" && req.method === "POST") {
      const { sheets } = await body();
      const { sheets: clean, warnings } = normalizeDB(sheets);
      const old = await loadDB();
      await backupDB(old, "prima del caricamento Excel");
      const db = await saveDB({ ...(old || {}), sheets: clean });
      return json({ ok: true, warnings, updatedAt: db.updatedAt,
        counts: Object.fromEntries(Object.entries(clean).map(([k, v]) => [k, v.length])) });
    }

    if (route === "partitella" && req.method === "POST") {
      const { data, righe } = await body();
      const db = await loadDB();
      if (!db) throw new HttpError(409, "Carica prima il file Excel.");
      const iso = toISO(data);
      if (!iso) throw new HttpError(400, "Data della partitella non valida.");
      if (!Array.isArray(righe) || !righe.length) throw new HttpError(400, "Nessun giocatore indicato.");
      const known = new Set(db.sheets.Giocatori.map(p => p.Giocatore));
      for (const r of righe) {
        if (!known.has(r.giocatore)) throw new HttpError(400, `Giocatore sconosciuto: ${r.giocatore}`);
        if (!ESITI.includes(r.risultato)) throw new HttpError(400, `Risultato non valido per ${r.giocatore}.`);
      }
      if (!righe.some(r => r.risultato !== "Assente")) throw new HttpError(400, "Tutti assenti: partitella non salvata.");
      const n = Math.max(0, ...db.sheets.Partitelle.map(r => r.Giornata || 0)) + 1;
      righe.forEach(r => db.sheets.Partitelle.push({ Giornata: n, Data: iso, Giocatore: r.giocatore, Risultato: r.risultato }));
      await saveDB(db);
      return json({ ok: true, giornata: n });
    }

    if (route === "partitella" && req.method === "DELETE") {
      const g = +url.searchParams.get("giornata");
      const db = await loadDB();
      if (!db || !g) throw new HttpError(400, "Partitella non indicata.");
      const before = db.sheets.Partitelle.length;
      db.sheets.Partitelle = db.sheets.Partitelle.filter(r => r.Giornata !== g);
      if (db.sheets.Partitelle.length === before) throw new HttpError(404, "Partitella non trovata.");
      await backupDB(await loadDB(), `prima dell'eliminazione della partitella ${g}`);
      await saveDB(db);
      return json({ ok: true });
    }

    if (route === "sync" && req.method === "POST") return json(await syncTuttocampo());

    return json({ error: "Funzione non trovata." }, 404);
  } catch (e) { return fail(e); }
};
export const config = { path: "/api/admin/*" };
