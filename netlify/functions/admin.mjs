// Funzioni riservate all'amministratore: /api/admin/...
//   GET    /api/admin/status
//   POST   /api/admin/import     {sheets}                         sostituisce tutti i dati (file Excel)
//   POST   /api/admin/partitella {data, righe:[{giocatore, risultato}]}
//   DELETE /api/admin/partitella?giornata=N
//   POST   /api/admin/risultati  {giornata, risultati:[{casa, trasferta, gc, gt}]}
//   POST   /api/admin/tabellino  {giornata, righe:[{giocatore, stato, sostituito, gol, assist}]}
import { loadDB, saveDB, backupDB, normalizeDB, requireAdmin, ESITI, STATI, TEAM,
         json, fail, HttpError, toISO } from "../lib/common.mjs";

const intOrNull = (v, max, what) => {
  if (v == null || v === "") return null;
  const n = Number(v);
  if (!Number.isInteger(n) || n < 0 || n > max) throw new HttpError(400, `${what}: valore non valido (${v}).`);
  return n;
};

export default async req => {
  try {
    requireAdmin(req);
    const url = new URL(req.url);
    const route = url.pathname.replace(/^\/api\/admin\/?/, "").replace(/\/$/, "");
    const body = () => req.json().catch(() => { throw new HttpError(400, "Richiesta non valida."); });
    const needDB = async () => { const db = await loadDB(); if (!db) throw new HttpError(409, "Carica prima il file Excel."); return db; };

    if (route === "status" && req.method === "GET") {
      const db = await loadDB();
      return json({ hasData: !!db, updatedAt: db?.updatedAt || null,
        counts: db ? Object.fromEntries(Object.entries(db.sheets).map(([k, v]) => [k, v.length])) : null });
    }

    if (route === "import" && req.method === "POST") {
      const { sheets } = await body();
      const { sheets: clean, warnings } = normalizeDB(sheets);
      const old = await loadDB();
      await backupDB(old, "prima del caricamento Excel");
      const db = await saveDB({ ...(old || {}), sheets: clean });
      return json({ ok: true, warnings, updatedAt: db.updatedAt });
    }

    // ---------------- risultati della giornata ----------------
    if (route === "risultati" && req.method === "POST") {
      const { giornata, risultati } = await body();
      const db = await needDB();
      const g = Number(giornata);
      const rows = db.sheets.Risultati_Girone.filter(r => r.Giornata === g);
      if (!rows.length) throw new HttpError(404, `Giornata ${giornata} non trovata.`);
      if (!Array.isArray(risultati)) throw new HttpError(400, "Nessun risultato indicato.");
      for (const x of risultati) {
        const r = rows.find(r => r.Casa === x.casa && r.Trasferta === x.trasferta);
        if (!r) throw new HttpError(400, `Partita ${x.casa}-${x.trasferta} non presente nella giornata ${g}.`);
        const a = intOrNull(x.gc, 40, `${x.casa}-${x.trasferta}`), b = intOrNull(x.gt, 40, `${x.casa}-${x.trasferta}`);
        if ((a == null) !== (b == null))
          throw new HttpError(400, `${x.casa}-${x.trasferta}: inserisci i gol di entrambe le squadre, oppure lascia vuote tutte e due le caselle (partita rinviata).`);
        r.Gol_Casa = a; r.Gol_Trasferta = b;
      }
      await saveDB(db);
      return json({ ok: true, completa: rows.every(r => r.Gol_Casa != null && r.Gol_Trasferta != null) });
    }

    // ---------------- tabellino del Cussanio ----------------
    if (route === "tabellino" && req.method === "POST") {
      const { giornata, righe } = await body();
      const db = await needDB();
      const g = Number(giornata);
      const match = db.sheets.Risultati_Girone.find(r => r.Giornata === g && (r.Casa === TEAM || r.Trasferta === TEAM));
      if (!match) throw new HttpError(404, `Nella giornata ${giornata} il ${TEAM} non gioca.`);
      if (!Array.isArray(righe)) throw new HttpError(400, "Dati del tabellino mancanti.");
      const known = new Set(db.sheets.Giocatori.map(p => p.Giocatore));
      const seen = new Set();
      const nuove = righe.map(r => {
        if (!known.has(r.giocatore)) throw new HttpError(400, `Giocatore sconosciuto: ${r.giocatore}`);
        if (seen.has(r.giocatore)) throw new HttpError(400, `${r.giocatore} compare due volte.`);
        seen.add(r.giocatore);
        if (!STATI.includes(r.stato)) throw new HttpError(400, `Stato non valido per ${r.giocatore}.`);
        const giocato = r.stato !== "Non entrato";
        const gol = giocato ? intOrNull(r.gol, 20, `Gol di ${r.giocatore}`) : null;
        const ass = giocato ? intOrNull(r.assist, 20, `Assist di ${r.giocatore}`) : null;
        return { Giornata: g, Giocatore: r.giocatore, Stato: r.stato,
                 Sostituito: giocato ? (r.sostituito ? "Sì" : "No") : null,
                 Gol: gol || null, Assist: ass || null };
      });
      // ordine: titolari, subentrati, panchina
      nuove.sort((a, b) => STATI.indexOf(a.Stato) - STATI.indexOf(b.Stato));
      const altre = db.sheets.Tabellino.filter(r => r.Giornata !== g);
      const pos = altre.findIndex(r => (r.Giornata ?? 0) > g);
      db.sheets.Tabellino = pos < 0 ? [...altre, ...nuove] : [...altre.slice(0, pos), ...nuove, ...altre.slice(pos)];
      await saveDB(db);
      return json({ ok: true, righe: nuove.length });
    }

    // ---------------- partitelle ----------------
    if (route === "partitella" && req.method === "POST") {
      const { data, righe } = await body();
      const db = await needDB();
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
      const db = await needDB();
      const before = db.sheets.Partitelle.length;
      const old = JSON.parse(JSON.stringify(db));
      db.sheets.Partitelle = db.sheets.Partitelle.filter(r => r.Giornata !== g);
      if (db.sheets.Partitelle.length === before) throw new HttpError(404, "Partitella non trovata.");
      await backupDB(old, `prima dell'eliminazione della partitella ${g}`);
      await saveDB(db);
      return json({ ok: true });
    }

    return json({ error: "Funzione non trovata." }, 404);
  } catch (e) { return fail(e); }
};
export const config = { path: "/api/admin/*" };
