// =====================================================================
// Codice condiviso dalle funzioni Netlify del sito Cussanio
// =====================================================================
import { getStore } from "@netlify/blobs";
import crypto from "node:crypto";

export const TEAM = process.env.TEAM_NAME || "Cussanio";

// ---------------------------------------------------------------------
// Struttura dei dati (identica ai fogli del file Excel)
// ---------------------------------------------------------------------
export const SHEETS = {
  Giocatori:        { cols: ["Giocatore", "Nome", "Cognome", "Ruolo", "Foto_URL", "Attivo", "ID_Tuttocampo"], num: ["ID_Tuttocampo"] },
  Squadre:          { cols: ["Squadra", "Penalità", "Logo_URL", "ID_Tuttocampo"], num: ["Penalità", "ID_Tuttocampo"] },
  Risultati_Girone: { cols: ["Giornata", "Data", "Casa", "Trasferta", "Gol_Casa", "Gol_Trasferta"], num: ["Giornata", "Gol_Casa", "Gol_Trasferta"], date: ["Data"] },
  Tabellino:        { cols: ["Giornata", "Giocatore", "Stato", "Sostituito", "Gol", "Assist"], num: ["Giornata", "Gol", "Assist"] },
  Partitelle:       { cols: ["Giornata", "Data", "Giocatore", "Risultato"], num: ["Giornata"], date: ["Data"] },
};
export const ESITI = ["Vittoria", "Pareggio", "Sconfitta", "Assente"];
export const STATI = ["Titolare", "Subentrato", "Non entrato"];

// ---------------------------------------------------------------------
// Archivio (Netlify Blobs)
// ---------------------------------------------------------------------
const store = () => getStore({ name: "cussanio", consistency: "strong" });
export const loadDB   = async () => (await store().get("db", { type: "json" })) || null;
export const saveDB   = async db => { db.updatedAt = new Date().toISOString(); await store().setJSON("db", db); return db; };
export const loadMeta = async () => null;   // (aggiornamento da Tuttocampo disattivato)
export async function backupDB(db, reason){
  if (!db) return;
  const key = "backup/" + new Date().toISOString().replace(/[:.]/g, "-");
  await store().setJSON(key, { reason, db });
}

// ---------------------------------------------------------------------
// Risposte HTTP
// ---------------------------------------------------------------------
export const json = (obj, status = 200) => new Response(JSON.stringify(obj), {
  status, headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" },
});
export class HttpError extends Error { constructor(status, msg){ super(msg); this.status = status; } }
export const fail = e => json({ error: e.message || String(e) }, e.status || 500);

// ---------------------------------------------------------------------
// Accesso amministratore (password in variabile d'ambiente ADMIN_PASSWORD)
// ---------------------------------------------------------------------
function secret(){
  const p = process.env.ADMIN_PASSWORD;
  if (!p) throw new HttpError(500, "Variabile ADMIN_PASSWORD non configurata su Netlify.");
  return crypto.createHash("sha256").update("cussanio-token:" + p).digest();
}
export function checkPassword(pw){
  const p = process.env.ADMIN_PASSWORD;
  if (!p) throw new HttpError(500, "Variabile ADMIN_PASSWORD non configurata su Netlify.");
  const a = crypto.createHash("sha256").update(String(pw ?? "")).digest();
  const b = crypto.createHash("sha256").update(p).digest();
  return crypto.timingSafeEqual(a, b);
}
export function makeToken(hours = 12){
  const exp = String(Date.now() + hours * 3600e3);
  return exp + "." + crypto.createHmac("sha256", secret()).update(exp).digest("base64url");
}
export function requireAdmin(req){
  const t = (req.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  const [exp, sig] = t.split(".");
  if (!exp || !sig || +exp < Date.now()) throw new HttpError(401, "Sessione scaduta: effettua di nuovo l'accesso.");
  const good = crypto.createHmac("sha256", secret()).update(exp).digest("base64url");
  if (sig.length !== good.length || !crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(good)))
    throw new HttpError(401, "Accesso non valido.");
}

// ---------------------------------------------------------------------
// Pulizia e controllo dei dati caricati da Excel
// ---------------------------------------------------------------------
const pad = n => String(n).padStart(2, "0");
export function toISO(v){
  if (v == null || v === "") return null;
  if (typeof v === "number" && isFinite(v)) return new Date(Math.round((v - 25569) * 864e5)).toISOString().slice(0, 10);
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
  m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})/);
  if (m) return `${m[3].length === 2 ? "20" + m[3] : m[3]}-${pad(m[2])}-${pad(m[1])}`;
  return undefined;
}
const clean = v => {
  if (v == null) return null;
  if (typeof v === "string") { const t = v.trim(); return t === "" ? null : t; }
  return v;
};

export function normalizeDB(input){
  const warnings = [];
  if (!input || typeof input !== "object") throw new HttpError(400, "Dati mancanti.");
  const sheets = {};
  for (const [name, def] of Object.entries(SHEETS)) {
    const rows = input[name];
    if (!Array.isArray(rows)) throw new HttpError(400, `Manca il foglio "${name}" nel file.`);
    const out = [];
    rows.forEach((r, i) => {
      const o = {};
      for (const c of def.cols) {
        let v = clean(r[c]);
        if (v != null && def.num?.includes(c)) {
          const n = typeof v === "number" ? v : parseFloat(String(v).replace(",", "."));
          if (isNaN(n)) { warnings.push(`${name}, riga ${i + 2}: "${c}" non è un numero (${v}), lasciato vuoto.`); v = null; }
          else v = n;
        }
        if (v != null && def.date?.includes(c)) {
          const iso = toISO(v);
          if (iso === undefined) { warnings.push(`${name}, riga ${i + 2}: data non riconosciuta (${v}).`); v = null; }
          else v = iso;
        }
        o[c] = v ?? null;
      }
      if (Object.values(o).some(v => v != null)) out.push(o);
    });
    sheets[name] = out;
  }
  const seen = new Set();
  sheets.Giocatori = sheets.Giocatori.filter(p => {
    if (!p.Giocatore && p.Nome) p.Giocatore = (!p.Cognome || p.Cognome === "?") ? p.Nome : `${p.Nome[0]}. ${p.Cognome}`;
    if (!p.Giocatore) return false;
    if (!p.Attivo) p.Attivo = "Sì";
    if (seen.has(p.Giocatore)) warnings.push(`Giocatori: "${p.Giocatore}" compare due volte.`);
    seen.add(p.Giocatore);
    return true;
  });
  const teams = new Set(sheets.Squadre.map(t => t.Squadra).filter(Boolean));
  sheets.Risultati_Girone.forEach((m, i) => {
    for (const k of ["Casa", "Trasferta"]) if (m[k] && !teams.has(m[k]))
      warnings.push(`Risultati_Girone, riga ${i + 2}: squadra "${m[k]}" non presente nel foglio Squadre.`);
    if ((m.Gol_Casa == null) !== (m.Gol_Trasferta == null))
      warnings.push(`Risultati_Girone, riga ${i + 2}: c'è un solo punteggio (${m.Casa}-${m.Trasferta}), la partita non viene contata.`);
  });
  for (const sh of ["Tabellino", "Partitelle"])
    sheets[sh].forEach((r, i) => { if (r.Giocatore && !seen.has(r.Giocatore))
      warnings.push(`${sh}, riga ${i + 2}: giocatore "${r.Giocatore}" non presente nel foglio Giocatori.`); });
  sheets.Partitelle.forEach((r, i) => { if (r.Risultato && !ESITI.includes(r.Risultato))
    warnings.push(`Partitelle, riga ${i + 2}: risultato "${r.Risultato}" non valido (usa ${ESITI.join(" / ")}).`); });
  return { sheets, warnings };
}
