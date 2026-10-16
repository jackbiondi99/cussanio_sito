// =====================================================================
// Codice condiviso dalle funzioni Netlify del sito Cussanio
// =====================================================================
import { getStore } from "@netlify/blobs";
import crypto from "node:crypto";

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

// ---------------------------------------------------------------------
// Archivio (Netlify Blobs)
// ---------------------------------------------------------------------
const store = () => getStore({ name: "cussanio", consistency: "strong" });
export const loadDB   = async () => (await store().get("db", { type: "json" })) || null;
export const saveDB   = async db => { db.updatedAt = new Date().toISOString(); await store().setJSON("db", db); return db; };
export const loadMeta = async () => (await store().get("sync", { type: "json" })) || null;
export const saveMeta = async m => store().setJSON("sync", m);
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
  if (typeof v === "number" && isFinite(v)) {
    const d = new Date(Math.round((v - 25569) * 864e5));
    return d.toISOString().slice(0, 10);
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
  m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})/);
  if (m) return `${m[3].length === 2 ? "20" + m[3] : m[3]}-${pad(m[2])}-${pad(m[1])}`;
  return undefined; // non riconosciuta
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

  // Giocatori: nome breve ricavato se manca, Attivo di default "Sì"
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

// =====================================================================
// AGGIORNAMENTO AUTOMATICO DA TUTTOCAMPO
// =====================================================================
const BASE = () => (process.env.TUTTOCAMPO_BASE || "https://www.tuttocampo.it/Piemonte/TerzaCategoria/GironeBCuneo").replace(/\/$/, "");
const slug = name => name.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^A-Za-z0-9]/g, "");
const ENT = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " ", agrave: "à", egrave: "è", eacute: "é", igrave: "ì", ograve: "ò", ugrave: "ù", Agrave: "À", Egrave: "È" };
export function htmlToText(html){
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n))
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&([a-z]+);/gi, (m, n) => ENT[n] ?? m)
    .replace(/\s+/g, " ");
}
const reEsc = s => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

// Estrae le partite dal testo di una pagina "Calendario" di squadra.
// Formati riconosciuti (testo della pagina):
//   "1 Cussanio 4 0 Villar 2025 INFO"               -> partita giocata
//   "3 Cussanio - 09/10 - Villafalletto INFO"        -> partita da giocare
export function parseCalendar(text, teamNames){
  const alt = teamNames.slice().sort((a, b) => b.length - a.length).map(reEsc).join("|");
  const out = [];
  const played = new RegExp(`(?:^|\\s)(\\d{1,2})\\s+(${alt})\\s+(\\d{1,2})\\s*-?\\s*(\\d{1,2})\\s+(${alt})(?=\\s|$)`, "g");
  const todo   = new RegExp(`(?:^|\\s)(\\d{1,2})\\s+(${alt})\\s+-\\s+(\\d{1,2})\\/(\\d{1,2})(?:\\/(\\d{2,4}))?\\s+-\\s+(${alt})(?=\\s|$)`, "g");
  let m;
  while ((m = played.exec(text))) out.push({ g: +m[1], home: m[2], away: m[5], gh: +m[3], ga: +m[4] });
  while ((m = todo.exec(text)))   out.push({ g: +m[1], home: m[2], away: m[6], day: +m[3], month: +m[4], year: m[5] ? +m[5] : null });
  return out.filter(x => x.home !== x.away && x.g >= 1 && x.g <= 60);
}

async function fetchText(url){
  const ctl = new AbortController();
  const t = setTimeout(() => ctl.abort(), 12000);
  try {
    const r = await fetch(url, { signal: ctl.signal, headers: { "user-agent": "Mozilla/5.0 (compatible; CussanioStats/1.0)", "accept-language": "it-IT,it;q=0.9" } });
    if (!r.ok) throw new Error("HTTP " + r.status);
    return await r.text();
  } finally { clearTimeout(t); }
}

export async function syncTuttocampo({ fetcher = fetchText } = {}){
  const db = await loadDB();
  if (!db) throw new HttpError(409, "Nessun dato nel sito: carica prima il file Excel dall'area amministratore.");
  const S = db.sheets;
  const teams = S.Squadre.filter(t => t.Squadra && t.ID_Tuttocampo);
  const names = S.Squadre.map(t => t.Squadra).filter(Boolean);
  const meta = { at: new Date().toISOString(), ok: false, pages: 0, pagesFailed: [], updated: [], conflicts: [], added: [], errors: [] };

  // 1. Lettura delle pagine (in parallelo, il limite delle funzioni programmate è 30 s)
  const pages = await Promise.all(teams.map(async t => {
    const url = `${BASE()}/Squadra/${slug(t.Squadra)}/${t.ID_Tuttocampo}/Calendario`;
    try { return parseCalendar(htmlToText(await fetcher(url)), names); }
    catch (e) { meta.pagesFailed.push(`${t.Squadra}: ${e.message}`); return null; }
  }));
  meta.pages = pages.filter(Boolean).length;
  const found = pages.filter(Boolean).flat();
  if (!meta.pages) { meta.errors.push("Nessuna pagina di Tuttocampo è stata letta."); await saveMeta(meta); return meta; }
  if (!found.length) { meta.errors.push("Pagine lette, ma nessuna partita riconosciuta: probabilmente Tuttocampo ha cambiato formato."); await saveMeta(meta); return meta; }

  // 2. Unione dei dati letti da più pagine (ogni partita compare in due calendari)
  const merged = new Map();
  for (const x of found) {
    const k = `${x.g}|${x.home}|${x.away}`;
    const prev = merged.get(k);
    if (!prev) { merged.set(k, x); continue; }
    if (x.gh != null && prev.gh != null && (x.gh !== prev.gh || x.ga !== prev.ga)) prev.incoerente = true;
    if (x.gh != null && prev.gh == null) merged.set(k, Object.assign(x, { incoerente: prev.incoerente }));
  }

  // anno di inizio stagione: dalla prima data presente, altrimenti dall'anno corrente
  const dates = S.Risultati_Girone.map(r => r.Data).filter(Boolean).sort();
  const y0 = dates.length ? +dates[0].slice(0, 4) : new Date().getFullYear();
  const isoOf = x => `${x.year ? (x.year < 100 ? 2000 + x.year : x.year) : (x.month >= 7 ? y0 : y0 + 1)}-${pad(x.month)}-${pad(x.day)}`;
  const label = r => `G${r.Giornata} ${r.Casa}-${r.Trasferta}`;

  // 3. Applicazione: si riempiono solo i campi vuoti, i risultati già presenti non vengono mai sovrascritti
  let changed = false;
  for (const x of merged.values()) {
    if (x.incoerente) { meta.conflicts.push(`G${x.g} ${x.home}-${x.away}: punteggi diversi tra due pagine di Tuttocampo, non aggiornato.`); continue; }
    let row = S.Risultati_Girone.find(r => r.Giornata === x.g && r.Casa === x.home && r.Trasferta === x.away);
    const swapped = !row && S.Risultati_Girone.find(r => r.Giornata === x.g && r.Casa === x.away && r.Trasferta === x.home);
    const hasScore = r => r.Gol_Casa != null && r.Gol_Trasferta != null;

    if (!row && swapped) {
      if (hasScore(swapped)) {
        if (x.gh != null) meta.conflicts.push(`${label(swapped)}: su Tuttocampo risulta giocata a campi invertiti (${x.home}-${x.away} ${x.gh}-${x.ga}). Non modificato.`);
        continue;
      }
      [swapped.Casa, swapped.Trasferta] = [x.home, x.away];
      meta.updated.push(`${label(swapped)}: corretto casa/trasferta secondo Tuttocampo.`);
      row = swapped; changed = true;
    }
    if (!row) {
      row = { Giornata: x.g, Data: x.day ? isoOf(x) : null, Casa: x.home, Trasferta: x.away, Gol_Casa: null, Gol_Trasferta: null };
      S.Risultati_Girone.push(row);
      meta.added.push(`${label(row)}: partita aggiunta (non era nel calendario).`);
      changed = true;
    }
    if (x.gh != null) {
      if (!hasScore(row)) {
        row.Gol_Casa = x.gh; row.Gol_Trasferta = x.ga;
        meta.updated.push(`${label(row)}: risultato ${x.gh}-${x.ga}.`); changed = true;
      } else if (row.Gol_Casa !== x.gh || row.Gol_Trasferta !== x.ga) {
        meta.conflicts.push(`${label(row)}: nel sito ${row.Gol_Casa}-${row.Gol_Trasferta}, su Tuttocampo ${x.gh}-${x.ga}. Tenuto quello del sito.`);
      }
    } else if (x.day && !hasScore(row)) {
      const iso = isoOf(x);
      if (row.Data !== iso) {
        meta.updated.push(`${label(row)}: data ${row.Data ? row.Data.split("-").reverse().join("/") : "–"} → ${iso.split("-").reverse().join("/")}.`);
        row.Data = iso; changed = true;
      }
    }
  }
  S.Risultati_Girone.sort((a, b) => (a.Giornata - b.Giornata) || String(a.Casa).localeCompare(String(b.Casa)));
  if (changed) { await backupDB(await loadDB(), "prima dell'aggiornamento da Tuttocampo"); await saveDB(db); }
  meta.ok = true; meta.changed = changed;
  await saveMeta(meta);
  return meta;
}
