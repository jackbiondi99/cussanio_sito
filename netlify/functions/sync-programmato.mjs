// Aggiornamento automatico da Tuttocampo.
// Orari in UTC: 05:00 e 21:00 di sabato, domenica, lunedì (07:00 / 23:00 ora italiana estiva).
import { syncTuttocampo } from "../lib/common.mjs";

export default async () => {
  try {
    const m = await syncTuttocampo();
    console.log("Sync Tuttocampo:", JSON.stringify({ pagine: m.pages, aggiornati: m.updated.length, conflitti: m.conflicts.length, errori: m.errors }));
  } catch (e) { console.error("Sync Tuttocampo fallito:", e.message); }
};
export const config = { schedule: "0 5,21 * * 0,1,6" };
