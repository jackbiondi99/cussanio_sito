# Sito Cussanio – istruzioni

## Contenuto
- `public/index.html` – il sito pubblico
- `public/admin.html` – area riservata (indirizzo: tuosito.netlify.app/admin.html)
- `public/xlsx.js` – lettura/scrittura dei file Excel nel browser
- `netlify/functions/` – funzioni lato server (dati, accesso, area riservata, aggiornamento automatico)
- `netlify/lib/common.mjs` – codice condiviso (archivio dati, lettura di Tuttocampo)

## Messa online (una volta sola)
1. Crea un account gratuito su github.com e un repository nuovo (anche privato), es. `cussanio-sito`.
2. Nel repository: "Add file" > "Upload files" e trascina TUTTO il contenuto di questa cartella
   (file e sottocartelle). Conferma con "Commit changes".
3. Su Netlify, apri il sito esistente > Site configuration > Build & deploy > Continuous deployment >
   "Link repository" e scegli il repository (così l'indirizzo del sito resta lo stesso).
   Nessun comando di build: le impostazioni sono già in `netlify.toml`.
4. Site configuration > Environment variables > Add a variable:
   - `ADMIN_PASSWORD` = la password dell'area riservata (scegline una robusta)
5. Deploys > Trigger deploy > Deploy site.
6. Apri `/admin.html`, entra con la password, tab "Dati Excel" > carica il tuo file Excel attuale.
7. Tab "Tuttocampo" > "Aggiorna adesso" per verificare che la lettura funzioni.

## Uso
- Risultati e date: si aggiornano da soli (sab, dom, lun alle 7:00 e alle 23:00 circa, ora italiana).
  I risultati già presenti non vengono mai sovrascritti; le differenze compaiono nel tab "Tuttocampo".
- Partitelle: tab "Partitella" dell'area riservata.
- Tutto il resto (rosa, tabellino, correzioni): tab "Dati Excel" > scarica, modifica, ricarica.
  Il file caricato sostituisce tutti i dati; prima viene salvata una copia di sicurezza (Netlify > Blobs).

## Variabili facoltative
- `TUTTOCAMPO_BASE` – indirizzo del girone su Tuttocampo, se cambia
  (default: https://www.tuttocampo.it/Piemonte/TerzaCategoria/GironeBCuneo)
