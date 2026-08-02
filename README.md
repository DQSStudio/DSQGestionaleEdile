# Software di Computazione Edile — Desearq Studio

## Come pubblicarlo su Netlify

### Opzione A — la più veloce (drag & drop, senza GitHub)
1. Sul tuo computer, dentro questa cartella esegui:
   ```
   npm install
   npm run build
   ```
   Si crea una cartella `dist/`.
2. Vai su https://app.netlify.com → apri il progetto **gestionale-edile** (già creato in precedenza nel team DQS Studio) → **Deploys** → trascina la cartella `dist` nella zona di drag & drop.
3. Fatto: il sito si aggiorna subito con l'ultima versione.

### Opzione B — collegato a GitHub (per aggiornamenti automatici futuri)
1. Crea un repository GitHub e carica tutti i file di questa cartella.
2. Su Netlify → il progetto **gestionale-edile** → *Site configuration* → *Link repository* → collega il repo.
3. Build command: `npm run build` — Publish directory: `dist` (già preconfigurato in `netlify.toml`).
4. Da quel momento ogni push su GitHub ripubblica automaticamente il sito.

## Sviluppo locale
```
npm install
npm run dev
```
Apre l'app su http://localhost:5173

## Cosa contiene
Tutta la piattaforma sviluppata finora in un unico componente (`src/App.jsx`):
Dashboard, Listino prezzi (multiplo, con prezzo impresa/cliente e formule),
Progetti con computi versionati e drag&drop dal listino, Confronto revisioni,
Fornitori con confronto prezzi, Team, esportazione Excel/PDF reale, IVA configurabile.

## Nota
Il file `src/App.jsx` è volutamente un unico grande componente (nato come prototipo
interattivo). Funziona perfettamente così, ma se in futuro volete continuare a
svilupparlo con più persone in parallelo, conviene suddividerlo in file più piccoli
per cartella (components/, pages/) — ditemelo quando volete farlo.
