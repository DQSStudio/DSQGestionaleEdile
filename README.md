# Software di Computazione Edile — Desearq Studio

Piattaforma per listino prezzi, computi metrici, progetti e fornitori,
con dati salvati su un database reale (Supabase) e accesso vero per tutto il team.

## Prima pubblicazione — cosa fare in ordine

### 1. Carica i file su GitHub
Se stai aggiornando un repository già esistente, questi sono i file cambiati rispetto
all'ultima volta — aggiornali tutti:
- `package.json` (aggiunta la libreria `@supabase/supabase-js`)
- `src/supabaseClient.js` (nuovo file — collega l'app al database)
- `src/App.jsx` (login, Team e salvataggio dati ora reali)

Gli altri file (`index.html`, `src/main.jsx`, `netlify.toml`) restano invariati.

### 2. Collega a Netlify (se non l'hai già fatto)
Vedi la sezione "Come pubblicarlo su Netlify" più sotto.

### 3. Al primo avvio del sito online
1. Vai sul sito pubblicato, clicca **"Crea un account"**
2. Registrati con **nicola@desearq.com** e una password a tua scelta (minimo 6 caratteri)
   → diventi automaticamente **admin** (sei la prima persona a registrarti)
3. Se Supabase chiede conferma email, controlla la posta e conferma prima di accedere
4. Da **Team**, invita le altre persone (nome, email, ruolo Admin/Membro)
5. Ognuna di loro, sul sito, clicca **"Crea un account"** e si registra con **la stessa email**
   con cui è stata invitata → viene collegata automaticamente al workspace

Da questo momento: dati salvati per sempre, visibili a tutto il team, refresh della
pagina non disconnette più nessuno.

## Come pubblicarlo su Netlify

### Opzione A — la più veloce (drag & drop, senza GitHub)
1. Sul tuo computer, dentro questa cartella esegui:
   ```
   npm install
   npm run build
   ```
   Si crea una cartella `dist/`.
2. Vai su https://app.netlify.com/projects/gestionale-edile → **Deploys** → trascina la cartella `dist`.
3. Fatto: il sito si aggiorna subito.

### Opzione B — collegato a GitHub (aggiornamenti automatici)
1. Carica tutti i file di questa cartella su un repository GitHub.
2. Su Netlify → il progetto **gestionale-edile** → *Site configuration* → *Link repository* → collega il repo.
3. Build command: `npm run build` — Publish directory: `dist` (già in `netlify.toml`).
4. Da quel momento ogni aggiornamento su GitHub ripubblica automaticamente il sito.

## Sviluppo locale
```
npm install
npm run dev
```
Apre l'app su http://localhost:5173

## Database (Supabase)
- Progetto: **Desearq Studio Manager**, schema dedicato **`cea`** (separato dagli altri
  programmi che condividono lo stesso account Supabase)
- Tabelle: `cea.team_members` (persone e ruoli, collegate all'autenticazione reale),
  `cea.app_state` (un documento unico con progetti, listini e fornitori)
- Login gestito da Supabase Auth: password vere, sessione persistente, nessun dato finto

## Cosa contiene la piattaforma
Dashboard, Listino prezzi (multiplo, con prezzo impresa/cliente e formule),
Progetti con computi versionati, drag&drop dal listino, sommatorie parziali,
macrosezioni personalizzate, IVA configurabile per sezione, Confronto revisioni,
Fornitori con confronto prezzi collegato ai computi, import da Excel, planimetrie
con punti collegati al listino, Team con ruoli, esportazione Excel/PDF reale.

## Nota
`src/App.jsx` è un unico grande componente (nato come prototipo interattivo).
Funziona bene così; se in futuro serve più organizzazione per lavorarci in più
persone in parallelo, si può suddividere in più file — basta chiederlo.
