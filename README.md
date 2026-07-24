# crm-cdg-app

App browser del CRM CDG. Database Excel su Dropbox, hosting gratuito su GitHub Pages,
automazione su GitHub Actions. Progetto in **testing**.

> Regola di sicurezza: questo repository è **pubblico**. Contiene **solo codice**.
> Mai dati clienti, mai il refresh token, mai l'app secret. I dati stanno su Dropbox;
> i segreti nei GitHub Secrets. Il file `.gitignore` blocca per sicurezza xlsx/json/token.

## Come funziona

- **Database di record**: `crm-database.xlsx` nella App Folder di Dropbox (`Apps/CRM CDG - GitHub/`).
- **Automazione** (`.github/workflows/build-snapshot.yml`): ogni 15 minuti GitHub Actions esegue
  `scripts/build_snapshot.py`, che legge il database da Dropbox e ricarica `snapshot.json` nella App Folder.
- **App** (`index.html`): l'utente accede a Dropbox (login PKCE, senza segreti), scarica `snapshot.json`
  e lo visualizza. Senza login mostra dati dimostrativi per verificare il layout.

## Setup (una tantum)

### 1. Dropbox — aggiungi il Redirect URI
App Console → app "CRM CDG - GitHub" → Settings → OAuth 2 → Redirect URIs → aggiungi esattamente:
```
https://pigironcompany-cdg.github.io/crm-cdg-app/
```

### 2. GitHub — aggiungi i Secrets
Repo → Settings → Secrets and variables → Actions → New repository secret:
- `DROPBOX_APP_KEY` = `jeihle42vxibgiv`
- `DROPBOX_REFRESH_TOKEN` = *(il refresh token generato con tools/dropbox-refresh-token.html)*

### 3. GitHub — attiva Pages
Repo → Settings → Pages → Source = **Deploy from a branch** → Branch = **main** / **/ (root)** → Save.
L'app sarà su `https://pigironcompany-cdg.github.io/crm-cdg-app/`.

### 4. Dropbox — carica il database (seed iniziale)
Metti `crm-database.xlsx` nella App Folder `Apps/CRM CDG - GitHub/` (una volta sola).

### 5. Genera il primo snapshot
Repo → Actions → "Build snapshot da Dropbox" → **Run workflow**. Poi apri l'app e accedi a Dropbox.

## Struttura

```
crm-cdg-app/
  index.html                    app browser (login Dropbox + vista clienti)
  config.js                     configurazione pubblica (app key, redirect, percorsi)
  scripts/
    build_snapshot.py           Dropbox -> snapshot.json (usato da Actions)
    requirements.txt
  .github/workflows/
    build-snapshot.yml          automazione schedulata
  .gitignore
```

## Roadmap

- **v0.1 (questa)**: vista read-only multi-utente. Snapshot generato da Actions.
- **v0.2**: scrittura dall'app (note, follow-up) → log su Dropbox → reconcile nel database via Actions.
- **v0.3**: consolidamento script, dashboard, gestione multi-utente su cartella condivisa.

## Note tecniche

- Le app Dropbox "App Folder" danno a ogni utente una propria cartella: in fase multi-utente
  reale si valuterà l'accesso a una cartella condivisa. Per il testing (utente singolo) va bene così.
- Il database non viene mai scritto a mano dagli utenti: è di competenza degli script (regola del progetto).
