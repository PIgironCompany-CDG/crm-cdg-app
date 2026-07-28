#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
backup_daily.py — Backup giornaliero (GitHub Actions, la sera).
Copia i file chiave della App Folder in una cartella datata:
  /backup/AAAA-MM-GG/<file>
Usa la copia server-side di Dropbox (nessun download/upload). Idempotente: se il
backup del giorno esiste già, lo salta (per-file).

Env (GitHub Secrets): DROPBOX_APP_KEY, DROPBOX_REFRESH_TOKEN
"""
import os, json, datetime, sys, requests

APP_KEY = os.environ["DROPBOX_APP_KEY"]
REFRESH = os.environ["DROPBOX_REFRESH_TOKEN"]
FILES = ["/crm-database.xlsx", "/snapshot.json", "/interazioni.json", "/edits.jsonl", "/edits-applied.jsonl"]
TOKEN_URL = "https://api.dropbox.com/oauth2/token"
COPY_URL  = "https://api.dropboxapi.com/2/files/copy_v2"

def tok():
    r = requests.post(TOKEN_URL, data={"grant_type":"refresh_token","refresh_token":REFRESH,"client_id":APP_KEY}, timeout=30)
    r.raise_for_status(); return r.json()["access_token"]

def copy(t, frm, to):
    return requests.post(COPY_URL, headers={"Authorization":"Bearer "+t, "Content-Type":"application/json"},
        data=json.dumps({"from_path":frm, "to_path":to, "autorename":False}), timeout=60)

def main():
    t = tok()
    day = datetime.date.today().isoformat()
    done, skipped, missing = [], [], []
    for f in FILES:
        to = f"/backup/{day}{f}"
        r = copy(t, f, to)
        if r.status_code == 200:
            done.append(f)
        else:
            err = r.text
            if "to/conflict" in err or "already" in err:
                skipped.append(f)               # backup del giorno già presente
            elif "from_lookup/not_found" in err:
                missing.append(f)               # file non ancora esistente (es. edits vuoti)
            else:
                print("Attenzione su", f, r.status_code, err[:200])
    print(f"Backup {day}: copiati {len(done)}, già presenti {len(skipped)}, assenti {len(missing)}.")
    print("  copiati:", ", ".join(done) or "-")
    if missing: print("  assenti:", ", ".join(missing))
    return 0

if __name__ == "__main__":
    try: sys.exit(main())
    except requests.HTTPError as e:
        print("Errore HTTP Dropbox:", e.response.status_code, e.response.text[:300]); sys.exit(1)
