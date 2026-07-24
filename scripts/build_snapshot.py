#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_snapshot.py — Automazione (GitHub Actions).
Legge il database Excel dalla App Folder di Dropbox, costruisce snapshot.json
(la vista che l'app browser mostra) e lo ricarica nella stessa App Folder.

NON scrive nulla nel repository: i dati restano su Dropbox.

Variabili d'ambiente (dai GitHub Secrets):
  DROPBOX_APP_KEY        app key dell'app Dropbox (pubblica)
  DROPBOX_REFRESH_TOKEN  refresh token generato una tantum (segreto)

Percorsi nella App Folder (relativi alla radice dell'app):
  /crm-database.xlsx   database di record
  /snapshot.json       vista generata per l'app
"""
import os, io, json, datetime, sys
import requests
import openpyxl

APP_KEY = os.environ["DROPBOX_APP_KEY"]
REFRESH_TOKEN = os.environ["DROPBOX_REFRESH_TOKEN"]
DB_PATH = "/crm-database.xlsx"
SNAPSHOT_PATH = "/snapshot.json"

TOKEN_URL = "https://api.dropbox.com/oauth2/token"
DOWNLOAD_URL = "https://content.dropboxapi.com/2/files/download"
UPLOAD_URL = "https://content.dropboxapi.com/2/files/upload"

# mappa: chiave_json -> intestazione colonna nel foglio Anagrafica_Master
FIELDS = {
 "codice":"Codice","ragione":"Ragione sociale","piva":"P.IVA","settore":"Settore",
 "tipologia":"Tipologia","status":"Status","priorita":"Priorità","stato_attivita":"Stato attività",
 "regione":"Regione","provincia":"Provincia","comune":"Comune","indirizzo":"Indirizzo",
 "email":"Email","telefono":"Telefono","referente":"Referente acquisti","proprieta":"Proprietà",
 "owner":"Owner","note":"Note","anno_bilancio":"Anno bilancio","fatturato_bilancio":"Fatturato bilancio (€)",
 "utile":"Utile/Perdita (€)","dipendenti":"N. dipendenti","rating":"Rating credito",
 "punteggio":"Punteggio credito","limite_credito":"Limite credito report (€)",
 "gg_silenzio":"Giorni di silenzio","fatturato_storico":"Fatturato storico (€)","n_ordini":"N. ordini",
}

def get_access_token():
    r = requests.post(TOKEN_URL, data={
        "grant_type": "refresh_token",
        "refresh_token": REFRESH_TOKEN,
        "client_id": APP_KEY,
    }, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]

def download(token, path):
    r = requests.post(DOWNLOAD_URL, headers={
        "Authorization": f"Bearer {token}",
        "Dropbox-API-Arg": json.dumps({"path": path}),
    }, timeout=60)
    r.raise_for_status()
    return r.content

def upload(token, path, data: bytes):
    r = requests.post(UPLOAD_URL, headers={
        "Authorization": f"Bearer {token}",
        "Dropbox-API-Arg": json.dumps({"path": path, "mode": "overwrite", "mute": True}),
        "Content-Type": "application/octet-stream",
    }, data=data, timeout=60)
    r.raise_for_status()
    return r.json()

def conv(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    return v

def build(xlsx_bytes: bytes) -> dict:
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
    M = wb["Anagrafica_Master"]; H = [c.value for c in M[1]]; HX = {h:i for i,h in enumerate(H) if h}

    # Registro operativo (foglio Pipeline righe 5-123) = fonte affidabile dei follow-up
    reg = {}
    try:
        P = wb["Pipeline"]
        for row in P.iter_rows(min_row=5, max_row=123, values_only=True):
            if row and row[0]:
                reg[str(row[0]).strip()] = {
                    "ultimo": row[6], "esito": row[8], "prossima": row[9],
                    "followup": row[10], "ntent": row[11], "stato": row[12],
                }
    except Exception:
        pass

    clients = []
    for row in M.iter_rows(min_row=2, values_only=True):
        cod = row[HX["Codice"]] if "Codice" in HX else None
        if not cod:
            continue
        rec = {}
        for k, col in FIELDS.items():
            i = HX.get(col)
            rec[k] = conv(row[i]) if (i is not None and i < len(row)) else None
        rg = reg.get(str(cod).strip())
        if rg:
            rec["esito"] = conv(rg.get("esito"))
            rec["prossima_azione"] = conv(rg.get("prossima"))
            rec["data_ultima_azione"] = conv(rg.get("ultimo"))
            rec["data_followup"] = conv(rg.get("followup"))
            rec["n_tentativi"] = rg.get("ntent")
            rec["stato_followup"] = conv(rg.get("stato"))
        clients.append(rec)
    wb.close()
    return {
        "generato": datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
        "n_clienti": len(clients),
        "clienti": clients,
    }

def main():
    token = get_access_token()
    print("Token ottenuto. Scarico il database...")
    xlsx = download(token, DB_PATH)
    print(f"Database scaricato ({len(xlsx)} byte). Costruisco lo snapshot...")
    snap = build(xlsx)
    blob = json.dumps(snap, ensure_ascii=False).encode("utf-8")
    upload(token, SNAPSHOT_PATH, blob)
    print(f"Snapshot pubblicato: {snap['n_clienti']} clienti, {len(blob)} byte.")

if __name__ == "__main__":
    try:
        sys.exit(main())
    except requests.HTTPError as e:
        print("Errore HTTP Dropbox:", e.response.status_code, e.response.text[:500])
        sys.exit(1)
