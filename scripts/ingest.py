#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ingest.py — Ingestione caricamenti (GitHub Actions, 10:00 e 18:00).
Controlla /caricamenti nella App Folder. Se ci sono file nuovi:
  - riconosce il tipo (esposizione / uscite annuali / altro),
  - estrae i dati e li consolida in UN UNICO file centrale: /operativa.json,
  - archivia TUTTI i file in /caricamenti/archivio/AAAA-MM-GG/ (storico/backup),
  - svuota /caricamenti.
Lo snapshot viene poi rigenerato da build_snapshot.py (che legge /operativa.json).

Env: DROPBOX_APP_KEY, DROPBOX_REFRESH_TOKEN
"""
import os, io, json, datetime, sys, requests, openpyxl

APP_KEY = os.environ["DROPBOX_APP_KEY"]; REFRESH = os.environ["DROPBOX_REFRESH_TOKEN"]
CARIC = "/caricamenti"; OPERATIVA = "/operativa.json"
TOKEN_URL="https://api.dropbox.com/oauth2/token"
DL="https://content.dropboxapi.com/2/files/download"
UL="https://content.dropboxapi.com/2/files/upload"
LIST="https://api.dropboxapi.com/2/files/list_folder"
MOVE="https://api.dropboxapi.com/2/files/move_v2"
MESI={"gennaio":"01","febbraio":"02","marzo":"03","aprile":"04","maggio":"05","giugno":"06",
      "luglio":"07","agosto":"08","settembre":"09","ottobre":"10","novembre":"11","dicembre":"12"}

def tok():
    r=requests.post(TOKEN_URL,data={"grant_type":"refresh_token","refresh_token":REFRESH,"client_id":APP_KEY},timeout=30)
    r.raise_for_status(); return r.json()["access_token"]
def listing(t,path):
    r=requests.post(LIST,headers={"Authorization":"Bearer "+t,"Content-Type":"application/json"},
        data=json.dumps({"path":path,"recursive":False}),timeout=60)
    if r.status_code!=200: return []
    return [e for e in r.json().get("entries",[]) if e.get(".tag")=="file"]
def dl(t,path):
    r=requests.post(DL,headers={"Authorization":"Bearer "+t,"Dropbox-API-Arg":json.dumps({"path":path})},timeout=120)
    return r.content if r.status_code==200 else None
def ul(t,path,data):
    r=requests.post(UL,headers={"Authorization":"Bearer "+t,
        "Dropbox-API-Arg":json.dumps({"path":path,"mode":"overwrite","mute":True}),
        "Content-Type":"application/octet-stream"},data=data,timeout=120); r.raise_for_status()
def move(t,frm,to):
    return requests.post(MOVE,headers={"Authorization":"Bearer "+t,"Content-Type":"application/json"},
        data=json.dumps({"from_path":frm,"to_path":to,"autorename":True}),timeout=60)

def n(v):
    try: return round(float(v),2)
    except: return None

def parse_esposizione(b):
    wb=openpyxl.load_workbook(io.BytesIO(b),data_only=True)
    righe=[]; giac=None
    if "GHISA" in wb.sheetnames:
        for row in wb["GHISA"].iter_rows(values_only=True):
            lab=row[0]
            if isinstance(lab,str) and lab.strip():
                l=lab.strip()
                if "GIACENZA COMPLESSIVA" in l.upper(): giac=n(row[1]) if len(row)>1 else None; continue
                usd=n(row[1]) if len(row)>1 else None; eur=n(row[2]) if len(row)>2 else None
                if usd is not None or eur is not None: righe.append({"voce":l,"usd":usd,"eur":eur})
    giacenze=[]
    if "Giacenze" in wb.sheetnames:
        for row in wb["Giacenze"].iter_rows(min_row=2,values_only=True):
            if row and row[0] and str(row[0]).strip():
                giacenze.append({"materiale":str(row[0]).strip(),"origine":str(row[1] or "").strip(),
                    "ton":n(row[2]),"prezzo_usd":n(row[3]),"prezzo_eur":n(row[4]),"val_usd":n(row[5]),"val_eur":n(row[6])})
    wb.close()
    return {"esposizione":{"righe":righe,"giacenza_ton":giac}, "giacenze":giacenze}

def parse_uscite(b):
    wb=openpyxl.load_workbook(io.BytesIO(b),data_only=True)
    year=str(datetime.date.today().year)
    ws=wb[year] if year in wb.sheetnames else wb.worksheets[-1]
    out={m:0 for m in MESI.values()}
    for row in ws.iter_rows(values_only=True):
        if row and isinstance(row[0],str) and row[0].strip().lower() in MESI:
            out[MESI[row[0].strip().lower()]]=n(row[1]) or 0
    wb.close()
    return [{"mese":m,"tonnellate":out[m]} for m in sorted(out)]

def classify(name):
    u=name.upper()
    if u.endswith(".XLSX") and "ESPOSIZIONE" in u: return "esposizione"
    if u.endswith(".XLSX") and ("USCITE" in u or "USCITEANNUALI" in u): return "uscite"
    return "altro"

def main():
    t=tok()
    files=listing(t,CARIC)
    if not files:
        print("Cartella caricamenti vuota: nulla da fare."); return 0
    print(f"File trovati in caricamenti: {len(files)}")
    # operativa esistente (aggiorno solo le parti fornite)
    cur=dl(t,OPERATIVA); op=json.loads(cur) if cur else {}
    op["aggiornato"]=datetime.date.today().isoformat()
    changed=False
    for e in files:
        name=e["name"]; path=e["path_lower"]; kind=classify(name)
        print(f" - {name} -> {kind}")
        if kind=="esposizione":
            d=parse_esposizione(dl(t,path)); op["esposizione"]=d["esposizione"]; op["giacenze"]=d["giacenze"]; changed=True
        elif kind=="uscite":
            op["uscite_mensili"]=parse_uscite(dl(t,path)); changed=True
    if changed:
        ul(t,OPERATIVA,json.dumps(op,ensure_ascii=False).encode("utf-8"))
        print("operativa.json aggiornato.")
    # archivia TUTTO e svuota caricamenti
    day=datetime.date.today().isoformat()
    for e in files:
        r=move(t,e["path_lower"],f"{CARIC}/archivio/{day}/{e['name']}")
        if r.status_code!=200: print("   ! archiviazione fallita:",e["name"],r.status_code,r.text[:150])
    print(f"Archiviati {len(files)} file in {CARIC}/archivio/{day}/ e caricamenti svuotata.")
    return 0

if __name__=="__main__":
    try: sys.exit(main())
    except requests.HTTPError as ex:
        print("Errore HTTP Dropbox:",ex.response.status_code,ex.response.text[:300]); sys.exit(1)
