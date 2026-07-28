#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ingest.py — Ingestione caricamenti (GitHub Actions, 10:00 e 18:00).
Controlla /caricamenti nella App Folder. Se ci sono file nuovi:
  - riconosce il tipo (esposizione / magazzino invenduto / uscite),
  - estrae i dati e li consolida in UN UNICO file centrale: /operativa.json,
  - archivia TUTTI i file in /caricamenti/archivio/AAAA-MM-GG/ (storico/backup),
  - svuota /caricamenti.
Crea /caricamenti se non esiste. Lo snapshot lo rigenera build_snapshot.py.

Env: DROPBOX_APP_KEY, DROPBOX_REFRESH_TOKEN
"""
import os, io, json, datetime, sys, re, requests, openpyxl

APP_KEY=os.environ["DROPBOX_APP_KEY"]; REFRESH=os.environ["DROPBOX_REFRESH_TOKEN"]
CARIC="/caricamenti"; OPERATIVA="/operativa.json"
TOKEN_URL="https://api.dropbox.com/oauth2/token"
DL="https://content.dropboxapi.com/2/files/download"
UL="https://content.dropboxapi.com/2/files/upload"
LIST="https://api.dropboxapi.com/2/files/list_folder"
MOVE="https://api.dropboxapi.com/2/files/move_v2"
MKDIR="https://api.dropboxapi.com/2/files/create_folder_v2"
MESI={"gennaio":"01","febbraio":"02","marzo":"03","aprile":"04","maggio":"05","giugno":"06",
      "luglio":"07","agosto":"08","settembre":"09","ottobre":"10","novembre":"11","dicembre":"12"}

def tok():
    r=requests.post(TOKEN_URL,data={"grant_type":"refresh_token","refresh_token":REFRESH,"client_id":APP_KEY},timeout=30)
    r.raise_for_status(); return r.json()["access_token"]
def listing(t,path):
    r=requests.post(LIST,headers={"Authorization":"Bearer "+t,"Content-Type":"application/json"},
        data=json.dumps({"path":path,"recursive":False}),timeout=60)
    return [e for e in r.json().get("entries",[]) if e.get(".tag")=="file"] if r.status_code==200 else []
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
def mkdir(t,path):
    requests.post(MKDIR,headers={"Authorization":"Bearer "+t,"Content-Type":"application/json"},
        data=json.dumps({"path":path}),timeout=30)

def num(v):
    try: return round(float(v),2)
    except: return None
def cs(row): return [("" if c is None else str(c).strip()) for c in row]

def parse_esposizione(b):
    wb=openpyxl.load_workbook(io.BytesIO(b),data_only=True); righe=[]; giac=None
    if "GHISA" in wb.sheetnames:
        for row in wb["GHISA"].iter_rows(values_only=True):
            lab=row[0]
            if isinstance(lab,str) and lab.strip():
                l=lab.strip()
                if "GIACENZA COMPLESSIVA" in l.upper(): giac=num(row[1]) if len(row)>1 else None; continue
                u=num(row[1]) if len(row)>1 else None; e=num(row[2]) if len(row)>2 else None
                if u is not None or e is not None: righe.append({"voce":l,"usd":u,"eur":e})
    wb.close()
    return {"righe":righe,"giacenza_ton":giac}

def parse_magazzino(b):
    wb=openpyxl.load_workbook(io.BytesIO(b),data_only=True); voci=[]; usc={m:0 for m in MESI.values()}
    for ws in wb.worksheets:
        if ws.title.strip().lower().startswith("uscite"):
            for row in ws.iter_rows(values_only=True):
                c=cs(row)
                for i,x in enumerate(c):
                    if x.lower() in MESI and i+1<len(c) and num(c[i+1]) is not None: usc[MESI[x.lower()]]=num(c[i+1])
            continue
        qual=ws.title.strip(); prodotto=nave=prezzo=None
        for row in ws.iter_rows(values_only=True):
            c=cs(row); j=" ".join(c)
            for x in c:
                if re.search(r'PIG IRON|EMATITE|NODULAR|BASIC|FOUNDRY|BM\b',x) and x==x.upper() and len(x)>3 and "€" not in x and "$" not in x: prodotto=x
            for x in c:
                if re.search(r'ENTERPRISE|HERMES|NAREE|TBN|SELECTA|MANX|KARANFIL|IZUMO|ELLAN|MALLIKA|SAGA',x,re.I): nave=x
            for x in c:
                if "/mt" in x.lower(): prezzo=x
            if "disponibilit" in j.lower():
                disp=None
                for i,x in enumerate(c):
                    if "disponibilit" in x.lower():
                        for k in range(i+1,len(c)):
                            if num(c[k]) is not None: disp=num(c[k]); break
                voci.append({"qualita":qual,"prodotto":prodotto,"nave":nave,"prezzo":prezzo,"disponibile":disp})
                nave=prezzo=None
    wb.close()
    return {"magazzino":{"disponibile_totale":round(sum(v["disponibile"] or 0 for v in voci)),"voci":voci},
            "uscite_mensili":[{"mese":m,"tonnellate":usc[m]} for m in sorted(usc)]}

def parse_uscite(b):
    wb=openpyxl.load_workbook(io.BytesIO(b),data_only=True)
    year=str(datetime.date.today().year); ws=wb[year] if year in wb.sheetnames else wb.worksheets[-1]
    usc={m:0 for m in MESI.values()}
    for row in ws.iter_rows(values_only=True):
        c=cs(row)
        if c and c[0].lower() in MESI and len(c)>1 and num(c[1]) is not None: usc[MESI[c[0].lower()]]=num(c[1])
    wb.close()
    return [{"mese":m,"tonnellate":usc[m]} for m in sorted(usc)]

def classify(name):
    u=name.upper()
    if not u.endswith(".XLSX"): return "altro"
    if "ESPOSIZIONE" in u: return "esposizione"
    if "MAGAZZINO" in u or "INVENDUTO" in u: return "magazzino"
    if "USCITE" in u: return "uscite"
    return "altro"

def main():
    t=tok(); mkdir(t,CARIC)
    files=listing(t,CARIC)
    if not files:
        print("Cartella caricamenti vuota: nulla da fare."); return 0
    print(f"File in caricamenti: {len(files)}")
    cur=dl(t,OPERATIVA); op=json.loads(cur) if cur else {}
    op["aggiornato"]=datetime.date.today().isoformat(); changed=False
    for e in files:
        name=e["name"]; kind=classify(name); print(f" - {name} -> {kind}")
        if kind=="esposizione": op["esposizione"]=parse_esposizione(dl(t,e["path_lower"])); changed=True
        elif kind=="magazzino":
            d=parse_magazzino(dl(t,e["path_lower"])); op["magazzino"]=d["magazzino"]; op["uscite_mensili"]=d["uscite_mensili"]; changed=True
        elif kind=="uscite":
            if not op.get("uscite_mensili"): op["uscite_mensili"]=parse_uscite(dl(t,e["path_lower"])); changed=True
    if changed:
        ul(t,OPERATIVA,json.dumps(op,ensure_ascii=False).encode("utf-8")); print("operativa.json aggiornato.")
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
