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
import os, io, json, datetime, sys, re, zipfile, subprocess, tempfile, requests, openpyxl

APP_KEY=os.environ["DROPBOX_APP_KEY"]; REFRESH=os.environ["DROPBOX_REFRESH_TOKEN"]
CARIC="/caricamenti"; OPERATIVA="/operativa.json"; DBX="/crm-database.xlsx"
CREDIT=[("Rating credito","rating",False),("Punteggio credito","punteggio",True),
        ("Limite credito report (€)","limite",True),("Proprietà","prop",False),
        ("Segnalazioni credito","segnalazioni",False),("Fatturato bilancio (€)","fatt_stim",True)]
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
    if u.endswith(".PDF") and re.match(r'[A-Z]{2,4}\d{3}', u): return "bilancio"   # CODICE - Ragione.pdf
    if u.endswith(".XLSX"):
        if "ESPOSIZIONE" in u: return "esposizione"
        if "MAGAZZINO" in u or "INVENDUTO" in u: return "magazzino"
        if "USCITE" in u: return "uscite"
    return "altro"

# ---------- BILANCI (report reportaziende.it) ----------
def code_from_name(name):
    m=re.match(r'\s*([A-Za-z]{2,4}\d{3})', name); return m.group(1).upper() if m else None
def colletter(i):
    i+=1; s=""
    while i>0: i,r=divmod(i-1,26); s=chr(65+r)+s
    return s
def val_eur(s):
    m=re.search(r'([\d.,]+)\s*(Mln|Mld|Md|Mrd|K)?', s or "")
    if not m: return None
    try: v=float(m.group(1).replace('.','').replace(',','.'))
    except: return None
    u=(m.group(2) or '').lower()
    if u=='mln': v*=1_000_000
    elif u in ('mld','md','mrd'): v*=1_000_000_000
    elif u=='k': v*=1000
    return round(v)
def parse_report(pdf_bytes):
    with tempfile.NamedTemporaryFile(suffix=".pdf",delete=False) as f: f.write(pdf_bytes); path=f.name
    try: t=subprocess.run(["pdftotext","-layout",path,"-"],capture_output=True,text=True).stdout
    finally:
        try: os.unlink(path)
        except: pass
    lines=[l.rstrip() for l in t.splitlines()]; j="\n".join(lines); out={}
    m=re.search(r'Valutazione complessiva\s*\n\s*([A-D][+-]?)\s*\n', j); out["rating"]=m.group(1) if m else None
    m=re.search(r'(\d{1,3})\s*/\s*100', j); out["punteggio"]=int(m.group(1)) if m else None
    m=re.search(r'Limite di credito\s*\n\s*€\s*([\d.,]+\s*(?:Mln|Mld|Md|K)?)', j) or re.search(r'Limite di credito\s*€\s*([\d.,]+\s*(?:Mln|Mld|Md|K)?)', j)
    out["limite"]=val_eur(m.group(1)) if m else None
    m=re.search(r'Fatturato Stimato\s*\n?\s*€\s*([\d.,]+\s*(?:Mln|Mld|Md|K)?)', j); out["fatt_stim"]=val_eur(m.group(1)) if m else None
    prop=None
    for l in lines:
        mm=re.search(r'(?:A\.U\.|AMMINISTRATORE UNICO|PRESIDENTE[^A-Z]*|Pres\.?\s*CdA|LEGALE RAPPRESENTANTE)\s{2,}([A-ZÀ-Ù][A-ZÀ-Ù \.&\']{3,})\s*$', l)
        if mm: prop=mm.group(1).strip(); break
    out["prop"]=prop
    bad=[]
    if "PRESENZA DI PROTESTI" in j: bad.append("protesti")
    if "PRESENZA DI PREGIUDIZIEVOLI" in j: bad.append("pregiudizievoli")
    if "PRESENZA DI PROCEDURE" in j: bad.append("procedure")
    out["segnalazioni"]=("PRESENTI: "+", ".join(bad)) if bad else "nessuna"
    return out
def master_index(xlsx_bytes):
    wb=openpyxl.load_workbook(io.BytesIO(xlsx_bytes),read_only=True,data_only=True)
    M=wb["Anagrafica_Master"]; H=[c.value for c in M[1]]; HX={h:i for i,h in enumerate(H) if h}
    rows={}
    for ri,row in enumerate(M.iter_rows(min_row=2,values_only=True),2):
        c=row[HX["Codice"]] if "Codice" in HX else None
        if c: rows[str(c).strip()]=ri
    wb.close(); return HX, rows
def apply_master_credit(xlsx_bytes, updates):
    z=zipfile.ZipFile(io.BytesIO(xlsx_bytes)); parts={n:z.read(n) for n in z.namelist()}; infos=z.infolist(); z.close()
    wbx=parts['xl/workbook.xml'].decode()
    rid=re.search(r'<sheet name="Anagrafica_Master"[^>]*r:id="(rId\d+)"',wbx).group(1)
    tgt=re.search(r'Id="%s"[^>]*Target="([^"]*)"'%rid,parts['xl/_rels/workbook.xml.rels'].decode()).group(1)
    sf="xl/"+tgt.replace("\\","/"); s=parts[sf].decode()
    def esc(v): return str(v).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    def cidx(l):
        n=0
        for ch in l: n=n*26+(ord(ch)-64)
        return n
    def cx(ref,val,isnum): return f'<c r="{ref}"><v>{val}</v></c>' if isnum else f'<c r="{ref}" t="inlineStr"><is><t>{esc(val)}</t></is></c>'
    for rn,ups in updates.items():
        m=re.search(r'(<row r="%d"[^>]*>)(.*?)(</row>)'%rn,s,re.DOTALL)
        if not m: continue
        nodes=re.findall(r'<c r="[A-Z]+%d"[^>]*?(?:/>|>.*?</c>)'%rn,m.group(2),re.DOTALL)
        cm={re.match(r'<c r="([A-Z]+)\d+"',nd).group(1):nd for nd in nodes}
        for col,val,isnum in ups: cm[col]=cx(f"{col}{rn}",val,isnum)
        s=s[:m.start()]+m.group(1)+"".join(cm[c] for c in sorted(cm,key=cidx))+m.group(3)+s[m.end():]
    parts[sf]=s.encode()
    parts['xl/workbook.xml']=re.sub(r'(<calcPr[^/]*?)(/>)',lambda mm:mm.group(1)+(' fullCalcOnLoad="1"' if 'fullCalcOnLoad' not in mm.group(1) else '')+mm.group(2),wbx,count=1).encode()
    if 'xl/calcChain.xml' in parts:
        parts.pop('xl/calcChain.xml',None); infos=[i for i in infos if i.filename!='xl/calcChain.xml']
        parts['[Content_Types].xml']=re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>','',parts['[Content_Types].xml'].decode()).encode()
        parts['xl/_rels/workbook.xml.rels']=re.sub(r'<Relationship[^>]*Target="calcChain.xml"[^>]*/>','',parts['xl/_rels/workbook.xml.rels'].decode()).encode()
    out=io.BytesIO()
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as zw:
        for i in infos: zw.writestr(i,parts[i.filename])
    return out.getvalue()

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
    # BILANCI: aggiornano i campi credito del cliente nel database (Master) -> si propagano alle schede
    bil=[e for e in files if classify(e["name"])=="bilancio"]
    if bil:
        xlsx=dl(t,DBX)
        if not xlsx:
            print("   ATTENZIONE: crm-database.xlsx non trovato: bilanci non applicati.")
        else:
            HX,rows=master_index(xlsx); updates={}
            for e in bil:
                cod=code_from_name(e["name"])
                if not cod or cod not in rows: print(f"   bilancio ignorato (cliente non trovato): {e['name']}"); continue
                rep=parse_report(dl(t,e["path_lower"])); ups=[]
                for col,key,isnum in CREDIT:
                    v=rep.get(key)
                    if col in HX and v not in (None,""): ups.append((colletter(HX[col]),v,isnum))
                if ups: updates[rows[cod]]=ups; print(f"   bilancio {cod}: {len(ups)} campi credito")
            if updates:
                ul(t,DBX,apply_master_credit(xlsx,updates)); print(f"crm-database.xlsx aggiornato: credito di {len(updates)} clienti.")
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
