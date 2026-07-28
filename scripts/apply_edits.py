#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
apply_edits.py — Automazione (GitHub Actions).
Applica le modifiche accodate dall'app (edits.jsonl su Dropbox) dentro il database
Excel (crm-database.xlsx nella App Folder), poi svuota la coda (archiviando gli applicati).

Instrada ogni campo come il vecchio 'sync pipeline.py':
  - follow-up (esito, prossima azione, data follow-up, stato) -> foglio Pipeline (Registro, righe 5-123)
  - contatti/stato (note, referente, email, telefono, tipologia, priorita, stato attivita) -> Anagrafica_Master
Non usa openpyxl in scrittura: chirurgia XML sullo zip; rimuove calcChain; imposta fullCalcOnLoad.
"""
import os, io, json, re, datetime, zipfile
import requests, openpyxl

APP_KEY = os.environ["DROPBOX_APP_KEY"]
REFRESH = os.environ["DROPBOX_REFRESH_TOKEN"]
DB_PATH = "/crm-database.xlsx"
EDITS_PATH = "/edits.jsonl"
ARCHIVE_PATH = "/edits-applied.jsonl"

TOKEN_URL="https://api.dropbox.com/oauth2/token"
DL="https://content.dropboxapi.com/2/files/download"
UL="https://content.dropboxapi.com/2/files/upload"
EPOCH=datetime.date(1899,12,30)

# campo app -> ("REG", colonna_Pipeline) | ("MAS", intestazione_Master); tipo
ROUTE = {
 "esito":("REG","I","str"), "prossima_azione":("REG","J","str"),
 "data_followup":("REG","K","date"), "stato_followup":("REG","M","str"),
 "note":("MAS","Note","str"), "referente":("MAS","Referente acquisti","str"),
 "email":("MAS","Email","str"), "telefono":("MAS","Telefono","str"),
 "tipologia":("MAS","Tipologia","str"), "priorita":("MAS","Priorità","str"),
 "stato_attivita":("MAS","Stato attività","str"),
 # anagrafica (Fase C: campi prima in sola lettura, ora modificabili dall'app)
 "ragione":("MAS","Ragione sociale","str"), "piva":("MAS","P.IVA","str"),
 "settore":("MAS","Settore","str"), "proprieta":("MAS","Proprietà","str"),
 "indirizzo":("MAS","Indirizzo","str"), "comune":("MAS","Comune","str"),
 "provincia":("MAS","Provincia","str"), "regione":("MAS","Regione","str"),
 "owner":("MAS","Owner","str"),
}
NEW_PATH = "/nuovi.jsonl"
NEW_ARCHIVE = "/nuovi-applied.jsonl"

def tok():
    r=requests.post(TOKEN_URL,data={"grant_type":"refresh_token","refresh_token":REFRESH,"client_id":APP_KEY},timeout=30)
    r.raise_for_status(); return r.json()["access_token"]
def dl_bytes(t,path):
    r=requests.post(DL,headers={"Authorization":"Bearer "+t,"Dropbox-API-Arg":json.dumps({"path":path})},timeout=60)
    return r.content if r.status_code==200 else None
def ul_bytes(t,path,data):
    r=requests.post(UL,headers={"Authorization":"Bearer "+t,
        "Dropbox-API-Arg":json.dumps({"path":path,"mode":"overwrite","mute":True}),
        "Content-Type":"application/octet-stream"},data=data,timeout=60)
    r.raise_for_status()

def ser(v):
    if not v: return None
    try: d=datetime.datetime.strptime(str(v)[:10],"%Y-%m-%d").date(); return (d-EPOCH).days
    except: return None
def esc(v): return str(v).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
def colidx(l):
    n=0
    for ch in l: n=n*26+(ord(ch)-64)
    return n
def colletter(i):
    i+=1; s=""
    while i>0: i,r=divmod(i-1,26); s=chr(65+r)+s
    return s

def apply_to_xlsx(xlsx_bytes, agg):
    # agg: {codice: {campo: valore}}
    wb=openpyxl.load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
    M=wb["Anagrafica_Master"]; H=[c.value for c in M[1]]; HX={h:i for i,h in enumerate(H) if h}
    mrow={}
    for ri,row in enumerate(M.iter_rows(min_row=2,values_only=True),2):
        c=row[HX["Codice"]] if "Codice" in HX else None
        if c: mrow[str(c).strip()]=ri
    P=wb["Pipeline"]; prow={}
    for ri,row in enumerate(P.iter_rows(min_row=5,max_row=123,values_only=True),5):
        if row and row[0]: prow[str(row[0]).strip()]=ri
    wb.close()

    z=zipfile.ZipFile(io.BytesIO(xlsx_bytes)); parts={n:z.read(n) for n in z.namelist()}; infos=z.infolist(); z.close()
    wbx=parts['xl/workbook.xml'].decode()
    def sheetfile(name):
        rid=re.search(r'<sheet name="%s"[^>]*r:id="(rId\d+)"'%name,wbx).group(1)
        tgt=re.search(r'Id="%s"[^>]*Target="([^"]*)"'%rid,parts['xl/_rels/workbook.xml.rels'].decode()).group(1)
        return "xl/"+tgt.replace("\\","/")
    sfM=sheetfile("Anagrafica_Master"); sfP=sheetfile("Pipeline")
    sM=parts[sfM].decode(); sP=parts[sfP].decode()
    mK=re.search(r'<c r="K\d+" s="(\d+)"><v>4\d{4}',sP); dP=mK.group(1) if mK else None

    def getcell(body,ref):
        m=re.search(r'<c r="%s"[^>]*?(?:/>|>.*?</c>)'%ref,body,re.DOTALL); return m.group(0) if m else None
    def style_of(cellxml):
        s=re.search(r'\ss="(\d+)"',cellxml or ""); return s.group(1) if s else None
    def setrow(s,rn,cells):
        m=re.search(r'(<row r="%d"[^>]*>)(.*?)(</row>)'%rn,s,re.DOTALL)
        if not m: return s
        nodes=re.findall(r'<c r="[A-Z]+%d"[^>]*?(?:/>|>.*?</c>)'%rn,m.group(2),re.DOTALL)
        cm={re.match(r'<c r="([A-Z]+)\d+"',nd).group(1):nd for nd in nodes}
        for col,xml in cells.items(): cm[col]=xml
        return s[:m.start()]+m.group(1)+"".join(cm[c] for c in sorted(cm,key=colidx))+m.group(3)+s[m.end():]
    def strcell(ref,val,style=None):
        sa=f' s="{style}"' if style else ''
        return f'<c r="{ref}"{sa} t="inlineStr"><is><t>{esc(val)}</t></is></c>'
    def datecell(ref,val,style=None):
        sv=ser(val); sa=f' s="{style}"' if style else ''
        return f'<c r="{ref}"{sa}><v>{sv}</v></c>' if sv is not None else f'<c r="{ref}"{sa}/>'

    Mcells={}; Pcells={}
    for cod,fields in agg.items():
        mr=mrow.get(cod); pr=prow.get(cod)
        for f,val in fields.items():
            if f not in ROUTE: continue
            dest,key,kind=ROUTE[f]
            if dest=="REG":
                if not pr: continue
                ref=f"{key}{pr}"
                if kind=="date":
                    st=style_of(getcell(re.search(r'<row r="%d"[^>]*>.*?</row>'%pr,sP,re.DOTALL).group(0),ref)) or dP
                    Pcells.setdefault(pr,{})[key]=datecell(ref,val,st)
                else:
                    Pcells.setdefault(pr,{})[key]=strcell(ref,val)
            else:
                if not mr: continue
                col=colletter(HX[key]); ref=f"{col}{mr}"
                Mcells.setdefault(mr,{})[col]=strcell(ref,val)
                if f=="tipologia" and pr:  # rispecchia tipologia nel Registro (col E)
                    Pcells.setdefault(pr,{})["E"]=strcell(f"E{pr}",val)
    for rn,cells in Mcells.items(): sM=setrow(sM,rn,cells)
    for rn,cells in Pcells.items(): sP=setrow(sP,rn,cells)
    parts[sfM]=sM.encode(); parts[sfP]=sP.encode()

    parts['xl/workbook.xml']=re.sub(r'(<calcPr[^/]*?)(/>)',
        lambda mm:mm.group(1)+(' fullCalcOnLoad="1"' if 'fullCalcOnLoad' not in mm.group(1) else '')+mm.group(2),wbx,count=1).encode()
    if 'xl/calcChain.xml' in parts:
        parts.pop('xl/calcChain.xml',None); infos=[i for i in infos if i.filename!='xl/calcChain.xml']
        parts['[Content_Types].xml']=re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>','',parts['[Content_Types].xml'].decode()).encode()
        parts['xl/_rels/workbook.xml.rels']=re.sub(r'<Relationship[^>]*Target="calcChain.xml"[^>]*/>','',parts['xl/_rels/workbook.xml.rels'].decode()).encode()
    out=io.BytesIO()
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as zw:
        for i in infos: zw.writestr(i,parts[i.filename])
    return out.getvalue()

def add_clients(xlsx_bytes, nuovi):
    """Aggiunge nuove righe cliente in Anagrafica_Master (chirurgia XML)."""
    wb=openpyxl.load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
    M=wb["Anagrafica_Master"]; H=[c.value for c in M[1]]; HX={h:i for i,h in enumerate(H) if h}
    esistenti=set(); last=1
    for ri,row in enumerate(M.iter_rows(min_row=2,values_only=True),2):
        c=row[HX["Codice"]] if "Codice" in HX else None
        if c: esistenti.add(str(c).strip()); last=ri
    wb.close()
    da_agg=[n for n in nuovi if str(n.get("codice","")).strip() and str(n["codice"]).strip() not in esistenti]
    if not da_agg: return xlsx_bytes, 0

    z=zipfile.ZipFile(io.BytesIO(xlsx_bytes)); parts={n:z.read(n) for n in z.namelist()}; infos=z.infolist(); z.close()
    wbx=parts['xl/workbook.xml'].decode()
    rid=re.search(r'<sheet name="Anagrafica_Master"[^>]*r:id="(rId\d+)"',wbx).group(1)
    tgt=re.search(r'Id="%s"[^>]*Target="([^"]*)"'%rid,parts['xl/_rels/workbook.xml.rels'].decode()).group(1)
    sf="xl/"+tgt.replace("\\","/"); s=parts[sf].decode()
    tabfile=next((n for n in parts if n.startswith("xl/tables/") and b'name="tbl_Master"' in parts[n]),None)

    # mappa campo app -> intestazione Master (solo destinazioni MAS)
    field2col={k:v[1] for k,v in ROUTE.items() if v[0]=="MAS"}
    maxcol=max(HX.values())+1
    rows_xml=[]; r=last+1
    for n in da_agg:
        f=n.get("fields",{}) or {}; cells={}
        def put(colname,val):
            if colname in HX and val not in (None,""):
                cells[colletter(HX[colname])]=strcell_plain(f"{colletter(HX[colname])}{r}",val)
        put("Codice", n["codice"])
        for k,v in f.items():
            if k in field2col: put(field2col[k], v)
        put("Status","Prospect"); put("Stato attività","Da fare")
        if cells:
            rows_xml.append(f'<row r="{r}" spans="1:{maxcol}">'+"".join(cells[c] for c in sorted(cells,key=colidx))+'</row>')
            r+=1
    newlast=r-1
    s=s.replace('</sheetData>',"".join(rows_xml)+'</sheetData>',1)
    s=re.sub(r'<dimension ref="A1:[A-Z]+\d+"/>',f'<dimension ref="A1:{colletter(maxcol-1)}{newlast}"/>',s)
    parts[sf]=s.encode()
    if tabfile:
        tx=parts[tabfile].decode()
        m=re.search(r'ref="A1:([A-Z]+)(\d+)"',tx)
        if m: tx=tx.replace(f'ref="A1:{m.group(1)}{m.group(2)}"',f'ref="A1:{m.group(1)}{newlast}"')
        parts[tabfile]=tx.encode()
    parts['xl/workbook.xml']=re.sub(r'(<calcPr[^/]*?)(/>)',lambda mm:mm.group(1)+(' fullCalcOnLoad="1"' if 'fullCalcOnLoad' not in mm.group(1) else '')+mm.group(2),wbx,count=1).encode()
    if 'xl/calcChain.xml' in parts:
        parts.pop('xl/calcChain.xml',None); infos=[i for i in infos if i.filename!='xl/calcChain.xml']
        parts['[Content_Types].xml']=re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>','',parts['[Content_Types].xml'].decode()).encode()
        parts['xl/_rels/workbook.xml.rels']=re.sub(r'<Relationship[^>]*Target="calcChain.xml"[^>]*/>','',parts['xl/_rels/workbook.xml.rels'].decode()).encode()
    out=io.BytesIO()
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as zw:
        for i in infos: zw.writestr(i,parts[i.filename])
    return out.getvalue(), len(rows_xml)

def strcell_plain(ref,val):
    return f'<c r="{ref}" t="inlineStr"><is><t>{esc(val)}</t></is></c>'

def main():
    t=tok()
    # ---- 1. NUOVI CLIENTI creati dall'app ----
    newRaw=dl_bytes(t,NEW_PATH)
    newlines=[l for l in (newRaw.decode("utf-8").splitlines() if newRaw else []) if l.strip()]
    nuovi=[]
    for l in newlines:
        try: nuovi.append(json.loads(l))
        except: pass
    # ---- 2. MODIFICHE in coda ----
    editsRaw=dl_bytes(t,EDITS_PATH)
    lines=[l for l in (editsRaw.decode("utf-8").splitlines() if editsRaw else []) if l.strip()]
    edits=[]
    for l in lines:
        try: edits.append(json.loads(l))
        except: pass
    if not lines and not newlines:
        print("Nessuna modifica né nuovo cliente in coda."); return 0
    n=len(lines)
    agg={}
    for e in edits:
        cod=str(e.get("codice","")).strip()
        if not cod: continue
        agg.setdefault(cod,{}).update(e.get("fields",{}) or {})
    print(f"In coda: {len(nuovi)} nuovi clienti, {len(agg)} clienti modificati ({n} eventi).")

    xlsx=dl_bytes(t,DB_PATH)
    if not xlsx: print("ERRORE: database non trovato su Dropbox."); return 1

    added=0
    if nuovi:
        xlsx,added=add_clients(xlsx,nuovi)
        print(f"Nuovi clienti inseriti nel Master: {added}")
    if agg:
        xlsx=apply_to_xlsx(xlsx,agg)
    ul_bytes(t,DB_PATH,xlsx)
    print("Database aggiornato su Dropbox.")

    # archivio + pulizia code
    if lines:
        prevArch=dl_bytes(t,ARCHIVE_PATH); prevArch=prevArch.decode("utf-8") if prevArch else ""
        ul_bytes(t,ARCHIVE_PATH,(prevArch+"\n".join(lines)+"\n").encode("utf-8"))
        curRaw=dl_bytes(t,EDITS_PATH); cur=[l for l in (curRaw.decode("utf-8").splitlines() if curRaw else []) if l.strip()]
        remainder=cur[n:] if len(cur)>n else []
        ul_bytes(t,EDITS_PATH,("\n".join(remainder)+("\n" if remainder else "")).encode("utf-8"))
        print(f"Coda modifiche ripulita ({len(remainder)} in attesa).")
    if newlines and added:
        prevN=dl_bytes(t,NEW_ARCHIVE); prevN=prevN.decode("utf-8") if prevN else ""
        ul_bytes(t,NEW_ARCHIVE,(prevN+"\n".join(newlines)+"\n").encode("utf-8"))
        curN=dl_bytes(t,NEW_PATH); curl=[l for l in (curN.decode("utf-8").splitlines() if curN else []) if l.strip()]
        rem=curl[len(newlines):] if len(curl)>len(newlines) else []
        ul_bytes(t,NEW_PATH,("\n".join(rem)+("\n" if rem else "")).encode("utf-8"))
        print(f"Coda nuovi clienti ripulita ({len(rem)} in attesa).")
    return 0

if __name__=="__main__":
    import sys
    try: sys.exit(main())
    except requests.HTTPError as e:
        print("Errore HTTP Dropbox:", e.response.status_code, e.response.text[:400]); sys.exit(1)
