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
    if u.endswith(".PDF"):
        if "DETTAGLIO" in u: return "dettaglio"                       # DETTAGLIO ... .pdf (ordini + esposizione per cliente)
        if re.match(r'[A-Z]{2,4}\d{3}', u): return "bilancio"         # CODICE - Ragione.pdf
        return "altro"
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

# ---------- DETTAGLIO (ordini aperti + esposizione per cliente) ----------
def money(s):
    s=str(s).strip().replace('.','').replace(',','.')
    try: return round(float(s),2)
    except: return None
def parse_dettaglio(pdf_bytes):
    with tempfile.NamedTemporaryFile(suffix=".pdf",delete=False) as f: f.write(pdf_bytes); path=f.name
    try: t=subprocess.run(["pdftotext","-layout",path,"-"],capture_output=True,text=True).stdout
    finally:
        try: os.unlink(path)
        except: pass
    cur=None; section=None; espos={}; ordini={}
    EPOCH=datetime.date(1899,12,30)
    for ln in t.splitlines():
        m=re.search(r'Cliente\s+([A-Z]{2,4}\d{3})\s+', ln)
        if m: cur=m.group(1)
        if "Movimenti Contabili" in ln: section="contabili"
        elif "Movimenti Contratti-Ordini" in ln: section="ordini"
        elif "Movimenti Pre-Fatture" in ln: section="pre"
        me=re.search(r'Totali senza Eff\.\s+([\d.,]+)', ln)
        if me and cur: espos[cur]=money(me.group(1))
        if section=="ordini" and cur:
            mo=re.search(r'\b(\d{2}/G\d{3,4})\b', ln)
            if mo:
                doc=mo.group(1); toks=ln.split()
                decs=[i for i,tk in enumerate(toks) if re.match(r'^[\d.]+,\d{2}$',tk)]
                imp=res=None
                if decs:
                    eur_i=decs[-2] if len(decs)>=2 else decs[-1]; imp=money(toks[eur_i])
                    # La quantità residua è il token subito prima dell'Importo EUR e PUÒ avere
                    # decimali (es. "648,52"). Cercando solo interi si prendeva per sbaglio
                    # l'Importo.Orig (es. 499000), gonfiando la quantità di 1000 volte.
                    for k in range(eur_i-1,-1,-1):
                        if re.match(r'^[\d.]+(?:,\d+)?$',toks[k]): res=money(toks[k]); break
                ds=None; md=re.search(r'(\d{2}/\d{2}/\d{4})', ln)
                if md:
                    try: ds=(datetime.datetime.strptime(md.group(1),"%d/%m/%Y").date()-EPOCH).days
                    except: ds=None
                # quantità/importo ORIGINALI del contratto: servono a ricavare il consegnato
                # (consegnato = originale - residuo), dato che i file non riportano le
                # quantità delle singole fatture.
                i_doc=toks.index(doc) if doc in toks else 0
                nums=[x for x in toks[i_doc+1:] if re.match(r'^[\d.]+(?:,\d+)?$',x)]
                q_or=money(nums[0]) if len(nums)>=1 else None
                i_or=money(nums[1]) if len(nums)>=2 else None
                ordini[doc]={"cli":cur,"qty_kg":round((res or 0)*1000),"imp":imp or 0,"data_serial":ds,
                             "q_orig_kg":(round(q_or*1000) if q_or is not None else None),
                             "imp_orig":i_or}
    return espos,ordini


def read_vendite_full(xlsx_bytes):
    """Layout del foglio Vendite + stato per contratto.

    Ritorna (hdr, col, last, stato) dove stato[documento] = {
      'fat_kg','fat_imp'   somma delle righe già registrate come Fattura
      'ord_row'            riga della quota ancora aperta (Stato = Ordine), se c'è
      'ord_kg','ord_imp'   valori attualmente scritti su quella riga
      'eredita'            Prodotto/Qualità/Produttore/Pagamento presi dalle righe esistenti
    }
    """
    wb=openpyxl.load_workbook(io.BytesIO(xlsx_bytes),read_only=True,data_only=True)
    V=wb["Vendite"]; hdr=None; col={}; stato={}; last=1
    for ri,row in enumerate(V.iter_rows(values_only=True),1):
        vals=[str(c).strip() if c is not None else "" for c in row]
        if hdr is None:
            if "Data" in vals and "Cod.Cli" in vals and "N. documento" in vals:
                hdr=ri
                for i,v in enumerate(vals,1):
                    if v: col[v]=i
            continue
        g=lambda n: row[col[n]-1] if n in col and col[n]-1<len(row) else None
        doc=g("N. documento")
        if doc in (None,""): continue
        d=str(doc).strip(); last=ri
        st=stato.setdefault(d,{"fat_kg":0.0,"fat_imp":0.0,"ord_row":None,
                               "ord_kg":None,"ord_imp":None,"eredita":{},
                               "cli":None,"ord_data":None})
        if st["cli"] is None and g("Cod.Cli") not in (None,""): st["cli"]=str(g("Cod.Cli")).strip()
        for nome in ("Prodotto","Qualità","Produttore","Pagamento"):
            v=g(nome)
            if v not in (None,"") and nome not in st["eredita"]: st["eredita"][nome]=v
        s=str(g("Stato") or "").strip()
        if s=="Fattura":
            st["fat_kg"]+=g("Quantità (kg)") or 0; st["fat_imp"]+=g("Importo") or 0
        elif s=="Ordine":
            st["ord_row"]=ri; st["ord_kg"]=g("Quantità (kg)"); st["ord_imp"]=g("Importo")
            dv=g("Data")
            if isinstance(dv,(datetime.date,datetime.datetime)):
                dv=dv.date() if isinstance(dv,datetime.datetime) else dv
                st["ord_data"]=(dv-datetime.date(1899,12,30)).days
            elif isinstance(dv,(int,float)): st["ord_data"]=int(dv)
    wb.close(); return hdr,col,last,stato

def read_vendite_layout(xlsx_bytes):
    wb=openpyxl.load_workbook(io.BytesIO(xlsx_bytes),read_only=True,data_only=True)
    V=wb["Vendite"]; hdr=None; col={}
    for ri,row in enumerate(V.iter_rows(min_row=1,max_row=40,values_only=True),1):
        vals=[str(c).strip() if c is not None else "" for c in row]
        if "Data" in vals and "Cod.Cli" in vals and "N. documento" in vals:
            hdr=ri
            for i,v in enumerate(vals,1):
                if v: col[v]=i
            break
    docs=set(); last=hdr or 1
    if hdr:
        dcol=col["N. documento"]-1; ccol=col["Cod.Cli"]-1
        for ri,row in enumerate(V.iter_rows(min_row=hdr+1,values_only=True),hdr+1):
            d=row[dcol] if dcol<len(row) else None; c=row[ccol] if ccol<len(row) else None
            if d not in (None,""): docs.add(str(d).strip())
            if c not in (None,""): last=ri
    wb.close(); return hdr,col,docs,last

SOGLIA_T = 0.5      # tonnellate: sotto questa differenza non si registra nulla
SOGLIA_EUR = 100.0  # euro


def pianifica_vendite(contratti, stato, clienti_noti, clienti_nel_report=None):
    """Decide cosa scrivere nel foglio Vendite, senza toccare il file.

    - consegne: righe Fattura nuove per la quota consegnata e non ancora registrata
                (consegnato = quantità originale del contratto - residuo attuale)
    - residui : aggiornamento della riga Ordine quando il residuo è cambiato
    - nuovi   : contratti mai visti, da inserire come Ordine aperto
    - evasi   : righe Ordine il cui contratto non è più nel DETTAGLIO -> Stato "Evaso"
    """
    consegne, residui, nuovi, evasi = [], [], {}, []
    for doc, o in contratti.items():
        if clienti_noti and o.get("cli") not in clienti_noti:
            continue
        st = stato.get(doc)
        if st is None:
            nuovi[doc] = o
            continue
        if o.get("q_orig_kg") is not None and o.get("imp_orig") is not None:
            cons_kg = o["q_orig_kg"] - o["qty_kg"]
            cons_eur = o["imp_orig"] - o["imp"]
            d_kg = cons_kg - st["fat_kg"]
            d_eur = cons_eur - st["fat_imp"]
            # solo consegne in più: se a database risulta di più, il contratto è stato
            # ampliato o rifornito oltre l'originale -> non si tocca nulla
            if d_kg > SOGLIA_T * 1000 or d_eur > SOGLIA_EUR:
                kg = round(max(d_kg, 0)); eur = round(max(d_eur, 0.0), 2)
                # coerenza: quantità e valore devono muoversi insieme. Se una delle due è
                # a zero mentre l'altra è rilevante, il dato del PDF è incoerente
                # (capita su contratti senza importo residuo): meglio segnalare che scrivere.
                if (kg <= 0 and eur > SOGLIA_EUR) or (eur <= 0 and kg > SOGLIA_T * 1000):
                    print(f"   ! {doc} ({o['cli']}): consegna incoerente "
                          f"({kg/1000:.2f} t / {eur:,.2f} EUR) - ignorata, da verificare a mano.")
                else:
                    consegne.append({"doc": doc, "cli": o["cli"], "kg": kg, "imp": eur,
                                     "data_serial": o.get("data_serial"),
                                     "eredita": st["eredita"]})
        if st["ord_row"]:
            dk = abs((st["ord_kg"] or 0) - o["qty_kg"])
            de = abs((st["ord_imp"] or 0) - o["imp"])
            if dk > SOGLIA_T * 1000 or de > SOGLIA_EUR:
                residui.append({"row": st["ord_row"], "kg": o["qty_kg"], "imp": o["imp"], "doc": doc})
    # Chiusure: un ordine si considera evaso SOLO se il suo cliente compare nel DETTAGLIO
    # (quindi il report lo copre) ma il contratto non è più tra quelli aperti. Se il cliente
    # manca del tutto, il report è parziale su di lui e non si conclude nulla.
    # Il residuo che sparisce è merce consegnata: va registrato, non cancellato.
    for doc, st in stato.items():
        if not st["ord_row"] or doc in contratti: continue
        if (st["ord_kg"] or 0) <= 0 and (st["ord_imp"] or 0) <= 0: continue
        cli = st.get("cli")
        if clienti_nel_report is not None and cli not in clienti_nel_report: continue
        evasi.append({"row": st["ord_row"], "doc": doc})
        if (st["ord_kg"] or 0) > 0 or (st["ord_imp"] or 0) > 0:
            consegne.append({"doc": doc, "cli": cli, "kg": round(st["ord_kg"] or 0),
                             "imp": round(st["ord_imp"] or 0.0, 2),
                             "data_serial": st.get("ord_data"), "eredita": st["eredita"],
                             "finale": True})
    return consegne, residui, nuovi, evasi


def _set_cell(s, letter, row, value, isnum):
    """Sostituisce il valore di una cella esistente conservandone lo stile."""
    ref = f"{letter}{row}"
    m = re.search(r'<c r="%s"([^>]*?)(?:/>|>(.*?)</c>)' % ref, s, re.DOTALL)
    if not m:
        return s, False
    attrs = m.group(1)
    attrs = re.sub(r'\s+t="[^"]*"', '', attrs)      # via il tipo precedente
    if isnum:
        nuovo = f'<c r="{ref}"{attrs}><v>{value}</v></c>'
    else:
        v = str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        nuovo = f'<c r="{ref}"{attrs} t="inlineStr"><is><t>{v}</t></is></c>'
    return s[:m.start()] + nuovo + s[m.end():], True


def add_orders(xlsx_bytes,new_orders,hdr,col,last,consegne=None,residui=None,evasi=None):
    """Scrive nel foglio Vendite: nuovi ordini aperti, consegne (righe Fattura),
    aggiornamento dei residui e chiusura degli ordini evasi. Sempre in chirurgia
    XML: openpyxl in scrittura distruggerebbe formattazione condizionale e formule."""
    z=zipfile.ZipFile(io.BytesIO(xlsx_bytes)); parts={n:z.read(n) for n in z.namelist()}; infos=z.infolist(); z.close()
    wbx=parts['xl/workbook.xml'].decode()
    rid=re.search(r'<sheet name="Vendite"[^>]*r:id="(rId\d+)"',wbx).group(1)
    tgt=re.search(r'Id="%s"[^>]*Target="([^"]*)"'%rid,parts['xl/_rels/workbook.xml.rels'].decode()).group(1)
    sf="xl/"+tgt.replace("\\","/"); s=parts[sf].decode()
    tabfile=next((n for n in parts if n.startswith("xl/tables/") and b"tbl_Vendite" in parts[n]),None)
    # ATTENZIONE: col[] è 1-based (enumerate(...,1) in read_vendite_layout) mentre
    # colletter() è 0-based -> serve i-1. Senza il -1 ogni valore finiva una colonna
    # più a destra (data in "Cod.Cli", ecc.) e il controllo anti-duplicato, che legge
    # la colonna giusta, non trovava mai i documenti: ordini reinseriti a ogni giro.
    tx=parts[tabfile].decode(); L={n:colletter(i-1) for n,i in col.items()}; maxcol=max(col.values())
    style_of={}
    for name,idx in col.items():
        letter=colletter(idx-1); m=re.search(r'<c r="%s%d"([^>]*?)(?:/>|>)'%(letter,last),s)
        sm=re.search(r's="(\d+)"',m.group(1)) if m else None; style_of[name]=sm.group(1) if sm else None
    FORMULAS={}
    for chunk in tx.split('<tableColumn')[1:]:
        nm=re.search(r'name="([^"]+)"',chunk); cf=re.search(r'<calculatedColumnFormula>(.*?)</calculatedColumnFormula>',chunk,re.DOTALL)
        if nm and cf: FORMULAS[nm.group(1)]=cf.group(1)
    if "Anno" in col: FORMULAS["Anno"]='IF(A{R}=&quot;&quot;,&quot;&quot;,YEAR(A{R}))'
    def esc(v): return str(v).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    def cell(name,r,value=None,isnum=False):
        letter=L[name]; st=style_of.get(name); sa=f' s="{st}"' if st else ''
        if name in FORMULAS: return f'<c r="{letter}{r}"{sa}><f>{FORMULAS[name].replace("{R}",str(r))}</f></c>'
        if value in (None,""): return f'<c r="{letter}{r}"{sa}/>'
        if isnum: return f'<c r="{letter}{r}"{sa}><v>{value}</v></c>'
        return f'<c r="{letter}{r}"{sa} t="inlineStr"><is><t>{esc(value)}</t></is></c>'
    # --- 1) aggiornamento delle righe già presenti (residui e ordini evasi) ---
    n_res=n_eva=0
    for u in (residui or []):
        for nome,val,isnum in (("Quantità (kg)",u["kg"],True),("Importo",u["imp"],True)):
            if nome in L:
                s,done=_set_cell(s,L[nome],u["row"],val,isnum); n_res+=1 if done else 0
    for u in (evasi or []):
        if "Stato" in L:
            s,done=_set_cell(s,L["Stato"],u["row"],"Evaso",False); n_eva+=1 if done else 0
        # un ordine evaso non ha più residuo
        for nome in ("Quantità (kg)","Importo"):
            if nome in L: s,_=_set_cell(s,L[nome],u["row"],0,True)

    # --- 2) righe nuove (ordini aperti e consegne) ---
    rows_xml=[]; r=last+1
    def riga(o,doc,stato_val,eredita=None):
        cells=[]
        for name,idx in sorted(col.items(),key=lambda kv:kv[1]):
            if name=="Data": cells.append(cell(name,r,o.get("data_serial"),True) if o.get("data_serial") else cell(name,r))
            elif name=="Cod.Cli": cells.append(cell(name,r,o["cli"]))
            elif name=="N. documento": cells.append(cell(name,r,doc))
            elif name=="Quantità (kg)": cells.append(cell(name,r,o.get("qty_kg",o.get("kg")),True))
            elif name=="Importo": cells.append(cell(name,r,o["imp"],True))
            elif name=="Valuta": cells.append(cell(name,r,"EUR"))
            elif name=="Stato": cells.append(cell(name,r,stato_val))
            elif eredita and name in eredita: cells.append(cell(name,r,eredita[name]))
            else: cells.append(cell(name,r))
        return f'<row r="{r}" spans="1:{maxcol}">'+"".join(cells)+'</row>'
    for doc in sorted(new_orders):
        rows_xml.append(riga(new_orders[doc],doc,"Ordine")); r+=1
    for cg in sorted(consegne or [], key=lambda x:x["doc"]):
        rows_xml.append(riga(cg,cg["doc"],"Fattura",cg.get("eredita"))); r+=1
    newlast=r-1
    if residui or evasi:
        print(f"   Vendite: {n_res//2} residui aggiornati, {n_eva} ordini chiusi come Evasi.")
    s=s.replace('</sheetData>',"".join(rows_xml)+'</sheetData>',1)
    s=re.sub(r'<dimension ref="A1:[A-Z]+\d+"/>',f'<dimension ref="A1:{colletter(maxcol-1)}{newlast}"/>',s)
    parts[sf]=s.encode()
    oref=re.search(r'ref="A%d:([A-Z]+)%d"'%(hdr,last),tx)
    if oref: tx=tx.replace('ref="A%d:%s%d"'%(hdr,oref.group(1),last),'ref="A%d:%s%d"'%(hdr,oref.group(1),newlast))
    parts[tabfile]=tx.encode()
    parts['xl/workbook.xml']=re.sub(r'(<calcPr[^/]*?)(/>)',lambda mm:mm.group(1)+(' fullCalcOnLoad="1"' if 'fullCalcOnLoad' not in mm.group(1) else '')+mm.group(2),wbx,count=1).encode()
    if 'xl/calcChain.xml' in parts:
        parts.pop('xl/calcChain.xml',None); infos=[i for i in infos if i.filename!='xl/calcChain.xml']
        parts['[Content_Types].xml']=re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>','',parts['[Content_Types].xml'].decode()).encode()
        parts['xl/_rels/workbook.xml.rels']=re.sub(r'<Relationship[^>]*Target="calcChain.xml"[^>]*/>','',parts['xl/_rels/workbook.xml.rels'].decode()).encode()
    out=io.BytesIO()
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as zw:
        for i in infos: zw.writestr(i,parts[i.filename])
    return out.getvalue(),newlast

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
    # DATABASE: bilanci (credito) + DETTAGLIO (esposizione per cliente -> Master, ordini aperti -> Vendite)
    bil=[e for e in files if classify(e["name"])=="bilancio"]
    det=[e for e in files if classify(e["name"])=="dettaglio"]
    if bil or det:
        xlsx=dl(t,DBX)
        if not xlsx:
            print("   ATTENZIONE: crm-database.xlsx non trovato: bilanci/dettaglio non applicati.")
        else:
            HX,rows=master_index(xlsx); mupd={}
            def addm(cod,colname,val,isnum):
                if cod in rows and colname in HX and val not in (None,""):
                    mupd.setdefault(rows[cod],[]).append((colletter(HX[colname]),val,isnum))
            for e in bil:
                cod=code_from_name(e["name"])
                if not cod: print(f"   bilancio ignorato: {e['name']}"); continue
                rep=parse_report(dl(t,e["path_lower"]))
                for col,key,isnum in CREDIT: addm(cod,col,rep.get(key),isnum)
                print(f"   bilancio {cod} letto")
            new_orders={}; clienti_report=set()
            for e in det:
                espos,ordini=parse_dettaglio(dl(t,e["path_lower"]))
                for cod,v in espos.items(): addm(cod,"Esposizione corrente (€)",v,True)
                new_orders.update(ordini)
                clienti_report |= set(espos.keys()) | {o["cli"] for o in ordini.values() if o.get("cli")}
                print(f"   dettaglio: {len(espos)} esposizioni, {len(ordini)} ordini letti")
            if mupd:
                xlsx=apply_master_credit(xlsx,mupd); print(f"   Master aggiornato: {len(mupd)} clienti (credito/esposizione).")
            if new_orders:
                hdr,col,last,stato=read_vendite_full(xlsx)
                consegne,residui,nuovi,evasi=pianifica_vendite(new_orders,stato,set(rows.keys()),clienti_report)
                # sicurezza: se il DETTAGLIO è parziale, non chiudere mezzo portafoglio
                if len(evasi) > max(5, len(new_orders)//2):
                    print(f"   ATTENZIONE: {len(evasi)} ordini risulterebbero evasi: DETTAGLIO forse incompleto, chiusure ignorate.")
                    chiusi={e["doc"] for e in evasi}; evasi=[]
                    consegne=[c for c in consegne if not (c.get("finale") and c["doc"] in chiusi)]
                if consegne or residui or nuovi or evasi:
                    xlsx,_=add_orders(xlsx,nuovi,hdr,col,last,consegne=consegne,residui=residui,evasi=evasi)
                    tot_t=sum(c["kg"] for c in consegne)/1000; tot_e=sum(c["imp"] for c in consegne)
                    print(f"   Vendite: +{len(nuovi)} ordini nuovi, +{len(consegne)} consegne "
                          f"({tot_t:,.1f} t / {tot_e:,.2f} EUR).")
                else:
                    print("   Vendite: nessuna variazione da registrare.")
            ul(t,DBX,xlsx); print("crm-database.xlsx aggiornato.")
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
