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
OPERATIVA_PATH = "/operativa.json"   # dati esposizione/magazzino consolidati da ingest.py

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

def download_opt(token, path):
    """Download che ritorna None se il file non esiste (no eccezione)."""
    r = requests.post(DOWNLOAD_URL, headers={
        "Authorization": f"Bearer {token}",
        "Dropbox-API-Arg": json.dumps({"path": path}),
    }, timeout=60)
    return r.content if r.status_code == 200 else None

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

    # ---- Vendite: KPI YTD (pari periodo) + serie mensile + per-cliente ----
    import collections as _c
    vend = {}
    mensili = _c.OrderedDict((f"{m:02d}", {"fatturato":0.0,"tonnellate":0.0}) for m in range(1,13))
    ytd = {2025:{"fatt":0.0,"ton":0.0,"rows":0,"cli":set()}, 2026:{"fatt":0.0,"ton":0.0,"rows":0,"cli":set()}}
    ref_doy = datetime.date.today().timetuple().tm_yday
    grid = {}   # cruscotto in testa al foglio Vendite (righe 1-28)
    try:
        V = wb["Vendite"]; VX = {}
        for r, row in enumerate(V.iter_rows(min_row=1, values_only=True), 1):
            if r <= 28:
                for c, val in enumerate(row[:14], 1):
                    if val is not None: grid[(r,c)] = val
            elif r == 31:
                VX = {str(x).strip(): i for i, x in enumerate(row) if x}
            elif r >= 32 and VX:
                cod = row[VX["Cod.Cli"]] if "Cod.Cli" in VX else None
                if not cod: continue
                cod=str(cod).strip()
                dt = row[VX.get("Data")] if "Data" in VX else None
                if not isinstance(dt,(datetime.datetime,datetime.date)): continue
                imp = row[VX.get("Importo")] or 0
                q   = (row[VX.get("Quantità (kg)")] or 0)/1000
                if dt.year==2026:
                    d = vend.setdefault(cod, {"fatt":0.0,"ton":0.0}); d["fatt"]+=imp; d["ton"]+=q
                    mensili[f"{dt.month:02d}"]["fatturato"]+=imp; mensili[f"{dt.month:02d}"]["tonnellate"]+=q
                if dt.year in ytd and dt.timetuple().tm_yday<=ref_doy:
                    y=ytd[dt.year]; y["fatt"]+=imp; y["ton"]+=q; y["rows"]+=1; y["cli"].add(cod)
    except Exception:
        pass
    def _et(y): return round(y["fatt"]/y["ton"]) if y["ton"] else 0
    totali = {
        "fatturato": round(ytd[2026]["fatt"]), "tonnellate": round(ytd[2026]["ton"],1),
        "ordini": ytd[2026]["rows"], "clienti_attivi": len(ytd[2026]["cli"]), "et_medio": _et(ytd[2026]),
        "fatturato_2025": round(ytd[2025]["fatt"]), "tonnellate_2025": round(ytd[2025]["ton"],1), "et_2025": _et(ytd[2025]),
    }
    # ---- Vista operativa: esposizione finanziaria + magazzino (cruscotto foglio Vendite) ----
    def g(r,c): return grid.get((r,c))
    def gnum(r,c):
        v=g(r,c)
        try: return round(float(v),2)
        except: return None
    operativa = {}
    try:
        operativa["esposizione"] = {
            "clienti": gnum(8,1), "merce_a_terra": gnum(8,6), "totale": gnum(8,11),
            "voci": [{"voce":g(r,9), "eur":gnum(r,12), "usd":gnum(r,14)} for r in range(20,24) if g(r,9)],
        }
        operativa["magazzino"] = {
            "disponibile_ton": gnum(15,1), "uscite_ytd_ton": gnum(15,6), "mesi_copertura": gnum(15,11),
        }
        lotti=[]; r=20
        while g(r,1) and str(g(r,1)).strip().upper()!="TOTALE" and r<30:
            lotti.append({"qualita":g(r,1),"lotto":g(r,2),"ton":gnum(r,3),"prezzo_usd":gnum(r,4),"prezzo_eur":gnum(r,5),"note":g(r,6)})
            r+=1
        operativa["lotti"]=lotti
        mk=["01","02","03","04","05","06","07","08","09","10","11","12"]
        operativa["uscite_mensili"]=[{"mese":mk[i],"tonnellate":gnum(28,i+1) or 0} for i in range(12)]
    except Exception:
        operativa={}

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
        va = vend.get(str(cod).strip())
        rec["fatturato_2026"]  = round(va["fatt"]) if va else 0
        rec["tonnellate_2026"] = round(va["ton"],1) if va else 0
        clients.append(rec)
    wb.close()
    return {
        "generato": datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
        "n_clienti": len(clients),
        "clienti": clients,
        "vendite": {
            "mensili": [{"mese":k, "fatturato":round(v["fatturato"]), "tonnellate":round(v["tonnellate"],1)} for k,v in mensili.items()],
            "totali": totali,
        },
        "operativa": operativa,
    }

# Il campo "Produttore" mescola paesi e impianti: qui si traduce nell'origine che
# conta per il CBAM. Le sigle composite e i produttori incerti restano non attribuiti:
# meglio dichiarare un volume "da attribuire" che stimare un obbligo sbagliato.
ORIGINE_CBAM = {
    "BRASILE":"BRASILE", "BRAZIL":"BRASILE",
    "KOSAYA":"RUSSIA", "URAL STEEL":"RUSSIA", "URAL":"RUSSIA", "NLMK":"RUSSIA",
    "RUSSIA":"RUSSIA", "TULACHERMET":"RUSSIA",
    "ZAPORIZHSTAL":"UCRAINA", "ZAPO":"UCRAINA", "METINVEST":"UCRAINA", "UCRAINA":"UCRAINA",
}

def origine_cbam(produttore):
    p = str(produttore or "").strip().upper()
    if not p: return None
    if p in ORIGINE_CBAM: return ORIGINE_CBAM[p]
    if "/" in p: return None          # miscela di origini: non attribuibile
    for k,v in ORIGINE_CBAM.items():
        if k in p: return v
    return None

def costo_cbam(origine, anno, cfg, prezzo=None):
    """(valore predefinito x maggiorazione - benchmark x fattore) x prezzo certificato.
    Benchmark colonna B: è quello da usare quando si dichiara con i valori predefiniti."""
    o = (cfg.get("origini") or {}).get(origine)
    if not o: return 0.0
    y = str(anno)
    magg = (cfg.get("maggiorazione") or {}).get(y)
    fatt = (cfg.get("fattore") or {}).get(y)
    bm = (cfg.get("benchmark") or {}).get("B")
    if magg is None or fatt is None or bm is None: return 0.0
    netto = max(o.get("valore_predefinito",0)*(1+magg) - bm*fatt, 0)
    return netto * (prezzo if prezzo is not None else cfg.get("prezzo_certificato", 0))

CBAM_FALLBACK = {
    "prezzo_certificato": 75.28, "benchmark": {"A":1.089,"B":1.210},
    "maggiorazione": {"2026":0.10,"2027":0.20,"2028":0.30,"2029":0.30,"2030":0.30,
                      "2031":0.30,"2032":0.30,"2033":0.30,"2034":0.30},
    "fattore": {"2026":0.975,"2027":0.95,"2028":0.90,"2029":0.775,"2030":0.515,
                "2031":0.39,"2032":0.265,"2033":0.14,"2034":0},
    "origini": {"BRASILE":{"valore_predefinito":1.478},"RUSSIA":{"valore_predefinito":3.040},
                "UCRAINA":{"valore_predefinito":2.173},"UE":{"valore_predefinito":0}},
}

def prezzo_trimestre(cfg, anno, mese):
    """Prezzo del certificato del trimestre di riferimento.

    Nel 2026 il prezzo è pubblicato per trimestre e cambia: usare quello del periodo
    dà l'esposizione corretta invece di applicare un prezzo unico a tutto l'anno."""
    q = (mese - 1)//3 + 1
    for p in (cfg.get("prezzo_certificato_storico") or []):
        if str(p.get("periodo","")).upper() == f"{anno}-T{q}":
            return p.get("valore"), f"{anno}-T{q}"
    return cfg.get("prezzo_certificato"), "corrente"


def blocco_cbam(xlsx_bytes, cfg, anno):
    """Esposizione CBAM sulle merci vendute nell'anno.

    Conta solo ciò che è soggetto a CBAM: il regime definitivo parte dal 1° gennaio 2026,
    quindi le vendite di materiale importato prima non generano obbligo. Il prezzo del
    certificato è quello del trimestre della vendita, non una media annuale."""
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
    ws = wb["Vendite"]
    hdr=None; col={}
    per = {}; non_attr = {"tonnellate":0.0,"righe":0}
    for ri,row in enumerate(ws.iter_rows(values_only=True),1):
        vals=[str(c).strip() if c is not None else "" for c in row]
        if hdr is None:
            if "Data" in vals and "Cod.Cli" in vals and "N. documento" in vals:
                hdr=ri
                for i,v in enumerate(vals,1):
                    if v: col[v]=i
            continue
        g=lambda n: row[col[n]-1] if n in col and col[n]-1<len(row) else None
        if str(g("Stato") or "").strip()!="Fattura": continue
        d=g("Data")
        if isinstance(d, datetime.datetime): d=d.date()
        if not isinstance(d, datetime.date) or d.year!=anno: continue
        t=(g("Quantità (kg)") or 0)/1000.0
        if t<=0: continue
        orig=origine_cbam(g("Produttore"))
        if not orig:
            non_attr["tonnellate"]+=t; non_attr["righe"]+=1; continue
        if orig=="UE":
            continue
        prezzo,periodo = prezzo_trimestre(cfg, anno, d.month)
        eur_t = costo_cbam(orig, anno, cfg, prezzo)
        r=per.setdefault(orig,{"tonnellate":0.0,"costo":0.0,"trimestri":{}})
        r["tonnellate"]+=t; r["costo"]+=t*eur_t
        q=f"T{(d.month-1)//3+1}"
        tr=r["trimestri"].setdefault(q,{"tonnellate":0.0,"eur_t":round(eur_t,2),
                                        "prezzo":prezzo,"periodo":periodo,"costo":0.0})
        tr["tonnellate"]+=t; tr["costo"]+=t*eur_t
    wb.close()
    for orig,r in per.items():
        r["eur_t"]=round(r["costo"]/r["tonnellate"],2) if r["tonnellate"] else 0
        r["tonnellate"]=round(r["tonnellate"],1); r["costo"]=round(r["costo"])
        for q,tr in r["trimestri"].items():
            tr["tonnellate"]=round(tr["tonnellate"],1); tr["costo"]=round(tr["costo"])
    non_attr["tonnellate"]=round(non_attr["tonnellate"],1)
    return {"anno":anno, "per_origine":per, "non_attribuito":non_attr,
            "totale_costo":round(sum(r["costo"] for r in per.values())),
            "totale_tonnellate":round(sum(r["tonnellate"] for r in per.values()),1),
            "prezzo_certificato":cfg.get("prezzo_certificato"),
            "prezzi_trimestre":{p.get("periodo"):p.get("valore") for p in (cfg.get("prezzo_certificato_storico") or [])},
            "nota":"Solo vendite dell'anno con origine soggetta a CBAM. Il regime definitivo decorre dal 1° gennaio 2026.",
            "aggiornato":cfg.get("aggiornato")}

def blocco_commerciale(xlsx_bytes, offerte, documenti, anno):
    """Rapporto fra offerte presentate, contratti chiusi e vendite fatturate."""
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
    ws = wb["Vendite"]
    hdr=None; col={}
    fatt={"n":0,"tonnellate":0.0,"valore":0.0}
    contratti=set(); contratti_t=0.0; contratti_v=0.0
    mesi_off={}; mesi_con={}
    for ri,row in enumerate(ws.iter_rows(values_only=True),1):
        vals=[str(c).strip() if c is not None else "" for c in row]
        if hdr is None:
            if "Data" in vals and "Cod.Cli" in vals and "N. documento" in vals:
                hdr=ri
                for i,v in enumerate(vals,1):
                    if v: col[v]=i
            continue
        g=lambda n: row[col[n]-1] if n in col and col[n]-1<len(row) else None
        d=g("Data")
        if isinstance(d, datetime.datetime): d=d.date()
        if not isinstance(d, datetime.date) or d.year!=anno: continue
        st=str(g("Stato") or "").strip(); t=(g("Quantità (kg)") or 0)/1000.0; v=g("Importo") or 0
        doc=str(g("N. documento") or "").strip()
        if st=="Fattura":
            fatt["n"]+=1; fatt["tonnellate"]+=t; fatt["valore"]+=v
        if doc and st in ("Fattura","Ordine","Evaso") and doc not in contratti:
            contratti.add(doc); mesi_con[f"{d.month:02d}"]=mesi_con.get(f"{d.month:02d}",0)+1
    wb.close()
    # offerte: quelle create nell'app più quelle archiviate come documento
    n_off=0; val_off=0.0; ton_off=0.0
    for o in (offerte or {}).values():
        data=str(o.get("data") or "")
        if data[:4]!=str(anno): continue
        n_off+=1; mesi_off[data[5:7]]=mesi_off.get(data[5:7],0)+1
        q=float(o.get("quantita") or 0); p=float(o.get("prezzo_vendita") or 0)
        ton_off+=q; val_off+=q*p
    n_doc_off=0
    for dd in (documenti or {}).values():
        if dd.get("tipo")!="offerta": continue
        if str(dd.get("data") or "")[:4]!=str(anno): continue
        n_doc_off+=1; m=str(dd.get("data"))[5:7]; mesi_off[m]=mesi_off.get(m,0)+1
    n_doc_con=sum(1 for dd in (documenti or {}).values()
                  if dd.get("tipo")=="contratto" and str(dd.get("data") or "")[:4]==str(anno))
    tot_off=n_off+n_doc_off
    return {"anno":anno,
            "offerte":{"app":n_off,"archivio":n_doc_off,"totale":tot_off,
                       "tonnellate":round(ton_off,1),"valore":round(val_off)},
            "contratti":{"da_vendite":len(contratti),"documenti":n_doc_con},
            "vendite":{"righe":fatt["n"],"tonnellate":round(fatt["tonnellate"],1),"valore":round(fatt["valore"])},
            "conversione": round(100*len(contratti)/tot_off,1) if tot_off else None,
            "mensili":{"offerte":mesi_off,"contratti":mesi_con}}

# Quanto resta "caldo" un cliente dopo un'offerta, e quando si considera raffreddato.
GIORNI_CALDO = 14
GIORNI_RAFFREDDAMENTO = 45

def stato_commerciale(xlsx_bytes, offerte, documenti, oggi=None):
    """Per ogni cliente: contratti aperti, ultima offerta e stato che ne consegue.

    La regola commerciale è semplice e la si applica una volta sola, qui:
    - chi ha un contratto aperto è servito, non va ricontattato;
    - chi ha ricevuto un'offerta da meno di due settimane è caldo;
    - poi entra in raffreddamento, e oltre le sei settimane torna da lavorare.
    """
    oggi = oggi or datetime.date.today()
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
    ws = wb["Vendite"]
    hdr=None; col={}
    contratti={}      # codice -> {aperti, tonnellate, valore, elenco}
    ultima_vendita={}
    for ri,row in enumerate(ws.iter_rows(values_only=True),1):
        vals=[str(c).strip() if c is not None else "" for c in row]
        if hdr is None:
            if "Data" in vals and "Cod.Cli" in vals and "N. documento" in vals:
                hdr=ri
                for i,v in enumerate(vals,1):
                    if v: col[v]=i
            continue
        g=lambda n: row[col[n]-1] if n in col and col[n]-1<len(row) else None
        cod=str(g("Cod.Cli") or "").strip()
        if not cod: continue
        st=str(g("Stato") or "").strip()
        d=g("Data")
        if isinstance(d, datetime.datetime): d=d.date()
        if st=="Fattura" and isinstance(d, datetime.date):
            if cod not in ultima_vendita or d>ultima_vendita[cod]: ultima_vendita[cod]=d
        if st=="Ordine":
            t=(g("Quantità (kg)") or 0)/1000.0
            if t<=0: continue
            c=contratti.setdefault(cod,{"aperti":0,"tonnellate":0.0,"valore":0.0,"elenco":[]})
            c["aperti"]+=1; c["tonnellate"]+=t; c["valore"]+=(g("Importo") or 0)
            c["elenco"].append({"documento":str(g("N. documento") or ""),
                                "tonnellate":round(t,1),
                                "data":d.isoformat() if isinstance(d,datetime.date) else None})
    wb.close()

    # ultima offerta per cliente: quelle create nell'app e quelle archiviate
    ult_off={}
    def considera(cod, data, prezzo, margine, rif):
        if not cod or not data: return
        try: dd=datetime.date.fromisoformat(str(data)[:10])
        except Exception: return
        cur=ult_off.get(cod)
        if not cur or dd>cur["data"]:
            ult_off[cod]={"data":dd,"prezzo":prezzo,"margine":margine,"rif":rif}
    for o in (offerte or {}).values():
        considera(o.get("codice"), o.get("data"), o.get("prezzo_vendita"), o.get("margine"), o.get("numero"))
    for d in (documenti or {}).values():
        if d.get("tipo")=="offerta":
            considera(d.get("cliente"), d.get("data"), d.get("prezzo"), d.get("margine"), d.get("nome"))

    out={}
    codici=set(contratti)|set(ult_off)|set(ultima_vendita)
    for cod in codici:
        c=contratti.get(cod); o=ult_off.get(cod)
        gg=(oggi-o["data"]).days if o else None
        if c and c["aperti"]>0:
            stato="Contratto attivo"; motivo=f"{c['aperti']} contratti aperti per {round(c['tonnellate'],1)} t"
        elif gg is not None and gg<=GIORNI_CALDO:
            stato="Cliente caldo"; motivo=f"offerta di {gg} giorni fa"
        elif gg is not None and gg<=GIORNI_RAFFREDDAMENTO:
            stato="Cliente raffreddamento"; motivo=f"offerta di {gg} giorni fa, nessuna chiusura"
        elif gg is not None:
            stato="Da riprendere"; motivo=f"ultima offerta {gg} giorni fa"
        else:
            stato=None; motivo=None
        out[cod]={
            "stato_calcolato":stato,"motivo":motivo,
            "contratti_aperti":(c["aperti"] if c else 0),
            "contratti_tonnellate":round(c["tonnellate"],1) if c else 0,
            "contratti_valore":round(c["valore"]) if c else 0,
            "contratti_elenco":(c["elenco"][:6] if c else []),
            "ultima_offerta":(o["data"].isoformat() if o else None),
            "ultima_offerta_prezzo":(o["prezzo"] if o else None),
            "ultima_offerta_margine":(o["margine"] if o else None),
            "ultima_offerta_rif":(o["rif"] if o else None),
            "giorni_da_offerta":gg,
            "ultima_vendita":(ultima_vendita[cod].isoformat() if cod in ultima_vendita else None),
        }
    return out

def main():
    token = get_access_token()
    print("Token ottenuto. Scarico il database...")
    xlsx = download(token, DB_PATH)
    print(f"Database scaricato ({len(xlsx)} byte). Costruisco lo snapshot...")
    snap = build(xlsx)
    # Operativa CENTRALIZZATA: /operativa.json (prodotto da ingest.py dai file esposizione/magazzino).
    opraw = download_opt(token, OPERATIVA_PATH)
    try:
        snap["operativa"] = json.loads(opraw) if opraw else {}
    except Exception:
        snap["operativa"] = {}
    op = snap["operativa"]
    print(f"OPERATIVA: righe_esposizione={len((op.get('esposizione') or {}).get('righe',[]))} "
          f"giacenze={len(op.get('giacenze',[]))} "
          f"uscite_valorizzate={sum(1 for m in op.get('uscite_mensili',[]) if m.get('tonnellate'))} "
          f"aggiornato={op.get('aggiornato')}")
    if not op.get('esposizione') and not op.get('giacenze'):
        print("NOTA: operativa.json assente o vuoto. Carica i file in /caricamenti; l'ingest gira alle 10:00 e 18:00.")
    # --- CBAM e quadro commerciale ---
    anno = datetime.date.today().year
    def leggi(path, vuoto):
        raw = download_opt(token, path)
        try: return json.loads(raw) if raw else vuoto
        except Exception: return vuoto
    cfg = leggi("/cbam-config.json", {}) or {}
    for k,v in CBAM_FALLBACK.items(): cfg.setdefault(k,v)
    try:
        snap["cbam"] = blocco_cbam(xlsx, cfg, anno)
        c = snap["cbam"]
        print(f"CBAM {anno}: {c['totale_tonnellate']:,.0f} t attribuite, costo stimato "
              f"{c['totale_costo']:,} EUR | non attribuite {c['non_attribuito']['tonnellate']:,.0f} t")
    except Exception as e:
        print("CBAM non calcolato:", e); snap["cbam"] = {}
    # stato commerciale per cliente: contratti aperti e ricaduta delle offerte
    try:
        offerte_db = leggi("/offerte.json",{}); documenti_db = leggi("/documenti.json",{})
        stati = stato_commerciale(xlsx, offerte_db, documenti_db)
        for c in snap.get("clienti",[]):
            s = stati.get(str(c.get("codice") or "").strip())
            if s: c.update(s)
        from collections import Counter as _Cnt
        rip=_Cnt(v["stato_calcolato"] for v in stati.values() if v["stato_calcolato"])
        print("Stato commerciale:", dict(rip),
              "| clienti con contratto aperto:", sum(1 for v in stati.values() if v["contratti_aperti"]))
    except Exception as e:
        print("Stato commerciale non calcolato:", e)
        offerte_db = leggi("/offerte.json",{}); documenti_db = leggi("/documenti.json",{})
    try:
        snap["commerciale"] = blocco_commerciale(xlsx, offerte_db, documenti_db, anno)
        k = snap["commerciale"]
        print(f"Commerciale {anno}: offerte {k['offerte']['totale']} | contratti {k['contratti']['da_vendite']} "
              f"| conversione {k['conversione']}%")
    except Exception as e:
        print("Quadro commerciale non calcolato:", e); snap["commerciale"] = {}

    blob = json.dumps(snap, ensure_ascii=False).encode("utf-8")
    upload(token, SNAPSHOT_PATH, blob)
    print(f"Snapshot pubblicato: {snap['n_clienti']} clienti, {len(blob)} byte.")

if __name__ == "__main__":
    try:
        sys.exit(main())
    except requests.HTTPError as e:
        print("Errore HTTP Dropbox:", e.response.status_code, e.response.text[:500])
        sys.exit(1)
