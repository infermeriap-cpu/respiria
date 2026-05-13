#!/usr/bin/env python3
"""
RespirIA — Actualitzador automàtic CIMA → Excel → HTML
=======================================================
Autora: Sílvia Álvarez Vega · ICS Atenció Primària Girona
Versió: 5.1 · 2026-05-13

Modes d'execució:
  --detecta          Consulta CIMA, afegeix novetats al Excel, escriu novetats_count.txt
  --comprova-pendents Compta files amb col N no buida, escriu pendents_count.txt
  --comprova-publicar Compta files validades (col N buida), escriu publicar_count.txt
  --regenera         Regenera HTML amb NOMÉS els fàrmacs validats (col N buida)
  --tot              detecta + regenera (ús local)
"""

import requests, openpyxl, json, time, os, shutil, sys
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓ
# ═══════════════════════════════════════════════════════════════════════════════
EXCEL_PATH  = "inhaladors_MPOC_ASMA_Avanzado_FINAL.xlsx"
HTML_PATH   = "RespirIA_v2.html"
OUTPUT_HTML = "RespirIA_v2.html"
CIMA_BASE   = "https://cima.aemps.es/cima/rest"

I_CLASSE = 0;  I_IA = 1;  I_NOM = 2;   I_CN = 3
I_DISP   = 4;  I_DOSI = 5; I_TIPUS = 6; I_CO2 = 7
I_FLUX   = 8;  I_FLUX4 = 9; I_LINK = 10
I_DATA   = 11; I_ORIGEN = 12; I_ESTAT = 13
I_PHF    = 14; I_MATMA = 15

COLOR_NOU     = "FEF3C7"
COLOR_ATENCIO = "FEE2E2"

# ═══════════════════════════════════════════════════════════════════════════════
# DOSIS PHF CatSalut 2018 + GEMA 5.5
# ═══════════════════════════════════════════════════════════════════════════════
DOSIS_MPOC = {
    "salbutamol":               ("100-200 µg si cal",    "200 µg/6h",       True,  False),
    "terbutalina":              ("500 µg si cal",         "6.000 µg/24h",    False, False),
    "ipratropi":                ("40 µg si cal",          "240 µg/24h",      True,  False),
    "formoterol":               ("12 µg/12h",             "24 µg/12h",       True,  False),
    "salmeterol":               ("50 µg/12h",             "100 µg/12h",      True,  False),
    "indacaterol":              ("150 µg/24h",            "300 µg/24h",      True,  False),
    "olodaterol":               ("5 µg/24h",              "5 µg/24h",        False, False),
    "tiotropi":                 ("18 µg/24h (Handihaler) / 5 µg/24h (Respimat) / 10 µg/24h (Zonda)",
                                 "= pauta",               True,  False),
    "glicopirroni":             ("44 µg/24h",             "44 µg/24h",       False, False),
    "umeclidini":               ("55 µg/24h",             "55 µg/24h",       False, False),
    "aclidini":                 ("322 µg/12h",            "322 µg/12h",      False, False),
    "indacaterol/glicopirroni": ("85/43 µg/24h",          "85/43 µg/24h",    True,  False),
    "tiotropi/olodaterol":      ("5/5 µg/24h",            "5/5 µg/24h",      False, False),
    "umeclidini/vilanterol":    ("55/22 µg/24h",          "55/22 µg/24h",    False, False),
    "aclidini/formoterol":      ("340/12 µg/12h",         "340/12 µg/12h",   False, False),
    "salmeterol/fluticasona":   ("50/500 µg/12h",         "50/500 µg/12h",   False, False),
    "formoterol/budesonida":    ("9/320 µg/12h",          "36/1.280 µg/24h", False, False),
    "formoterol/beclometasona": ("12/200 µg/12h",         "12/400 µg/12h",   False, False),
    "vilanterol/fluticasona":   ("22/92 µg/24h",          "22/184 µg/24h",   False, False),
    "fluticasona":              ("250-500 µg/12h",         "500 µg/12h",      False, False),
    "budesonida":               ("200-400 µg/12h",         "800 µg/12h",      False, False),
    "beclometasona":            ("250-500 µg/12h",         "1.000 µg/12h",    False, False),
    "beclometasona/formoterol/glicopirroni": ("174/10/18 µg/12h","18/10/344 µg/12h",False,True),
    "fluticasona/umeclidini/vilanterol":     ("92/55/22 µg/24h","184/55/22 µg/24h",False,True),
    "budesonida/formoterol/glicopirroni":    ("160/10/14.4 µg/12h","160/10/14.4 µg/12h",False,True),
    "mometasona/indacaterol/glicopirroni":   ("136/150/50 µg/24h","136/150/50 µg/24h",False,True),
}
DOSIS_ASMA_GCI = {
    "budesonida":             ("200-400 µg/24h",  "401-800 µg/24h",   "801-1.600 µg/24h"),
    "beclometasona":          ("200-500 µg/24h",  "501-1.000 µg/24h", "1.001-2.000 µg/24h"),
    "beclometasona extrafina":("100-200 µg/24h",  "201-400 µg/24h",   ">400 µg/24h"),
    "ciclesonida":            ("80-160 µg/24h",   "161-320 µg/24h",   "321-1.280 µg/24h"),
    "fluticasona propionat":  ("100-250 µg/24h",  "251-500 µg/24h",   "501-1.000 µg/24h"),
    "fluticasona furoat":     ("92 µg/24h",        "92 µg/24h",        "184 µg/24h"),
    "mometasona":             ("200 µg/24h",       "400 µg/24h",       "800 µg/24h"),
}
DOSIS_ASMA_COMBO = {
    "salmeterol/fluticasona":   ("50/100 µg/12h",    "50/250 µg/12h",    "50/500 µg/12h"),
    "formoterol/budesonida":    ("4.5/160 µg/12h",   "9/320 µg/12h",     "2 inh de 9/320 µg/12h"),
    "formoterol/beclometasona": ("6/100: 1 inh/12h", "6/100: 2 inh/12h", "6/200: 2 inh/12h"),
    "vilanterol/fluticasona":   ("22/92 µg/24h",     "22/92 µg/24h",     "22/184 µg/24h"),
}
ATC_A_CLASSE = {
    "R03AC02":"SABA","R03AC03":"SABA","R03AC04":"SABA",
    "R03AC12":"LABA","R03AC13":"LABA","R03AC18":"LABA","R03AC19":"LABA","R03AC20":"LABA",
    "R03BB01":"SAMA","R03BB04":"LAMA","R03BB05":"LAMA","R03BB06":"LAMA","R03BB07":"LAMA",
    "R03BA01":"GCI","R03BA02":"GCI","R03BA05":"GCI","R03BA07":"GCI","R03BA08":"GCI","R03BA09":"GCI",
    "R03AL02":"LABA/LAMA","R03AL03":"LABA/LAMA","R03AL04":"LABA/LAMA","R03AL05":"LABA/LAMA",
    "R03AL06":"LABA/LAMA","R03AL09":"LABA/LAMA",
    "R03AK01":"SABA/GCI",
    "R03AK06":"LABA/GCI","R03AK07":"LABA/GCI","R03AK08":"LABA/GCI",
    "R03AK10":"LABA/GCI","R03AK11":"LABA/GCI","R03AK12":"LABA/GCI","R03AK13":"LABA/GCI",
    "R03AL08":"LAMA/LABA/GCI","R03AL10":"LAMA/LABA/GCI",
    "R03AL11":"LAMA/LABA/GCI","R03AL12":"LAMA/LABA/GCI",
}
DISPOSITIU_MAP = {
    "turbuhaler":  ("IPS-multi","🟢","50-60 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11893"),
    "accuhaler":   ("IPS-multi","🟢","60-90 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11813"),
    "genuair":     ("IPS-multi","🟢","60-90 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11878"),
    "ellipta":     ("IPS-multi","🟢","<50 l/m","Ràpida", "https://scientiasalut.gencat.cat/handle/11351/11818"),
    "novolizer":   ("IPS-multi","🟢","60-90 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11882"),
    "easyhaler":   ("IPS-multi","🟢","<50 l/m","Ràpida", "https://scientiasalut.gencat.cat/handle/11351/11817"),
    "nexthaler":   ("IPS-multi","🟢","60-90 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11881"),
    "spiromax":    ("IPS-multi","🟢","60-90 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11892"),
    "forspiro":    ("IPS-multi","🟢","60-90 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11819"),
    "twisthaler":  ("IPS-multi","🟢","<50 l/m","Ràpida", "https://scientiasalut.gencat.cat/handle/11351/11894"),
    "breezhaler":  ("IPS-uni", "🟢",">90 l/m","Ràpida",  "https://scientiasalut.gencat.cat/handle/11351/11816"),
    "aerolizer":   ("IPS-uni", "🟢",">90 l/m","Ràpida",  "https://scientiasalut.gencat.cat/handle/11351/11815"),
    "handihaler":  ("IPS-uni", "🟢","<50 l/m","Ràpida",  "https://scientiasalut.gencat.cat/handle/11351/11879"),
    "zonda":       ("IPS-uni", "🟢","<50 l/m","Ràpida",  "https://scientiasalut.gencat.cat/handle/11351/11895"),
    "respimat":    ("IVS",     "🟢","20-30 l/m","Lenta", "https://scientiasalut.gencat.cat/handle/11351/11891"),
    "modulite":    ("ICP",     "🔴","20-30 l/m","Lenta", "https://scientiasalut.gencat.cat/handle/11351/11880"),
    "aerosphere":  ("ICP",     "🔴","20-30 l/m","Lenta", "https://scientiasalut.gencat.cat/handle/11351/11880"),
    "inhalacion en envase a presion":("ICP","🔴","20-30 l/m","Lenta","https://scientiasalut.gencat.cat/handle/11351/11880"),
    "suspension para inhalacion":    ("ICP","🔴","20-30 l/m","Lenta","https://scientiasalut.gencat.cat/handle/11351/11880"),
}

# ═══════════════════════════════════════════════════════════════════════════════
# FUNCIONS AUXILIARS
# ═══════════════════════════════════════════════════════════════════════════════
def v(val):
    if val is None: return ""
    s = str(val).strip()
    return "" if s == "None" else s

def cima_get(endpoint, params=None, retries=3):
    url = f"{CIMA_BASE}/{endpoint}"
    for i in range(retries):
        try:
            r = requests.get(url, params=params, timeout=15)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            if i < retries - 1: time.sleep(2)
            else:
                print(f"  ⚠️  Error CIMA: {e}")
                return None

def get_tots_inhaladors_cima():
    print("📡 Consultant CIMA — filtrant per ATC R03 (inhaladors MPOC/Asma)...")
    resultats, pagina = [], 1
    while True:
        data = cima_get("presentaciones", {"atc": "R03", "comerc": 1, "pagina": pagina})
        if not data: break
        items = data.get("resultados", [])
        if not items: break
        resultats.extend(items)
        total = data.get("totalFilas", 0)
        print(f"  → Pàg {pagina}: {len(items)} ({len(resultats)}/{total})")
        if len(resultats) >= total: break
        pagina += 1
        time.sleep(0.3)
    print(f"  ✅ Total inhaladors R03: {len(resultats)}")
    return resultats

def get_fitxa_posologia(nregistro):
    try:
        data = cima_get("docSegmentado/contenido/1", {"nregistro": nregistro, "seccion": "4.2"})
        if not data or not data.get("secciones"): return ""
        soup = BeautifulSoup(data["secciones"][0].get("contenido",""), "html.parser")
        return soup.get_text(separator=" ", strip=True)[:400]
    except: return ""

def normalitza(text):
    t = str(text).lower()
    for p in ["bromur de ","bromur d'","propionat de ","furoat de ","dipropionat de "]:
        t = t.replace(p, "")
    return t.strip()

def cerca_dosi_mpoc(principis):
    ia = normalitza(principis)
    if ia in DOSIS_MPOC: return DOSIS_MPOC[ia]
    primer = ia.split("/")[0].strip().split(" ")[0]
    for k, val in DOSIS_MPOC.items():
        if primer in k: return val
    return None

def cerca_asma_gci(principis):
    ia = normalitza(principis)
    for k, val in DOSIS_ASMA_GCI.items():
        if k in ia or ia in k: return val
    return None

def cerca_asma_combo(principis):
    ia = normalitza(principis)
    for k, val in DOSIS_ASMA_COMBO.items():
        if k in ia or ia in k: return val
    return None

def infereix_classe(atcs):
    for atc in (atcs or []):
        c = atc.get("codigo","").upper()
        if c in ATC_A_CLASSE: return ATC_A_CLASSE[c]
        if c[:7] in ATC_A_CLASSE: return ATC_A_CLASSE[c[:7]]
    return "PENDENT"

def infereix_dispositiu(nom):
    nl = str(nom).lower()
    for k, val in DISPOSITIU_MAP.items():
        if k in nl: return val
    return ("ICP","🔴","20-30 l/m","Lenta","https://scientiasalut.gencat.cat/handle/11351/11880")

# ═══════════════════════════════════════════════════════════════════════════════
# MODES D'EXECUCIÓ
# ═══════════════════════════════════════════════════════════════════════════════

def mode_detecta():
    """Consulta CIMA, afegeix novetats al Excel, escriu novetats_count.txt"""
    print("🔍 Mode: detectar novetats CIMA")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active

    cns = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        cn = v(row[I_CN])
        if cn: cns.add(cn)
    print(f"  CNs existents: {len(cns)}")

    presentacions = get_tots_inhaladors_cima()
    novetats = [p for p in presentacions if v(p.get("cn","")) not in cns]
    print(f"  Novetats: {len(novetats)}")

    if not novetats:
        with open("novetats_count.txt", "w") as f: f.write("0")
        print("✅ Cap novetat")
        return

    shutil.copy(EXCEL_PATH, EXCEL_PATH.replace(".xlsx","_BACKUP.xlsx"))

    fill_nou     = PatternFill("solid", fgColor=COLOR_NOU)
    fill_atencio = PatternFill("solid", fgColor=COLOR_ATENCIO)
    font_base    = Font(name='Arial', size=10)
    alineacio    = Alignment(vertical="top", wrap_text=True)

    afegits = 0
    for i, p in enumerate(novetats[:50], 1):
        cn        = v(p.get("cn",""))
        nom       = v(p.get("nombre",""))
        nregistro = v(p.get("nregistro",""))
        print(f"  [{i}] {nom[:50]}")

        med       = cima_get("medicamento", {"cn": cn}) or {}
        atcs      = med.get("atcs", [])
        pas       = med.get("principiosActivos", [])
        principis = ", ".join([x.get("nombre","") for x in pas])

        classe = infereix_classe(atcs)
        tipus, co2, flux, flux4, link = infereix_dispositiu(nom)

        dosi_info  = cerca_dosi_mpoc(principis)
        asma_gci   = cerca_asma_gci(principis)
        asma_combo = cerca_asma_combo(principis)

        dosi_text = ""
        phf_val = matma_val = ""
        color = fill_nou

        if dosi_info:
            dosi_mpoc, dmx, starred, matma = dosi_info
            dosi_text = f"MPOC: {dosi_mpoc}\nD.màx: {dmx}"
            if starred:  phf_val   = "★"
            if matma:    matma_val = "MATMA"

        if asma_gci:
            dosi_text += f"\nAsma baixa: {asma_gci[0]}\nAsma mitjana: {asma_gci[1]}\nAsma alta: {asma_gci[2]}"
        elif asma_combo:
            dosi_text += f"\nAsma baixa: {asma_combo[0]}\nAsma mitjana: {asma_combo[1]}\nAsma alta: {asma_combo[2]}"

        if not dosi_text:
            posologia = get_fitxa_posologia(nregistro)
            dosi_text = f"FITXA TÈCNICA: {posologia[:200]}" if posologia else "PENDENT — consultar fitxa CIMA"
            color = fill_atencio

        fila = [
            classe, principis, nom, cn, nom,
            dosi_text, tipus, co2, flux, flux4, link,
            datetime.now().strftime("%Y-%m-%d"),
            f"https://cima.aemps.es/cima/publico/medicamento.html?nregistro={nregistro}",
            "NOU — PENDENT VALIDACIÓ CLÍNICA",
            phf_val, matma_val,
        ]

        ws.append(fila)
        nr = ws.max_row
        ws.row_dimensions[nr].height = 40
        for col in range(1, len(fila)+1):
            cel = ws.cell(row=nr, column=col)
            cel.fill = color
            cel.font = font_base
            cel.alignment = alineacio

        afegits += 1
        time.sleep(0.4)

    wb.save(EXCEL_PATH)
    with open("novetats_count.txt", "w") as f: f.write(str(afegits))
    print(f"✅ {afegits} novetats afegides al Excel")

def mode_comprova_pendents():
    """Compta files amb col N no buida (pendents de validació)"""
    print("🔍 Mode: comprovar pendents")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
    pendents = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) > I_ESTAT and v(row[I_ESTAT]):
            pendents += 1
    with open("pendents_count.txt", "w") as f: f.write(str(pendents))
    print(f"  Pendents: {pendents}")

def mode_comprova_publicar():
    """Compta files validades que encara no s'han publicat al HTML"""
    print("🔍 Mode: comprovar per publicar")
    wb_excel = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb_excel.active

    with open(HTML_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    pendents_publicar = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) <= I_ESTAT: continue
        nom    = v(row[I_NOM])
        data   = v(row[I_DATA])
        estat  = v(row[I_ESTAT])
        if not nom: continue
        if data and not estat:
            if nom not in html:
                pendents_publicar += 1

    with open("publicar_count.txt", "w") as f: f.write(str(pendents_publicar))
    print(f"  Per publicar: {pendents_publicar}")

def mode_regenera():
    """Regenera HTML amb NOMÉS els fàrmacs validats (col N buida)"""
    print(f"🔄 Regenerant HTML: {OUTPUT_HTML}")
    with open(HTML_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
    entrades = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row): continue
        nom   = v(row[I_NOM])
        if not nom: continue

        estat = v(row[I_ESTAT])
        if estat:
            print(f"  ⏭  Saltat (pendent): {nom[:40]}")
            continue

        classe = v(row[I_CLASSE])
        ia     = v(row[I_IA])
        disp   = v(row[I_DISP])
        dosi_r = v(row[I_DOSI])
        tipus_r= v(row[I_TIPUS])
        link   = v(row[I_LINK])

# ═══════════════════════════════════════════════════════════════════════════════
# ENTRADA PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    args = sys.argv[1:]
    if "--detecta" in args:
        mode_detecta()
    elif "--comprova-pendents" in args:
        mode_comprova_pendents()
    elif "--comprova-publicar" in args:
        mode_comprova_publicar()
    elif "--regenera" in args:
        mode_regenera()
    elif "--tot" in args:
        mode_detecta()
        mode_regenera()
    else:
        print("⚠️  Cap mode especificat. Usa: --detecta | --comprova-pendents | --comprova-publicar | --regenera | --tot")
