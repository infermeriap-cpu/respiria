#!/usr/bin/env python3
"""
RespirIA — Actualitzador automàtic CIMA → Excel → HTML
=======================================================
Autora: Sílvia Álvarez Vega · ICS Atenció Primària Girona
Versió: 4.0 · 2026-04-12

Canvis v4.0:
  - Llegeix columnes PHF (O) i MATMA (P) del nou Excel robust
  - Cap cel·la combinada — totes les files independents
  - starred i matma llegits directament del Excel (no inferits)
  - Flux de validació: columna N buida = validat, "NOU" = pendent

Ús:
    python respiria_cima_updater.py              # detecta novetats CIMA
    python respiria_cima_updater.py --regenera   # regenera HTML des del Excel
    python respiria_cima_updater.py --tot        # fa les dues coses

Requisits:
    pip install requests openpyxl beautifulsoup4
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
OUTPUT_HTML = "RespirIA_v2.html"   # sobreescriu el mateix fitxer
CIMA_BASE   = "https://cima.aemps.es/cima/rest"

# Columnes del Excel (índex base 1)
COL_CLASSE  = 1
COL_IA      = 2
COL_NOM     = 3
COL_CN      = 4
COL_DISP    = 5
COL_DOSI    = 6
COL_TIPUS   = 7
COL_CO2     = 8
COL_FLUX    = 9
COL_FLUX4   = 10
COL_LINK    = 11
COL_DATA    = 12
COL_ORIGEN  = 13
COL_ESTAT   = 14
COL_PHF     = 15   # NOU — columna O
COL_MATMA   = 16   # NOU — columna P

# Colors
COLOR_NOU      = "FEF3C7"  # groc  = pendent validació
COLOR_ATENCIO  = "FEE2E2"  # vermell = consultar fitxa
COLOR_VALIDAT  = "D1FAE5"  # verd = validat (manual)

# ═══════════════════════════════════════════════════════════════════════════════
# DOSIS DE REFERÈNCIA — PHF CatSalut 2018 + GEMA 5.5
# ═══════════════════════════════════════════════════════════════════════════════
DOSIS_MPOC = {
    "salbutamol":               ("100-200 µg si cal",    "200 µg/6h",          True,  False),
    "terbutalina":              ("500 µg si cal",         "6.000 µg/24h",       False, False),
    "ipratropi":                ("40 µg si cal",          "240 µg/24h",         True,  False),
    "formoterol":               ("12 µg/12h",             "24 µg/12h",          True,  False),
    "salmeterol":               ("50 µg/12h",             "100 µg/12h",         True,  False),
    "indacaterol":              ("150 µg/24h",            "300 µg/24h",         True,  False),
    "olodaterol":               ("5 µg/24h",              "5 µg/24h",           False, False),
    "tiotropi":                 ("18 µg/24h (Handihaler) / 5 µg/24h (Respimat) / 10 µg/24h (Zonda)",
                                 "= pauta",               True,  False),
    "glicopirroni":             ("44 µg/24h",             "44 µg/24h",          False, False),
    "umeclidini":               ("55 µg/24h",             "55 µg/24h",          False, False),
    "aclidini":                 ("322 µg/12h",            "322 µg/12h",         False, False),
    "indacaterol/glicopirroni": ("85/43 µg/24h",          "85/43 µg/24h",       True,  False),
    "tiotropi/olodaterol":      ("5/5 µg/24h",            "5/5 µg/24h",         False, False),
    "umeclidini/vilanterol":    ("55/22 µg/24h",          "55/22 µg/24h",       False, False),
    "aclidini/formoterol":      ("340/12 µg/12h",         "340/12 µg/12h",      False, False),
    "salmeterol/fluticasona":   ("50/500 µg/12h",         "50/500 µg/12h",      False, False),
    "formoterol/budesonida":    ("9/320 µg/12h",          "36/1.280 µg/24h",    False, False),
    "formoterol/beclometasona": ("12/200 µg/12h",         "12/400 µg/12h",      False, False),
    "vilanterol/fluticasona":   ("22/92 µg/24h",          "22/184 µg/24h",      False, False),
    "fluticasona":              ("250-500 µg/12h",         "500 µg/12h",         False, False),
    "budesonida":               ("200-400 µg/12h",         "800 µg/12h",         False, False),
    "beclometasona":            ("250-500 µg/12h",         "1.000 µg/12h",       False, False),
    "beclometasona/formoterol/glicopirroni": ("174/10/18 µg/12h","18/10/344 µg/12h",False,True),
    "fluticasona/umeclidini/vilanterol":     ("92/55/22 µg/24h", "184/55/22 µg/24h",False,True),
    "budesonida/formoterol/glicopirroni":    ("160/10/14.4 µg/12h","160/10/14.4 µg/12h",False,True),
    "mometasona/indacaterol/glicopirroni":   ("136/150/50 µg/24h","136/150/50 µg/24h",False,True),
}

DOSIS_ASMA_GCI = {
    "budesonida":            ("200-400 µg/24h",  "401-800 µg/24h",   "801-1.600 µg/24h"),
    "beclometasona":         ("200-500 µg/24h",  "501-1.000 µg/24h", "1.001-2.000 µg/24h"),
    "beclometasona extrafina":("100-200 µg/24h", "201-400 µg/24h",   ">400 µg/24h"),
    "ciclesonida":           ("80-160 µg/24h",   "161-320 µg/24h",   "321-1.280 µg/24h"),
    "fluticasona propionat": ("100-250 µg/24h",  "251-500 µg/24h",   "501-1.000 µg/24h"),
    "fluticasona furoat":    ("92 µg/24h",        "92 µg/24h",        "184 µg/24h"),
    "mometasona":            ("200 µg/24h",       "400 µg/24h",       "800 µg/24h"),
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
    "turbuhaler":   ("IPS-multi","🟢","50-60 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11893"),
    "accuhaler":    ("IPS-multi","🟢","60-90 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11813"),
    "genuair":      ("IPS-multi","🟢","60-90 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11878"),
    "ellipta":      ("IPS-multi","🟢","<50 l/m","Ràpida", "https://scientiasalut.gencat.cat/handle/11351/11818"),
    "novolizer":    ("IPS-multi","🟢","60-90 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11882"),
    "easyhaler":    ("IPS-multi","🟢","<50 l/m","Ràpida", "https://scientiasalut.gencat.cat/handle/11351/11817"),
    "nexthaler":    ("IPS-multi","🟢","60-90 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11881"),
    "spiromax":     ("IPS-multi","🟢","60-90 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11892"),
    "forspiro":     ("IPS-multi","🟢","60-90 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11819"),
    "twisthaler":   ("IPS-multi","🟢","<50 l/m","Ràpida", "https://scientiasalut.gencat.cat/handle/11351/11894"),
    "breezhaler":   ("IPS-uni", "🟢",">90 l/m","Ràpida",  "https://scientiasalut.gencat.cat/handle/11351/11816"),
    "aerolizer":    ("IPS-uni", "🟢",">90 l/m","Ràpida",  "https://scientiasalut.gencat.cat/handle/11351/11815"),
    "handihaler":   ("IPS-uni", "🟢","<50 l/m","Ràpida",  "https://scientiasalut.gencat.cat/handle/11351/11879"),
    "zonda":        ("IPS-uni", "🟢","<50 l/m","Ràpida",  "https://scientiasalut.gencat.cat/handle/11351/11895"),
    "respimat":     ("IVS",     "🟢","20-30 l/m","Lenta", "https://scientiasalut.gencat.cat/handle/11351/11891"),
    "modulite":     ("ICP",     "🔴","20-30 l/m","Lenta", "https://scientiasalut.gencat.cat/handle/11351/11880"),
    "aerosphere":   ("ICP",     "🔴","20-30 l/m","Lenta", "https://scientiasalut.gencat.cat/handle/11351/11880"),
    "inhalacion en envase a presion":("ICP","🔴","20-30 l/m","Lenta","https://scientiasalut.gencat.cat/handle/11351/11880"),
    "suspension para inhalacion":    ("ICP","🔴","20-30 l/m","Lenta","https://scientiasalut.gencat.cat/handle/11351/11880"),
}

# ═══════════════════════════════════════════════════════════════════════════════
# FUNCIONS AUXILIARS
# ═══════════════════════════════════════════════════════════════════════════════

def cima_get(endpoint, params=None, retries=3):
    url = f"{CIMA_BASE}/{endpoint}"
    for i in range(retries):
        try:
            r = requests.get(url, params=params, timeout=15)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            if i < retries - 1:
                time.sleep(2)
            else:
                print(f"  ⚠️  Error CIMA: {e}")
                return None

def get_tots_inhaladors_cima():
    print("📡 Consultant CIMA (sust=4&comerc=1)...")
    resultats, pagina = [], 1
    while True:
        data = cima_get("presentaciones", {"sust": 4, "comerc": 1, "pagina": pagina})
        if not data:
            break
        items = data.get("resultados", [])
        if not items:
            break
        resultats.extend(items)
        total = data.get("totalFilas", 0)
        print(f"  → Pàg {pagina}: {len(items)} ({len(resultats)}/{total})")
        if len(resultats) >= total:
            break
        pagina += 1
        time.sleep(0.3)
    print(f"  ✅ Total: {len(resultats)} presentacions")
    return resultats

def get_fitxa_posologia(nregistro):
    try:
        data = cima_get("docSegmentado/contenido/1",
                        {"nregistro": nregistro, "seccion": "4.2"})
        if not data or not data.get("secciones"):
            return ""
        soup = BeautifulSoup(data["secciones"][0].get("contenido",""), "html.parser")
        return soup.get_text(separator=" ", strip=True)[:500]
    except:
        return ""

def normalitza(text):
    t = str(text).lower()
    for p in ["bromur de ","bromur d'","propionat de ","furoat de ","dipropionat de "]:
        t = t.replace(p, "")
    return t.strip()

def cerca_dosi_mpoc(principis):
    ia = normalitza(principis)
    if ia in DOSIS_MPOC:
        return DOSIS_MPOC[ia]
    primer = ia.split("/")[0].strip().split(" ")[0]
    for k, v in DOSIS_MPOC.items():
        if primer in k:
            return v
    return None

def cerca_asma_gci(principis):
    ia = normalitza(principis)
    for k, v in DOSIS_ASMA_GCI.items():
        if k in ia or ia in k:
            return v
    return None

def cerca_asma_combo(principis):
    ia = normalitza(principis)
    for k, v in DOSIS_ASMA_COMBO.items():
        if k in ia or ia in k:
            return v
    return None

def infereix_classe(atcs):
    for atc in (atcs or []):
        c = atc.get("codigo","").upper()
        if c in ATC_A_CLASSE:
            return ATC_A_CLASSE[c]
        if c[:7] in ATC_A_CLASSE:
            return ATC_A_CLASSE[c[:7]]
    return "PENDENT"

def infereix_dispositiu(nom):
    nl = str(nom).lower()
    for k, v in DISPOSITIU_MAP.items():
        if k in nl:
            return v
    return ("ICP","🔴","20-30 l/m","Lenta",
            "https://scientiasalut.gencat.cat/handle/11351/11880")

def v(val):
    """Retorna string net o buit."""
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s == "None" else s

# ═══════════════════════════════════════════════════════════════════════════════
# LECTURA DEL EXCEL
# ═══════════════════════════════════════════════════════════════════════════════

def llegeix_excel():
    print(f"📊 Llegint Excel: {EXCEL_PATH}")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
    cns = set()
    files = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        cn = v(row[COL_CN - 1])
        if cn:
            cns.add(cn)
        files += 1
    print(f"  ✅ {files} files, {len(cns)} CNs registrats")
    return wb, ws, cns

# ═══════════════════════════════════════════════════════════════════════════════
# AFEGIR NOVETATS AL EXCEL
# ═══════════════════════════════════════════════════════════════════════════════

def afegeix_novetats(ws, novetats_raw):
    print(f"\n🧬 Enriquint i afegint {min(len(novetats_raw),50)} novetats...")

    fill_nou     = PatternFill("solid", fgColor=COLOR_NOU)
    fill_atencio = PatternFill("solid", fgColor=COLOR_ATENCIO)
    font_base    = Font(name='Arial', size=10)
    alineacio    = Alignment(vertical="top", wrap_text=True)

    afegits = 0
    for i, p in enumerate(novetats_raw[:50], 1):
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

        # Dosis
        dosi_info  = cerca_dosi_mpoc(principis)
        asma_gci   = cerca_asma_gci(principis)
        asma_combo = cerca_asma_combo(principis)

        dosi_text = ""
        phf_val   = ""
        matma_val = ""
        color     = fill_nou

        if dosi_info:
            dosi_mpoc, dmx, starred, matma = dosi_info
            dosi_text = f"MPOC: {dosi_mpoc}\nD.màx: {dmx}"
            if starred:
                phf_val = "★"
            if matma:
                matma_val = "MATMA"

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
            phf_val,
            matma_val,
        ]

        ws.append(fila)
        nova_fila = ws.max_row
        ws.row_dimensions[nova_fila].height = 40

        for col in range(1, len(fila) + 1):
            cel = ws.cell(row=nova_fila, column=col)
            cel.fill = color
            cel.font = font_base
            cel.alignment = alineacio

        afegits += 1
        time.sleep(0.4)

    return afegits

# ═══════════════════════════════════════════════════════════════════════════════
# REGENERACIÓ HTML
# ═══════════════════════════════════════════════════════════════════════════════

def regenera_html(ws):
    print(f"\n🔄 Generant HTML: {OUTPUT_HTML}")
    with open(HTML_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    entrades = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue

        nom   = v(row[COL_NOM - 1])
        if not nom:
            continue

        # Salta files pendents de validació
        estat = v(row[COL_ESTAT - 1])
        if "NOU" in estat.upper() or "PENDENT" in estat.upper():
            print(f"  ⏭  Saltat (pendent): {nom[:40]}")
            continue

        classe = v(row[COL_CLASSE - 1])
        ia     = v(row[COL_IA - 1])
        disp   = v(row[COL_DISP - 1])
        dosi_r = v(row[COL_DOSI - 1])
        tipus_r= v(row[COL_TIPUS - 1])
        link   = v(row[COL_LINK - 1])

        # Columnes noves
        phf_val   = v(row[COL_PHF - 1])
        matma_val = v(row[COL_MATMA - 1])

        starred = "★" in phf_val
        matma   = bool(matma_val)

        # Normalitza tipus
        t = tipus_r.lower().replace(" ","").replace("_","-")
        if "multi" in t:              tipus_js = "IPS-multi"
        elif "uni" in t:              tipus_js = "IPS-uni"
        elif "ivs" in t or "respimat" in t: tipus_js = "IVS"
        else:                         tipus_js = "ICP"

        # Parseja dosis
        mpoc = asma_b = asma_m = asma_a = dmx = ""
        for l in dosi_r.split("\n"):
            l = l.strip()
            ll = l.lower()
            if not l:
                continue
            if "mpoc" in ll or "epoc" in ll:
                mpoc = l.split(":")[-1].strip() if ":" in l else l
            elif "baixa" in ll or "baja" in ll:
                asma_b = l.split(":")[-1].strip() if ":" in l else l
            elif "mitjan" in ll or "media" in ll:
                asma_m = l.split(":")[-1].strip() if ":" in l else l
            elif "alta" in ll:
                asma_a = l.split(":")[-1].strip() if ":" in l else l
            elif "màx" in ll or "max" in ll:
                dmx = l.split(":")[-1].strip() if ":" in l else l
            elif not mpoc and not asma_b:
                mpoc = l

        obj = f"  {{cat:{json.dumps(classe, ensure_ascii=False)}"
        if ia:       obj += f",ia:{json.dumps(ia, ensure_ascii=False)}"
        if starred:  obj += ",starred:true"
        if matma:    obj += ",matma:true"
        obj += f",brands:{json.dumps(nom, ensure_ascii=False)}"
        obj += f",disp:{json.dumps(disp, ensure_ascii=False)}"
        if mpoc and not asma_b:
            obj += f",dose:{json.dumps(mpoc, ensure_ascii=False)}"
        elif mpoc:
            obj += f",mpoc:{json.dumps(mpoc, ensure_ascii=False)}"
        if asma_b:  obj += f",asma_baixa:{json.dumps(asma_b, ensure_ascii=False)}"
        if asma_m:  obj += f",asma_mitjana:{json.dumps(asma_m, ensure_ascii=False)}"
        if asma_a:  obj += f",asma_alta:{json.dumps(asma_a, ensure_ascii=False)}"
        if dmx:     obj += f",dmx:{json.dumps(dmx, ensure_ascii=False)}"
        obj += f",tipus:{json.dumps(tipus_js, ensure_ascii=False)}"
        if link:    obj += f",link:{json.dumps(link, ensure_ascii=False)}"
        obj += "}"
        entrades.append(obj)

    nou_catalog = "var catalog=[\n" + ",\n".join(entrades) + "\n];"
    start = html.index("var catalog=[")
    end   = html.index("];", start) + 2
    html_nou = html[:start] + nou_catalog + html[end:]

    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html_nou)
    print(f"  ✅ HTML generat amb {len(entrades)} entrades")

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    print()
    print("╔══════════════════════════════════════════════════════════╗")
    print("║  RespirIA — Actualitzador CIMA v4.0                      ║")
    print("║  PHF CatSalut 2018 · GEMA 5.5 · Columnes PHF+MATMA      ║")
    print(f"║  {datetime.now().strftime('%Y-%m-%d %H:%M')}                                        ║")
    print("╚══════════════════════════════════════════════════════════╝\n")

    for f in [EXCEL_PATH, HTML_PATH]:
        if not os.path.exists(f):
            print(f"❌ No trobo: {f}")
            print("   Posa el script a la mateixa carpeta que el Excel i el HTML.")
            return

    mode_regenera = "--regenera" in sys.argv
    mode_tot      = "--tot" in sys.argv

    wb, ws, cns_existents = llegeix_excel()

    if not mode_regenera:
        presentacions = get_tots_inhaladors_cima()
        novetats = [p for p in presentacions
                    if v(p.get("cn","")) not in cns_existents]
        print(f"\n🔍 Novetats: {len(novetats)}")

        if novetats:
            shutil.copy(EXCEL_PATH, EXCEL_PATH.replace(".xlsx","_BACKUP.xlsx"))
            afegits = afegeix_novetats(ws, novetats)
            wb.save(EXCEL_PATH)
            print(f"\n  ✅ Excel actualitzat: {afegits} fàrmacs nous afegits")
            print(f"  💾 Backup: {EXCEL_PATH.replace('.xlsx','_BACKUP.xlsx')}")
            print()
            print("━"*62)
            print("✋  REVISA EL EXCEL:")
            print("  · Files noves al final — columna N = 'NOU - PENDENT...'")
            print("  · Revisa dosi (col F), PHF (col O), MATMA (col P)")
            print("  · Quan validis una fila: esborra el text de la col N")
            print("  · Puja el Excel a GitHub")
            print("  · Executa: python respiria_cima_updater.py --regenera")
            print("━"*62)
        else:
            print("✅ Cap novetat. El catàleg ja està al dia.")

    if mode_regenera or mode_tot:
        wb2 = openpyxl.load_workbook(EXCEL_PATH)
        ws2 = wb2.active
        regenera_html(ws2)
        print(f"\n✅ HTML actualitzat: {OUTPUT_HTML}")

    print(f"\nFet! {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")

if __name__ == "__main__":
    main()
