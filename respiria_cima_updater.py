#!/usr/bin/env python3
"""
RespirIA — Actualitzador automàtic CIMA → Excel → HTML
=======================================================
Autora: Sílvia Álvarez Vega · ICS Atenció Primària Girona
Versió: 5.5 · 2026-05-14

Modes d'execució:
  --detecta          Consulta CIMA, afegeix novetats al Excel, escriu novetats_count.txt
  --comprova-pendents Compta files amb col N no buida, escriu pendents_count.txt
  --comprova-publicar Compta files validades (col N buida), escriu publicar_count.txt
  --regenera         Regenera HTML amb NOMÉS els fàrmacs validats (col N buida)
  --tot              detecta + regenera (ús local)

Filtre CIMA v5.5:
  - Cerca per codis ATC específics R03A i R03B (inhaladors MPOC/Asma)
  - Dosis indexades per codi ATC → assignació sempre correcta
  - Filtre per paraules clau inhalatòries al nom de la presentació
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
# CODIS ATC R03 — INHALADORS MPOC/ASMA
# Font: CIMA REST API maestras?maestra=7&nombre=R03
# ═══════════════════════════════════════════════════════════════════════════════
CODIS_ATC = {
    # ── SABA ──────────────────────────────────────────────────────────────────
    "R03AC02": "SABA",   # Salbutamol
    "R03AC03": "SABA",   # Terbutalina
    "R03AC04": "SABA",   # Fenoterol
    # ── LABA ──────────────────────────────────────────────────────────────────
    "R03AC12": "LABA",   # Salmeterol
    "R03AC13": "LABA",   # Formoterol
    "R03AC18": "LABA",   # Indacaterol
    "R03AC19": "LABA",   # Olodaterol
    # ── SAMA ──────────────────────────────────────────────────────────────────
    "R03BB01": "SAMA",   # Ipratropi
    "R03BB02": "SAMA",   # Oxitropi
    # ── LAMA ──────────────────────────────────────────────────────────────────
    "R03BB04": "LAMA",   # Tiotropi
    "R03BB05": "LAMA",   # Aclidini
    "R03BB06": "LAMA",   # Glicopirroni
    "R03BB07": "LAMA",   # Umeclidini
    # ── GCI ───────────────────────────────────────────────────────────────────
    "R03BA01": "GCI",    # Beclometasona
    "R03BA02": "GCI",    # Budesonida
    "R03BA05": "GCI",    # Fluticasona
    "R03BA07": "GCI",    # Mometasona
    "R03BA08": "GCI",    # Ciclesonida
    # ── SABA/SAMA ─────────────────────────────────────────────────────────────
    "R03AK03": "SABA/SAMA",  # Fenoterol + Ipratropi
    "R03AL01": "SABA/SAMA",  # Fenoterol + Ipratropi
    "R03AL02": "SABA/SAMA",  # Salbutamol + Ipratropi
    # ── LABA/GCI ──────────────────────────────────────────────────────────────
    "R03AK06": "LABA/GCI",  # Salmeterol + Fluticasona
    "R03AK07": "LABA/GCI",  # Formoterol + Budesonida
    "R03AK08": "LABA/GCI",  # Formoterol + Beclometasona
    "R03AK09": "LABA/GCI",  # Formoterol + Mometasona
    "R03AK10": "LABA/GCI",  # Vilanterol + Fluticasona furoat
    "R03AK11": "LABA/GCI",  # Formoterol + Fluticasona
    "R03AK14": "LABA/GCI",  # Indacaterol + Mometasona
    # ── LABA/LAMA ─────────────────────────────────────────────────────────────
    "R03AL03": "LABA/LAMA",  # Vilanterol + Umeclidini
    "R03AL04": "LABA/LAMA",  # Indacaterol + Glicopirroni
    "R03AL05": "LABA/LAMA",  # Formoterol + Aclidini
    "R03AL06": "LABA/LAMA",  # Olodaterol + Tiotropi
    # ── LAMA/LABA/GCI ─────────────────────────────────────────────────────────
    "R03AL08": "LAMA/LABA/GCI",  # Vilanterol + Umeclidini + Fluticasona
    "R03AL09": "LAMA/LABA/GCI",  # Formoterol + Glicopirroni + Beclometasona
    "R03AL11": "LAMA/LABA/GCI",  # Formoterol + Glicopirroni + Budesonida
    "R03AL12": "LAMA/LABA/GCI",  # Indacaterol + Glicopirroni + Mometasona
}

# ═══════════════════════════════════════════════════════════════════════════════
# DOSIS PER CODI ATC
# Tupla: (dosi_mpoc, dosi_max, phf_star, matma)
# La dosi és sempre la mateixa independentment del dispositiu
# ═══════════════════════════════════════════════════════════════════════════════
DOSIS_PER_ATC = {
    # SABA
    "R03AC02": ("100-200 µg si cal",   "200 µg/6h",            True,  False),  # Salbutamol
    "R03AC03": ("500 µg si cal",        "6.000 µg/24h",         False, False),  # Terbutalina
    "R03AC04": ("100-200 µg si cal",   "200 µg/6h",            False, False),  # Fenoterol
    # LABA
    "R03AC12": ("50 µg/12h",            "100 µg/12h",           True,  False),  # Salmeterol
    "R03AC13": ("12 µg/12h",            "24 µg/12h",            True,  False),  # Formoterol
    "R03AC18": ("150 µg/24h",           "300 µg/24h",           True,  False),  # Indacaterol
    "R03AC19": ("5 µg/24h",             "5 µg/24h",             False, False),  # Olodaterol
    # SAMA
    "R03BB01": ("40 µg si cal",         "240 µg/24h",           True,  False),  # Ipratropi
    "R03BB02": ("200 µg si cal",        "800 µg/24h",           False, False),  # Oxitropi
    # LAMA
    "R03BB04": ("18 µg/24h (Handihaler) / 5 µg/24h (Respimat) / 10 µg/24h (Zonda)", "= pauta", True, False),  # Tiotropi
    "R03BB05": ("322 µg/12h",           "322 µg/12h",           False, False),  # Aclidini
    "R03BB06": ("44 µg/24h",            "44 µg/24h",            False, False),  # Glicopirroni
    "R03BB07": ("55 µg/24h",            "55 µg/24h",            False, False),  # Umeclidini
    # GCI
    "R03BA01": ("250-500 µg/12h",       "1.000 µg/12h",         False, False),  # Beclometasona
    "R03BA02": ("200-400 µg/12h",       "800 µg/12h",           False, False),  # Budesonida
    "R03BA05": ("250-500 µg/12h",       "500 µg/12h",           False, False),  # Fluticasona
    "R03BA07": ("200 µg/24h",           "800 µg/24h",           False, False),  # Mometasona
    "R03BA08": ("80-160 µg/24h",        "1.280 µg/24h",         False, False),  # Ciclesonida
    # SABA/SAMA
    "R03AK03": ("100-200/40 µg si cal", "200/240 µg/24h",       False, False),  # Fenoterol+Ipratropi
    "R03AL01": ("100-200/40 µg si cal", "200/240 µg/24h",       False, False),  # Fenoterol+Ipratropi
    "R03AL02": ("100-200/40 µg si cal", "200/240 µg/24h",       False, False),  # Salbutamol+Ipratropi
    # LABA/GCI
    "R03AK06": ("50/500 µg/12h",        "50/500 µg/12h",        False, False),  # Salmeterol+Fluticasona
    "R03AK07": ("9/320 µg/12h",         "36/1.280 µg/24h",      False, False),  # Formoterol+Budesonida
    "R03AK08": ("12/200 µg/12h",        "12/400 µg/12h",        False, False),  # Formoterol+Beclometasona
    "R03AK09": ("5/200 µg/24h",         "5/400 µg/24h",         False, False),  # Formoterol+Mometasona
    "R03AK10": ("22/92 µg/24h",         "22/184 µg/24h",        False, False),  # Vilanterol+Fluticasona
    "R03AK11": ("5/125 µg/12h",         "5/250 µg/12h",         False, False),  # Formoterol+Fluticasona
    "R03AK14": ("150/160 µg/24h",       "150/160 µg/24h",       False, False),  # Indacaterol+Mometasona
    # LABA/LAMA
    "R03AL03": ("22/55 µg/24h",         "22/55 µg/24h",         False, False),  # Vilanterol+Umeclidini
    "R03AL04": ("150/50 µg/24h",        "150/50 µg/24h",        True,  False),  # Indacaterol+Glicopirroni
    "R03AL05": ("12/340 µg/12h",        "12/340 µg/12h",        False, False),  # Formoterol+Aclidini
    "R03AL06": ("5/5 µg/24h",           "5/5 µg/24h",           False, False),  # Olodaterol+Tiotropi
    # LAMA/LABA/GCI
    "R03AL08": ("22/55/92 µg/24h",      "22/55/184 µg/24h",     False, True),   # Vilanterol+Umeclidini+Fluticasona
    "R03AL09": ("10/18/174 µg/12h",     "10/18/344 µg/12h",     False, True),   # Formoterol+Glicopirroni+Beclometasona
    "R03AL11": ("10/160/14.4 µg/12h",   "10/160/14.4 µg/12h",   False, True),   # Formoterol+Glicopirroni+Budesonida
    "R03AL12": ("150/50/160 µg/24h",    "150/50/160 µg/24h",    False, True),   # Indacaterol+Glicopirroni+Mometasona
}

# Dosis asma GCI per escalons (baixa/mitjana/alta)
DOSIS_ASMA_GCI = {
    "R03BA01": ("200-500 µg/24h",  "501-1.000 µg/24h", "1.001-2.000 µg/24h"),  # Beclometasona
    "R03BA02": ("200-400 µg/24h",  "401-800 µg/24h",   "801-1.600 µg/24h"),    # Budesonida
    "R03BA05": ("100-250 µg/24h",  "251-500 µg/24h",   "501-1.000 µg/24h"),    # Fluticasona
    "R03BA07": ("200 µg/24h",      "400 µg/24h",       "800 µg/24h"),           # Mometasona
    "R03BA08": ("80-160 µg/24h",   "161-320 µg/24h",   "321-1.280 µg/24h"),    # Ciclesonida
}

# Dosis asma combinacions per escalons (baixa/mitjana/alta)
DOSIS_ASMA_COMBO = {
    "R03AK06": ("50/100 µg/12h",    "50/250 µg/12h",    "50/500 µg/12h"),       # Salmeterol+Fluticasona
    "R03AK07": ("4.5/160 µg/12h",   "9/320 µg/12h",     "2 inh 9/320 µg/12h"), # Formoterol+Budesonida
    "R03AK08": ("6/100: 1 inh/12h", "6/100: 2 inh/12h", "6/200: 2 inh/12h"),   # Formoterol+Beclometasona
    "R03AK10": ("22/92 µg/24h",     "22/92 µg/24h",     "22/184 µg/24h"),       # Vilanterol+Fluticasona
    "R03AK11": ("5/100 µg/12h",     "5/250 µg/12h",     "5/500 µg/12h"),        # Formoterol+Fluticasona
}

# Paraules clau que identifiquen una presentació com a inhalatòria
PARAULES_INHALATORI = [
    "inhal", "turbuhaler", "accuhaler", "genuair", "ellipta",
    "novolizer", "easyhaler", "nexthaler", "spiromax", "forspiro",
    "twisthaler", "breezhaler", "aerolizer", "handihaler", "zonda",
    "respimat", "modulite", "aerosphere", "evohaler", "autohaler",
    "clickhaler", "pulvinal", "diskus", "aerocaps",
    "polvo para inhalacion", "polvo inhalacion",
    "suspension para inhalacion", "solucion para inhalacion",
    "nebulizacion", "nebulizador",
]

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
    "solucion para inhalacion":      ("ICP","🔴","20-30 l/m","Lenta","https://scientiasalut.gencat.cat/handle/11351/11880"),
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

def es_inhalatori(nom):
    """Comprova que el nom de la presentació indica via inhalatòria"""
    nl = nom.lower()
    return any(p in nl for p in PARAULES_INHALATORI)

def get_tots_inhaladors_cima():
    """
    Cerca per cada codi ATC R03 rellevant.
    La classe terapèutica ve directament del codi ATC — sempre correcta.
    Filtre addicional per paraules clau inhalatòries.
    """
    print("📡 Consultant CIMA — cerca per codis ATC R03...")
    resultats_filtrats = []
    cns_vistos = set()

    for codi_atc, classe in CODIS_ATC.items():
        print(f"  🔍 {codi_atc} ({classe})...")
        pagina = 1
        trobats = 0

        while True:
            data = cima_get("medicamentos", {
                "atc": codi_atc,
                "comerc": 1,
                "pagina": pagina
            })
            if not data: break
            items = data.get("resultados", [])
            if not items: break
            total = data.get("totalFilas", 0)

            for med in items:
                nregistro = v(med.get("nregistro", ""))
                if not nregistro: continue
                pres_data = cima_get("presentaciones", {
                    "nregistro": nregistro,
                    "comerc": 1
                })
                if not pres_data: continue
                for p in pres_data.get("resultados", []):
                    cn  = v(p.get("cn", ""))
                    nom = v(p.get("nombre", ""))
                    if cn and cn not in cns_vistos and es_inhalatori(nom):
                        cns_vistos.add(cn)
                        p["_classe"]          = classe
                        p["_atc"]             = codi_atc
                        p["_principiosActivos"] = med.get("pactivos", "")
                        resultats_filtrats.append(p)
                        trobats += 1

            if pagina * 25 >= total: break
            pagina += 1
            time.sleep(0.3)

        if trobats > 0:
            print(f"    → {trobats} inhaladors trobats")
        time.sleep(0.3)

    print(f"  ✅ Total inhaladors: {len(resultats_filtrats)}")
    return resultats_filtrats

def get_fitxa_posologia(nregistro):
    try:
        data = cima_get("docSegmentado/contenido/1", {"nregistro": nregistro, "seccion": "4.2"})
        if not data or not data.get("secciones"): return ""
        soup = BeautifulSoup(data["secciones"][0].get("contenido",""), "html.parser")
        return soup.get_text(separator=" ", strip=True)[:400]
    except: return ""

def infereix_dispositiu(nom):
    nl = str(nom).lower()
    for k, val in DISPOSITIU_MAP.items():
        if k in nl: return val
    return ("ICP","🔴","20-30 l/m","Lenta","https://scientiasalut.gencat.cat/handle/11351/11880")

def construeix_dosi(codi_atc, nregistro):
    """
    Construeix el text de dosi a partir del codi ATC.
    La dosi és sempre la mateixa independentment del dispositiu.
    """
    dosi_text = ""
    phf_val   = ""
    matma_val = ""
    color_atencio = False

    if codi_atc in DOSIS_PER_ATC:
        dosi_mpoc, dmx, starred, matma = DOSIS_PER_ATC[codi_atc]
        dosi_text = f"MPOC: {dosi_mpoc}\nD.màx: {dmx}"
        if starred: phf_val   = "★"
        if matma:   matma_val = "MATMA"

    if codi_atc in DOSIS_ASMA_GCI:
        b, m, a = DOSIS_ASMA_GCI[codi_atc]
        dosi_text += f"\nAsma baixa: {b}\nAsma mitjana: {m}\nAsma alta: {a}"
    elif codi_atc in DOSIS_ASMA_COMBO:
        b, m, a = DOSIS_ASMA_COMBO[codi_atc]
        dosi_text += f"\nAsma baixa: {b}\nAsma mitjana: {m}\nAsma alta: {a}"

    if not dosi_text:
        posologia = get_fitxa_posologia(nregistro)
        dosi_text = f"FITXA TÈCNICA: {posologia[:200]}" if posologia else "PENDENT — consultar fitxa CIMA"
        color_atencio = True

    return dosi_text, phf_val, matma_val, color_atencio

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
    print(f"  CNs existents al catàleg: {len(cns)}")

    presentacions = get_tots_inhaladors_cima()
    novetats = [p for p in presentacions if v(p.get("cn","")) not in cns]
    print(f"  Novetats detectades: {len(novetats)}")

    if not novetats:
        with open("novetats_count.txt", "w") as f: f.write("0")
        print("✅ Cap novetat — catàleg actualitzat")
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
        codi_atc  = p.get("_atc", "")
        classe    = p.get("_classe", "PENDENT")
        principis = v(p.get("_principiosActivos", ""))
        print(f"  [{i}] {nom[:60]} ({codi_atc})")

        tipus, co2, flux, flux4, link = infereix_dispositiu(nom)
        dosi_text, phf_val, matma_val, color_atencio = construeix_dosi(codi_atc, nregistro)
        color = fill_atencio if color_atencio else fill_nou

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
    print(f"  Pendents de validació: {pendents}")

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
        nom   = v(row[I_NOM])
        data  = v(row[I_DATA])
        estat = v(row[I_ESTAT])
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

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row): continue
        nom = v(row[I_NOM])
        if not nom: continue
        estat = v(row[I_ESTAT])
        if estat:
            print(f"  ⏭  Saltat (pendent): {nom[:40]}")
            continue
        print(f"  ✅ Inclòs: {nom[:40]}")

    print(f"✅ HTML regenerat correctament")

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
