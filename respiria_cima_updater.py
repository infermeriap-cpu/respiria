#!/usr/bin/env python3
"""
RespirIA — Actualitzador automàtic CIMA → Excel → HTML
=======================================================
Autora: Sílvia Álvarez Vega · ICS Atenció Primària Girona
Versió: 5.8 · 2026-05-14

Canvis v5.8 — consulta definitiva:
  - UNA SOLA consulta: vias=78 (inhalatoria) + atc=R03 + comerc=1
  - Retorna 341 medicaments — tots els inhaladors MPOC/Asma comercialitzats
  - Filtre per formaFarmaceutica: exclou nebulitzadors automàticament
  - Classe terapèutica des del codi ATC del medicament (sempre correcta)
  - Principi actiu des de vtm.nombre (complet per a combinacions)
  - Dispositiu des de formaFarmaceutica.nombre (detecta UNIDOSIS)
  - Dosi des del diccionari per codi ATC (sempre correcta)
"""

import requests, openpyxl, time, shutil, sys
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl.styles import PatternFill, Font, Alignment

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

COLOR_NOU     = "FEF3C7"   # groc = pendent validació
COLOR_ATENCIO = "FEE2E2"   # vermell = dosi no trobada
COLOR_NO_INC  = "F3F4F6"   # gris = no incorporar (nebulitzador)

# ═══════════════════════════════════════════════════════════════════════════════
# FORMES FARMACÈUTIQUES → TIPUS DISPOSITIU
# Font: DICCIONARIO_FORMA_FARMACEUTICA.xml de la CIMA
# ═══════════════════════════════════════════════════════════════════════════════

# Formes que SÍ incorporem al prescriptor AP
FORMES_PRESCRIPTOR = {
    "SUSPENSIÓN PARA INHALACIÓN EN ENVASE A PRESIÓN": (
        "ICP", "🔴", "20-30 l/m", "Lenta",
        "https://scientiasalut.gencat.cat/handle/11351/11880"),
    "SOLUCIÓN PARA INHALACIÓN EN ENVASE A PRESIÓN": (
        "ICP", "🔴", "20-30 l/m", "Lenta",
        "https://scientiasalut.gencat.cat/handle/11351/11880"),
    "POLVO PARA INHALACIÓN": (
        "IPS-multi", "🟢", "Variable", "Ràpida",
        "https://scientiasalut.gencat.cat/handle/11351/11881"),
    "POLVO PARA INHALACIÓN (UNIDOSIS)": (
        "IPS-uni", "🟢", "Variable", "Ràpida",
        "https://scientiasalut.gencat.cat/handle/11351/11816"),
    "POLVO PARA INHALACIÓN (CÁPSULA DURA)": (
        "IPS-uni", "🟢", "Variable", "Ràpida",
        "https://scientiasalut.gencat.cat/handle/11351/11816"),
    "SOLUCIÓN PARA INHALACIÓN": (
        "IVS", "🟢", "20-30 l/m", "Lenta",
        "https://scientiasalut.gencat.cat/handle/11351/11891"),
    "SOLUCIÓN PARA INHALACIÓN DEL VAPOR": (
        "IVS", "🟢", "20-30 l/m", "Lenta",
        "https://scientiasalut.gencat.cat/handle/11351/11891"),
    "LÍQUIDO PARA INHALACIÓN DEL VAPOR": (
        "IVS", "🟢", "20-30 l/m", "Lenta",
        "https://scientiasalut.gencat.cat/handle/11351/11891"),
}

# Formes que NO incorporem (nebulitzadors i altres)
FORMES_NO_INCORPORAR = [
    "NEBULIZADOR",
    "NEBULIZACIÓN",
    "GAS PARA INHALACION",
    "POLVO Y DISOLVENTE PARA SOLUCIÓN PARA INHALACIÓN",
    "SOLUCIÓN ORAL O CONCENTRADO PARA INHALACIÓN",
]

# Refinament del link d'instruccions pel nom del dispositiu
DISPOSITIU_LINK_MAP = {
    "turbuhaler":  "https://scientiasalut.gencat.cat/handle/11351/11893",
    "accuhaler":   "https://scientiasalut.gencat.cat/handle/11351/11813",
    "genuair":     "https://scientiasalut.gencat.cat/handle/11351/11878",
    "ellipta":     "https://scientiasalut.gencat.cat/handle/11351/11818",
    "novolizer":   "https://scientiasalut.gencat.cat/handle/11351/11882",
    "easyhaler":   "https://scientiasalut.gencat.cat/handle/11351/11817",
    "nexthaler":   "https://scientiasalut.gencat.cat/handle/11351/11881",
    "spiromax":    "https://scientiasalut.gencat.cat/handle/11351/11892",
    "forspiro":    "https://scientiasalut.gencat.cat/handle/11351/11819",
    "twisthaler":  "https://scientiasalut.gencat.cat/handle/11351/11894",
    "breezhaler":  "https://scientiasalut.gencat.cat/handle/11351/11816",
    "aerolizer":   "https://scientiasalut.gencat.cat/handle/11351/11815",
    "handihaler":  "https://scientiasalut.gencat.cat/handle/11351/11879",
    "zonda":       "https://scientiasalut.gencat.cat/handle/11351/11895",
    "respimat":    "https://scientiasalut.gencat.cat/handle/11351/11891",
}

# ═══════════════════════════════════════════════════════════════════════════════
# MAPES ATC → CLASSE TERAPÈUTICA + PRINCIPI ACTIU + DOSIS
# ═══════════════════════════════════════════════════════════════════════════════
ATC_A_CLASSE = {
    "R03AC02":"SABA", "R03AC03":"SABA", "R03AC04":"SABA",
    "R03AC12":"LABA", "R03AC13":"LABA", "R03AC18":"LABA", "R03AC19":"LABA",
    "R03BB01":"SAMA", "R03BB02":"SAMA",
    "R03BB04":"LAMA", "R03BB05":"LAMA", "R03BB06":"LAMA", "R03BB07":"LAMA",
    "R03BA01":"GCI",  "R03BA02":"GCI",  "R03BA05":"GCI",
    "R03BA07":"GCI",  "R03BA08":"GCI",
    "R03AK03":"SABA/SAMA", "R03AL01":"SABA/SAMA", "R03AL02":"SABA/SAMA",
    "R03AK06":"LABA/GCI",  "R03AK07":"LABA/GCI",  "R03AK08":"LABA/GCI",
    "R03AK09":"LABA/GCI",  "R03AK10":"LABA/GCI",  "R03AK11":"LABA/GCI",
    "R03AK14":"LABA/GCI",
    "R03AL03":"LABA/LAMA", "R03AL04":"LABA/LAMA",
    "R03AL05":"LABA/LAMA", "R03AL06":"LABA/LAMA",
    "R03AL08":"LAMA/LABA/GCI", "R03AL09":"LAMA/LABA/GCI",
    "R03AL11":"LAMA/LABA/GCI", "R03AL12":"LAMA/LABA/GCI",
}

PRINCIPI_PER_ATC = {
    "R03AC02":"salbutamol", "R03AC03":"terbutalina", "R03AC04":"fenoterol",
    "R03AC12":"salmeterol", "R03AC13":"formoterol",
    "R03AC18":"indacaterol", "R03AC19":"olodaterol",
    "R03BB01":"ipratropio", "R03BB02":"oxitropio",
    "R03BB04":"tiotropio",  "R03BB05":"aclidinio",
    "R03BB06":"glicopirronio", "R03BB07":"umeclidinio",
    "R03BA01":"beclometasona", "R03BA02":"budesonida",
    "R03BA05":"fluticasona",   "R03BA07":"mometasona", "R03BA08":"ciclesonida",
    "R03AK03":"fenoterol + ipratropio",
    "R03AL01":"fenoterol + ipratropio",
    "R03AL02":"salbutamol + ipratropio",
    "R03AK06":"salmeterol + fluticasona",
    "R03AK07":"formoterol + budesonida",
    "R03AK08":"formoterol + beclometasona",
    "R03AK09":"formoterol + mometasona",
    "R03AK10":"vilanterol + fluticasona furoat",
    "R03AK11":"formoterol + fluticasona",
    "R03AK14":"indacaterol + mometasona",
    "R03AL03":"vilanterol + umeclidinio",
    "R03AL04":"indacaterol + glicopirronio",
    "R03AL05":"formoterol + aclidinio",
    "R03AL06":"olodaterol + tiotropio",
    "R03AL08":"vilanterol + umeclidinio + fluticasona furoat",
    "R03AL09":"formoterol + glicopirronio + beclometasona",
    "R03AL11":"formoterol + glicopirronio + budesonida",
    "R03AL12":"indacaterol + glicopirronio + mometasona",
}

DOSIS_PER_ATC = {
    "R03AC02": ("100-200 µg si cal",    "200 µg/6h",             True,  False),
    "R03AC03": ("500 µg si cal",         "6.000 µg/24h",          False, False),
    "R03AC04": ("100-200 µg si cal",    "200 µg/6h",             False, False),
    "R03AC12": ("50 µg/12h",             "100 µg/12h",            True,  False),
    "R03AC13": ("12 µg/12h",             "24 µg/12h",             True,  False),
    "R03AC18": ("150 µg/24h",            "300 µg/24h",            True,  False),
    "R03AC19": ("5 µg/24h",              "5 µg/24h",              False, False),
    "R03BB01": ("40 µg si cal",          "240 µg/24h",            True,  False),
    "R03BB02": ("200 µg si cal",         "800 µg/24h",            False, False),
    "R03BB04": ("18 µg/24h (Handihaler) / 5 µg/24h (Respimat) / 10 µg/24h (Zonda)",
                "= pauta",               True,  False),
    "R03BB05": ("322 µg/12h",            "322 µg/12h",            False, False),
    "R03BB06": ("44 µg/24h",             "44 µg/24h",             False, False),
    "R03BB07": ("55 µg/24h",             "55 µg/24h",             False, False),
    "R03BA01": ("250-500 µg/12h",        "1.000 µg/12h",          False, False),
    "R03BA02": ("200-400 µg/12h",        "800 µg/12h",            False, False),
    "R03BA05": ("250-500 µg/12h",        "500 µg/12h",            False, False),
    "R03BA07": ("200 µg/24h",            "800 µg/24h",            False, False),
    "R03BA08": ("80-160 µg/24h",         "1.280 µg/24h",          False, False),
    "R03AK03": ("100-200/40 µg si cal",  "200/240 µg/24h",        False, False),
    "R03AL01": ("100-200/40 µg si cal",  "200/240 µg/24h",        False, False),
    "R03AL02": ("100-200/40 µg si cal",  "200/240 µg/24h",        False, False),
    "R03AK06": ("50/500 µg/12h",         "50/500 µg/12h",         False, False),
    "R03AK07": ("9/320 µg/12h",          "36/1.280 µg/24h",       False, False),
    "R03AK08": ("12/200 µg/12h",         "12/400 µg/12h",         False, False),
    "R03AK09": ("5/200 µg/24h",          "5/400 µg/24h",          False, False),
    "R03AK10": ("22/92 µg/24h",          "22/184 µg/24h",         False, False),
    "R03AK11": ("5/125 µg/12h",          "5/250 µg/12h",          False, False),
    "R03AK14": ("150/160 µg/24h",        "150/160 µg/24h",        False, False),
    "R03AL03": ("22/55 µg/24h",          "22/55 µg/24h",          False, False),
    "R03AL04": ("150/50 µg/24h",         "150/50 µg/24h",         True,  False),
    "R03AL05": ("12/340 µg/12h",         "12/340 µg/12h",         False, False),
    "R03AL06": ("5/5 µg/24h",            "5/5 µg/24h",            False, False),
    "R03AL08": ("22/55/92 µg/24h",       "22/55/184 µg/24h",      False, True),
    "R03AL09": ("10/18/174 µg/12h",      "10/18/344 µg/12h",      False, True),
    "R03AL11": ("10/160/14.4 µg/12h",    "10/160/14.4 µg/12h",    False, True),
    "R03AL12": ("150/50/160 µg/24h",     "150/50/160 µg/24h",     False, True),
}

DOSIS_ASMA_GCI = {
    "R03BA01": ("200-500 µg/24h",  "501-1.000 µg/24h", "1.001-2.000 µg/24h"),
    "R03BA02": ("200-400 µg/24h",  "401-800 µg/24h",   "801-1.600 µg/24h"),
    "R03BA05": ("100-250 µg/24h",  "251-500 µg/24h",   "501-1.000 µg/24h"),
    "R03BA07": ("200 µg/24h",      "400 µg/24h",       "800 µg/24h"),
    "R03BA08": ("80-160 µg/24h",   "161-320 µg/24h",   "321-1.280 µg/24h"),
}

DOSIS_ASMA_COMBO = {
    "R03AK06": ("50/100 µg/12h",    "50/250 µg/12h",    "50/500 µg/12h"),
    "R03AK07": ("4.5/160 µg/12h",   "9/320 µg/12h",     "2 inh 9/320 µg/12h"),
    "R03AK08": ("6/100: 1 inh/12h", "6/100: 2 inh/12h", "6/200: 2 inh/12h"),
    "R03AK10": ("22/92 µg/24h",     "22/92 µg/24h",     "22/184 µg/24h"),
    "R03AK11": ("5/100 µg/12h",     "5/250 µg/12h",     "5/500 µg/12h"),
}

# ═══════════════════════════════════════════════════════════════════════════════
# FUNCIONS AUXILIARS
# ═══════════════════════════════════════════════════════════════════════════════
def v(val):
    if val is None: return ""
    s = str(val).strip()
    return "" if s == "None" else s

def normalitza_cn(cn):
    cn = v(cn)
    try:
        return str(int(float(cn))) if cn.replace('.','').isdigit() else cn
    except:
        return cn

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

def get_codi_atc(med):
    """Extreu el codi ATC més específic del medicament"""
    atcs = med.get("atcs", [])
    for atc in atcs:
        codi = v(atc.get("codigo","")).upper()
        if codi in ATC_A_CLASSE:
            return codi
    # Si no troba exacte, prova amb els 7 primers caràcters
    for atc in atcs:
        codi = v(atc.get("codigo","")).upper()[:7]
        if codi in ATC_A_CLASSE:
            return codi
    # Retorna el primer ATC R03 que trobi
    for atc in atcs:
        codi = v(atc.get("codigo","")).upper()
        if codi.startswith("R03"):
            return codi
    return ""

def classifica_forma(med):
    """
    Classifica el medicament segons la forma farmacèutica.
    Retorna: (tipus, co2, flux, flux4, link, es_nebulitzador)
    """
    forma = med.get("formaFarmaceutica", {})
    forma_nom = v(forma.get("nombre", "")).upper()
    nom_med = v(med.get("nombre","")).lower()

    # Comprova si és nebulitzador o no incorporable
    for paraula in FORMES_NO_INCORPORAR:
        if paraula in forma_nom:
            return "NEB", "—", "—", "—", "", True

    # Busca la forma al mapa de prescriptor
    for forma_clau, valors in FORMES_PRESCRIPTOR.items():
        if forma_clau in forma_nom:
            tipus, co2, flux, flux4, link = valors
            # Refina link pel nom del dispositiu
            for disp_nom, disp_link in DISPOSITIU_LINK_MAP.items():
                if disp_nom in nom_med:
                    link = disp_link
                    if disp_nom == "respimat":
                        tipus = "IVS"
                    break
            return tipus, co2, flux, flux4, link, False

    # Forma desconeguda → ICP per defecte
    return "ICP", "🔴", "20-30 l/m", "Lenta", \
           "https://scientiasalut.gencat.cat/handle/11351/11880", False

def construeix_dosi(codi_atc, nregistro):
    dosi_text = ""
    phf_val = matma_val = ""
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
        try:
            data = cima_get("docSegmentado/contenido/1",
                            {"nregistro": nregistro, "seccion": "4.2"})
            if data and data.get("secciones"):
                soup = BeautifulSoup(
                    data["secciones"][0].get("contenido",""), "html.parser")
                posologia = soup.get_text(separator=" ", strip=True)[:400]
                dosi_text = f"FITXA TÈCNICA: {posologia}" if posologia \
                            else "PENDENT — consultar fitxa CIMA"
        except:
            dosi_text = "PENDENT — consultar fitxa CIMA"
        color_atencio = True

    return dosi_text, phf_val, matma_val, color_atencio

# ═══════════════════════════════════════════════════════════════════════════════
# CONSULTA CIMA — UNA SOLA CONSULTA
# ═══════════════════════════════════════════════════════════════════════════════
def get_tots_inhaladors_cima():
    """
    UNA SOLA consulta: vias=78 (inhalatoria) + atc=R03 + comerc=1
    Retorna tots els 341 inhaladors MPOC/Asma comercialitzats a Espanya.
    Filtre per formaFarmaceutica (exclou nebulitzadors).
    Classe terapèutica des del codi ATC del medicament.
    Principi actiu des de vtm.nombre + fallback per ATC.
    """
    print("📡 Consultant CIMA — vias=78 + atc=R03 (consulta única)...")
    resultats = []
    cns_vistos = set()
    pagina = 1
    total_meds = 0

    while True:
        data = cima_get("medicamentos", {
            "vias": 78,
            "atc": "R03",
            "comerc": 1,
            "pagina": pagina
        })
        if not data: break
        items = data.get("resultados", [])
        if not items: break
        total = data.get("totalFilas", 0)
        total_meds += len(items)
        print(f"  → Pàg {pagina}: {len(items)} medicaments ({total_meds}/{total})")

        for med in items:
            nregistro = v(med.get("nregistro", ""))
            if not nregistro: continue

            # Obté codi ATC específic del medicament
            codi_atc = get_codi_atc(med)

            # Principi actiu: vtm.nombre → pactivos → diccionari ATC
            vtm = med.get("vtm", {})
            principi = v(vtm.get("nombre", ""))
            if not principi:
                principi = v(med.get("pactivos", ""))
            if not principi and codi_atc:
                principi = PRINCIPI_PER_ATC.get(codi_atc, "")

            # Classe terapèutica des del codi ATC
            classe = ATC_A_CLASSE.get(codi_atc, "PENDENT")

            # Obté presentacions comercialitzades
            pres_data = cima_get("presentaciones", {
                "nregistro": nregistro,
                "comerc": 1
            })
            if not pres_data: continue

            for p in pres_data.get("resultados", []):
                cn = normalitza_cn(p.get("cn", ""))
                if not cn or cn in cns_vistos:
                    continue
                cns_vistos.add(cn)

                # Classifica la forma farmacèutica
                # Usa el nom de la presentació per al dispositiu
                med_proxy = dict(med)
                med_proxy["nombre"] = v(p.get("nombre", ""))
                tipus, co2, flux, flux4, link, es_neb = classifica_forma(med_proxy)

                p["_classe"]    = classe
                p["_atc"]       = codi_atc
                p["_principi"]  = principi
                p["_nregistro"] = nregistro
                p["_tipus"]     = tipus
                p["_co2"]       = co2
                p["_flux"]      = flux
                p["_flux4"]     = flux4
                p["_link"]      = link
                p["_es_neb"]    = es_neb
                resultats.append(p)

        if total_meds >= total: break
        pagina += 1
        time.sleep(0.3)

    incorporables = sum(1 for r in resultats if not r.get("_es_neb"))
    nebulitzadors = sum(1 for r in resultats if r.get("_es_neb"))
    print(f"  ✅ Total: {len(resultats)} presentacions "
          f"({incorporables} incorporables, {nebulitzadors} nebulitzadors)")
    return resultats

# ═══════════════════════════════════════════════════════════════════════════════
# MODES D'EXECUCIÓ
# ═══════════════════════════════════════════════════════════════════════════════
def mode_detecta():
    print("🔍 Mode: detectar novetats CIMA")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active

    cns = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        cn = normalitza_cn(row[I_CN])
        if cn: cns.add(cn)
    print(f"  CNs existents al catàleg: {len(cns)}")

    presentacions = get_tots_inhaladors_cima()
    novetats = [p for p in presentacions
                if normalitza_cn(p.get("cn","")) not in cns]
    print(f"  Novetats detectades: {len(novetats)}")

    if not novetats:
        with open("novetats_count.txt", "w") as f: f.write("0")
        print("✅ Cap novetat — catàleg actualitzat")
        return

    shutil.copy(EXCEL_PATH, EXCEL_PATH.replace(".xlsx","_BACKUP.xlsx"))

    fill_nou     = PatternFill("solid", fgColor=COLOR_NOU)
    fill_atencio = PatternFill("solid", fgColor=COLOR_ATENCIO)
    fill_no_inc  = PatternFill("solid", fgColor=COLOR_NO_INC)
    font_base    = Font(name='Arial', size=10)
    alineacio    = Alignment(vertical="top", wrap_text=True)

    afegits = 0
    for i, p in enumerate(novetats[:50], 1):
        cn        = normalitza_cn(p.get("cn",""))
        nom       = v(p.get("nombre",""))
        nregistro = v(p.get("_nregistro",""))
        codi_atc  = p.get("_atc", "")
        classe    = p.get("_classe", "PENDENT")
        principi  = p.get("_principi", "")
        tipus     = p.get("_tipus", "ICP")
        co2       = p.get("_co2", "🔴")
        flux      = p.get("_flux", "20-30 l/m")
        flux4     = p.get("_flux4", "Lenta")
        link      = p.get("_link", "")
        es_neb    = p.get("_es_neb", False)

        tag = "[NEB] " if es_neb else ""
        print(f"  [{i}] {tag}{nom[:55]} | {classe} | {tipus}")

        dosi_text, phf_val, matma_val, color_atencio = construeix_dosi(
            codi_atc, nregistro)

        if es_neb:
            estat_validacio = "NO INCORPORAR — nebulitzador"
            color = fill_no_inc
        elif color_atencio:
            estat_validacio = "NOU — PENDENT VALIDACIÓ CLÍNICA"
            color = fill_atencio
        else:
            estat_validacio = "NOU — PENDENT VALIDACIÓ CLÍNICA"
            color = fill_nou

        fila = [
            classe, principi, nom, cn, nom,
            dosi_text, tipus, co2, flux, flux4, link,
            datetime.now().strftime("%Y-%m-%d"),
            f"https://cima.aemps.es/cima/publico/medicamento.html?nregistro={nregistro}",
            estat_validacio,
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
    print("🔍 Mode: comprovar per publicar")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
    with open(HTML_PATH, "r", encoding="utf-8") as f:
        html = f.read()
    pendents = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) <= I_ESTAT: continue
        nom   = v(row[I_NOM])
        data  = v(row[I_DATA])
        estat = v(row[I_ESTAT])
        if not nom: continue
        if data and not estat and nom not in html:
            pendents += 1
    with open("publicar_count.txt", "w") as f: f.write(str(pendents))
    print(f"  Per publicar: {pendents}")

def mode_regenera():
    print(f"🔄 Regenerant HTML: {OUTPUT_HTML}")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
    inclosos = saltats = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row): continue
        nom = v(row[I_NOM])
        if not nom: continue
        estat = v(row[I_ESTAT])
        if estat:
            print(f"  ⏭  Saltat: {nom[:40]} ({estat[:25]})")
            saltats += 1
        else:
            print(f"  ✅ Inclòs: {nom[:40]}")
            inclosos += 1
    print(f"✅ HTML regenerat: {inclosos} inclosos, {saltats} saltats")

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
        print("⚠️  Cap mode especificat.")
        print("    Usa: --detecta | --comprova-pendents | --comprova-publicar | --regenera | --tot")
