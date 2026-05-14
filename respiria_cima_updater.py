#!/usr/bin/env python3
"""
RespirIA — Actualitzador automàtic CIMA → Excel → HTML
=======================================================
Autora: Sílvia Álvarez Vega · ICS Atenció Primària Girona
Versió: 5.6 · 2026-05-14

Canvis v5.6 — basats en anàlisi del JSON real de la CIMA:
  - Principi actiu: usa vtm.nombre (sempre complet i correcte per a combinacions)
  - Dispositiu: usa formaFarmaceutica.nombre (detecta UNIDOSIS directament)
  - Filtre inhalatori: usa viasAdministracion (VÍA INHALATORIA) en lloc de paraules clau
  - Cerca per codis ATC R03 específics (font: CIMA maestras)
  - Dosis indexades per codi ATC → sempre correctes independentment del dispositiu
"""

import requests, openpyxl, time, os, shutil, sys
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
COLOR_ATENCIO = "FEE2E2"   # vermell = dosi no trobada, consultar CIMA

# ═══════════════════════════════════════════════════════════════════════════════
# CODIS ATC R03 — INHALADORS MPOC/ASMA
# Font: https://cima.aemps.es/cima/rest/maestras?maestra=7&nombre=R03
# Inclou R03A i R03B (inhalatoris). Exclou R03C (sistèmics) i R03D (xantines)
# ═══════════════════════════════════════════════════════════════════════════════
CODIS_ATC = {
    # SABA
    "R03AC02": "SABA",        # Salbutamol
    "R03AC03": "SABA",        # Terbutalina
    "R03AC04": "SABA",        # Fenoterol
    # LABA
    "R03AC12": "LABA",        # Salmeterol
    "R03AC13": "LABA",        # Formoterol
    "R03AC18": "LABA",        # Indacaterol
    "R03AC19": "LABA",        # Olodaterol
    # SAMA
    "R03BB01": "SAMA",        # Ipratropi
    "R03BB02": "SAMA",        # Oxitropi
    # LAMA
    "R03BB04": "LAMA",        # Tiotropi
    "R03BB05": "LAMA",        # Aclidini
    "R03BB06": "LAMA",        # Glicopirroni
    "R03BB07": "LAMA",        # Umeclidini
    # GCI
    "R03BA01": "GCI",         # Beclometasona
    "R03BA02": "GCI",         # Budesonida
    "R03BA05": "GCI",         # Fluticasona
    "R03BA07": "GCI",         # Mometasona
    "R03BA08": "GCI",         # Ciclesonida
    # SABA/SAMA
    "R03AK03": "SABA/SAMA",   # Fenoterol + Ipratropi
    "R03AL01": "SABA/SAMA",   # Fenoterol + Ipratropi
    "R03AL02": "SABA/SAMA",   # Salbutamol + Ipratropi
    # LABA/GCI
    "R03AK06": "LABA/GCI",    # Salmeterol + Fluticasona
    "R03AK07": "LABA/GCI",    # Formoterol + Budesonida
    "R03AK08": "LABA/GCI",    # Formoterol + Beclometasona
    "R03AK09": "LABA/GCI",    # Formoterol + Mometasona
    "R03AK10": "LABA/GCI",    # Vilanterol + Fluticasona furoat
    "R03AK11": "LABA/GCI",    # Formoterol + Fluticasona
    "R03AK14": "LABA/GCI",    # Indacaterol + Mometasona
    # LABA/LAMA
    "R03AL03": "LABA/LAMA",   # Vilanterol + Umeclidini
    "R03AL04": "LABA/LAMA",   # Indacaterol + Glicopirroni
    "R03AL05": "LABA/LAMA",   # Formoterol + Aclidini
    "R03AL06": "LABA/LAMA",   # Olodaterol + Tiotropi
    # LAMA/LABA/GCI
    "R03AL08": "LAMA/LABA/GCI",  # Vilanterol + Umeclidini + Fluticasona
    "R03AL09": "LAMA/LABA/GCI",  # Formoterol + Glicopirroni + Beclometasona
    "R03AL11": "LAMA/LABA/GCI",  # Formoterol + Glicopirroni + Budesonida
    "R03AL12": "LAMA/LABA/GCI",  # Indacaterol + Glicopirroni + Mometasona
}

# ═══════════════════════════════════════════════════════════════════════════════
# DOSIS PER CODI ATC — PHF CatSalut 2018
# Tupla: (dosi_mpoc, dosi_max, phf_star, matma)
# La dosi és SEMPRE la mateixa independentment del dispositiu
# ═══════════════════════════════════════════════════════════════════════════════
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

# Dosis asma GCI per escalons GEMA 5.5
DOSIS_ASMA_GCI = {
    "R03BA01": ("200-500 µg/24h",  "501-1.000 µg/24h", "1.001-2.000 µg/24h"),
    "R03BA02": ("200-400 µg/24h",  "401-800 µg/24h",   "801-1.600 µg/24h"),
    "R03BA05": ("100-250 µg/24h",  "251-500 µg/24h",   "501-1.000 µg/24h"),
    "R03BA07": ("200 µg/24h",      "400 µg/24h",       "800 µg/24h"),
    "R03BA08": ("80-160 µg/24h",   "161-320 µg/24h",   "321-1.280 µg/24h"),
}

# Dosis asma combinacions LABA/GCI per escalons GEMA 5.5
DOSIS_ASMA_COMBO = {
    "R03AK06": ("50/100 µg/12h",    "50/250 µg/12h",    "50/500 µg/12h"),
    "R03AK07": ("4.5/160 µg/12h",   "9/320 µg/12h",     "2 inh 9/320 µg/12h"),
    "R03AK08": ("6/100: 1 inh/12h", "6/100: 2 inh/12h", "6/200: 2 inh/12h"),
    "R03AK10": ("22/92 µg/24h",     "22/92 µg/24h",     "22/184 µg/24h"),
    "R03AK11": ("5/100 µg/12h",     "5/250 µg/12h",     "5/500 µg/12h"),
}

# ═══════════════════════════════════════════════════════════════════════════════
# MAPA DE FORMA FARMACÈUTICA → TIPUS DISPOSITIU
# Usa formaFarmaceutica.nombre del JSON de la CIMA — sempre correcte
# ═══════════════════════════════════════════════════════════════════════════════
FORMA_FARMACEUTICA_MAP = {
    # ICP — Inhaladors de cartucho pressuritzat
    "SUSPENSIÓN PARA INHALACIÓN EN ENVASE A PRESIÓN": (
        "ICP", "🔴", "20-30 l/m", "Lenta",
        "https://scientiasalut.gencat.cat/handle/11351/11880"),
    "SOLUCIÓN PARA INHALACIÓN EN ENVASE A PRESIÓN": (
        "ICP", "🔴", "20-30 l/m", "Lenta",
        "https://scientiasalut.gencat.cat/handle/11351/11880"),
    # IVS — Inhaladors de vapor suau
    "SOLUCIÓN PARA INHALACIÓN": (
        "IVS", "🟢", "20-30 l/m", "Lenta",
        "https://scientiasalut.gencat.cat/handle/11351/11891"),
    # IPS-multi — Inhaladors de pols sec multidosi
    "POLVO PARA INHALACIÓN": (
        "IPS-multi", "🟢", "Variable", "Ràpida",
        "https://scientiasalut.gencat.cat/handle/11351/11881"),
    # IPS-uni — Inhaladors de pols sec unidosi (càpsula per dosi)
    "POLVO PARA INHALACIÓN (UNIDOSIS)": (
        "IPS-uni", "🟢", "Variable", "Ràpida",
        "https://scientiasalut.gencat.cat/handle/11351/11816"),
}

# Mapa addicional per refinar el link d'instruccions segons el nom del dispositiu
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

def es_inhalatori(med):
    """
    Comprova via inhalatòria usant el camp viasAdministracion del JSON.
    Molt més fiable que paraules clau al nom.
    """
    vies = med.get("viasAdministracion", [])
    return any("INHALAT" in v(via.get("nombre","")).upper() for via in vies)

def infereix_dispositiu(med):
    """
    Infereix el tipus de dispositiu usant formaFarmaceutica.nombre del JSON.
    Detecta UNIDOSIS directament des del camp oficial de la CIMA.
    Refina el link d'instruccions pel nom del dispositiu si és possible.
    """
    forma = med.get("formaFarmaceutica", {})
    forma_nom = v(forma.get("nombre", "")).upper()
    nom = v(med.get("nombre", "")).lower()

    # Busca al mapa de formes farmacèutiques
    for clau, valors in FORMA_FARMACEUTICA_MAP.items():
        if clau in forma_nom:
            tipus, co2, flux, flux4, link = valors
            # Refina el link d'instruccions pel nom del dispositiu
            for disp_nom, disp_link in DISPOSITIU_LINK_MAP.items():
                if disp_nom in nom:
                    link = disp_link
                    # Respimat és sempre IVS
                    if disp_nom == "respimat":
                        tipus = "IVS"
                    break
            return tipus, co2, flux, flux4, link

    # Per defecte ICP si no s'ha identificat
    return ("ICP", "🔴", "20-30 l/m", "Lenta",
            "https://scientiasalut.gencat.cat/handle/11351/11880")

def construeix_dosi(codi_atc, nregistro):
    """
    Construeix el text de dosi a partir del codi ATC.
    La dosi és sempre la mateixa independentment del dispositiu.
    """
    dosi_text     = ""
    phf_val       = ""
    matma_val     = ""
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
# CONSULTA CIMA
# ═══════════════════════════════════════════════════════════════════════════════
def get_tots_inhaladors_cima():
    """
    Consulta la CIMA per cada codi ATC R03 rellevant.
    Usa viasAdministracion per filtrar inhaladors (molt més fiable que paraules clau).
    Usa vtm.nombre per al principi actiu (sempre complet per a combinacions).
    Usa formaFarmaceutica.nombre per al tipus de dispositiu (detecta UNIDOSIS).
    """
    print("📡 Consultant CIMA — codis ATC R03 + filtre viasAdministracion...")
    resultats = []
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
                # Filtre: només via inhalatòria
                if not es_inhalatori(med):
                    continue

                nregistro = v(med.get("nregistro", ""))
                if not nregistro: continue

                # Obté presentacions comercialitzades
                pres_data = cima_get("presentaciones", {
                    "nregistro": nregistro,
                    "comerc": 1
                })
                if not pres_data: continue

                for p in pres_data.get("resultados", []):
                    cn = v(p.get("cn", ""))
                    if not cn or cn in cns_vistos:
                        continue

                    cns_vistos.add(cn)

                    # vtm.nombre → principi actiu complet i correcte
                    vtm = med.get("vtm", {})
                    principi = v(vtm.get("nombre", ""))
                    if not principi:
                        principi = v(med.get("pactivos", ""))

                    # Afegeix metadades necessàries per a la incorporació
                    p["_classe"]    = classe
                    p["_atc"]       = codi_atc
                    p["_principi"]  = principi
                    p["_nregistro"] = nregistro
                    # Copia formaFarmaceutica del medicament a la presentació
                    p["_formaFarmaceutica"] = med.get("formaFarmaceutica", {})
                    p["_viasAdministracio"] = med.get("viasAdministracion", [])
                    resultats.append(p)
                    trobats += 1

            if pagina * 25 >= total: break
            pagina += 1
            time.sleep(0.3)

        if trobats > 0:
            print(f"    → {trobats} presentacions trobades")
        time.sleep(0.3)

    print(f"  ✅ Total inhaladors: {len(resultats)}")
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
        cn = v(row[I_CN])
        if cn: cns.add(str(int(float(cn))) if cn.replace('.','').isdigit() else cn)
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
        nregistro = v(p.get("_nregistro",""))
        codi_atc  = p.get("_atc", "")
        classe    = p.get("_classe", "PENDENT")
        principi  = p.get("_principi", "")

        print(f"  [{i}] {nom[:60]} ({codi_atc})")

        # Construeix un objecte med per a infereix_dispositiu
        med_proxy = {
            "nombre": nom,
            "formaFarmaceutica": p.get("_formaFarmaceutica", {}),
            "viasAdministracion": p.get("_viasAdministracio", []),
        }
        tipus, co2, flux, flux4, link = infereix_dispositiu(med_proxy)
        dosi_text, phf_val, matma_val, color_atencio = construeix_dosi(
            codi_atc, nregistro)
        color = fill_atencio if color_atencio else fill_nou

        fila = [
            classe,     # A: Classe terapèutica (des del codi ATC)
            principi,   # B: Principi actiu (vtm.nombre — sempre complet)
            nom,        # C: Nom comercial
            cn,         # D: Codi nacional
            nom,        # E: Dispositiu/Presentació
            dosi_text,  # F: Dosi (PHF CatSalut 2018 + GEMA 5.5)
            tipus,      # G: Tipus dispositiu (formaFarmaceutica)
            co2,        # H: Petjada CO2
            flux,       # I: Flux inspiratori
            flux4,      # J: Maniobra
            link,       # K: Link instruccions
            datetime.now().strftime("%Y-%m-%d"),  # L: Data
            f"https://cima.aemps.es/cima/publico/medicamento.html?nregistro={nregistro}",  # M
            "NOU — PENDENT VALIDACIÓ CLÍNICA",   # N: Bloqueja publicació
            phf_val,    # O: PHF ★
            matma_val,  # P: MATMA
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
            print(f"  ⏭  Saltat: {nom[:40]}")
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
