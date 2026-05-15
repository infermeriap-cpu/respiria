#!/usr/bin/env python3
"""
RespirIA — Actualitzador automàtic CIMA → Excel → HTML
=======================================================
Autora: Sílvia Álvarez Vega · ICS Atenció Primària Girona
Versió: 5.9.1 · 2026-05-15

Canvis v5.9.1:
  - Separació de múltiples CNs a la mateixa cel·la (ex: "700582 ,710249")
  - Evita falsos duplicats de fàrmacs com Formodual

Canvis v5.9:
  - Classe i dosi inferides des de vtm.nombre (sempre correcte)
  - Flux inspiratori específic per dispositiu (Turbuhaler 50-60, Accuhaler 60-90...)
  - Nebulitzadors descartats automàticament
  - Consulta única vias=78 + atc=R03 + comerc=1
"""

import requests, openpyxl, time, shutil, sys
from datetime import datetime
from openpyxl.styles import PatternFill, Font, Alignment

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓ
# ═══════════════════════════════════════════════════════════════════════════════
EXCEL_PATH  = "inhaladors_MPOC_ASMA_Avanzado_FINAL.xlsx"
HTML_PATH   = "RespirIA_v2.html"
CIMA_BASE   = "https://cima.aemps.es/cima/rest"

I_CLASSE = 0;  I_IA = 1;  I_NOM = 2;   I_CN = 3
I_DISP   = 4;  I_DOSI = 5; I_TIPUS = 6; I_CO2 = 7
I_FLUX   = 8;  I_FLUX4 = 9; I_LINK = 10
I_DATA   = 11; I_ORIGEN = 12; I_ESTAT = 13
I_PHF    = 14; I_MATMA = 15

COLOR_NOU     = "FEF3C7"
COLOR_ATENCIO = "FEE2E2"
COLOR_NO_INC  = "F3F4F6"

# ═══════════════════════════════════════════════════════════════════════════════
# MAPA VTM.NOMBRE → (classe, dosi_mpoc, dosi_max, phf_star, matma)
# ═══════════════════════════════════════════════════════════════════════════════
VTM_MAP = {
    # SABA
    "salbutamol":           ("SABA", "100-200 µg si cal",  "200 µg/6h",               True,  False),
    "terbutalina":          ("SABA", "500 µg si cal",       "6.000 µg/24h",            False, False),
    "fenoterol":            ("SABA", "100-200 µg si cal",  "200 µg/6h",               False, False),
    # SAMA
    "bromuro ipratropio":   ("SAMA", "40 µg si cal",        "240 µg/24h",              True,  False),
    "ipratropio":           ("SAMA", "40 µg si cal",        "240 µg/24h",              True,  False),
    "oxitropio":            ("SAMA", "200 µg si cal",       "800 µg/24h",              False, False),
    # LABA
    "salmeterol":           ("LABA", "50 µg/12h",           "100 µg/12h",              True,  False),
    "formoterol":           ("LABA", "12 µg/12h",           "24 µg/12h",               True,  False),
    "indacaterol":          ("LABA", "150 µg/24h",          "300 µg/24h",              True,  False),
    "olodaterol":           ("LABA", "5 µg/24h",            "5 µg/24h",                False, False),
    # LAMA
    "tiotropio":            ("LAMA", "18 µg/24h (Handihaler) / 5 µg/24h (Respimat) / 10 µg/24h (Zonda)", "= pauta", True, False),
    "aclidinio":            ("LAMA", "322 µg/12h",          "322 µg/12h",              False, False),
    "bromuro aclidinio":    ("LAMA", "322 µg/12h",          "322 µg/12h",              False, False),
    "glicopirronio":        ("LAMA", "44 µg/24h",           "44 µg/24h",               False, False),
    "bromuro umeclidinio":  ("LAMA", "55 µg/24h",           "55 µg/24h",               False, False),
    "umeclidinio":          ("LAMA", "55 µg/24h",           "55 µg/24h",               False, False),
    # GCI
    "beclometasona":        ("GCI",  "250-500 µg/12h",      "1.000 µg/12h",            False, False),
    "budesonida":           ("GCI",  "200-400 µg/12h",      "800 µg/12h",              False, False),
    "fluticasona":          ("GCI",  "250-500 µg/12h",      "500 µg/12h",              False, False),
    "mometasona":           ("GCI",  "200 µg/24h",          "800 µg/24h",              False, False),
    "ciclesonida":          ("GCI",  "80-160 µg/24h",       "1.280 µg/24h",            False, False),
    # SABA/SAMA
    "fenoterol + bromuro de ipratropio": ("SABA/SAMA", "100-200/40 µg si cal", "200/240 µg/24h", False, False),
    "salbutamol + bromuro ipratropio":   ("SABA/SAMA", "100-200/40 µg si cal", "200/240 µg/24h", False, False),
    # LABA/GCI
    "salmeterol + fluticasona":          ("LABA/GCI", "50/500 µg/12h",         "50/500 µg/12h",  False, False),
    "fluticasona + vilanterol":          ("LABA/GCI", "22/92 µg/24h",          "22/184 µg/24h",  False, False),
    "fluticasona + formoterol":          ("LABA/GCI", "5/125 µg/12h",          "5/250 µg/12h",   False, False),
    "budesonida + formoterol":           ("LABA/GCI", "9/320 µg/12h",          "36/1.280 µg/24h",False, False),
    "beclometasona + formoterol":        ("LABA/GCI", "12/200 µg/12h",         "12/400 µg/12h",  False, False),
    "indacaterol + mometasona":          ("LABA/GCI", "150/160 µg/24h",        "150/160 µg/24h", False, False),
    "formoterol + mometasona":           ("LABA/GCI", "5/200 µg/24h",          "5/400 µg/24h",   False, False),
    # LABA/LAMA
    "umeclidinio + vilanterol":          ("LABA/LAMA", "22/55 µg/24h",         "22/55 µg/24h",   False, False),
    "indacaterol + glicopirronio":       ("LABA/LAMA", "150/50 µg/24h",        "150/50 µg/24h",  True,  False),
    "aclidinio + formoterol":            ("LABA/LAMA", "12/340 µg/12h",        "12/340 µg/12h",  False, False),
    "olodaterol + tiotropio":            ("LABA/LAMA", "5/5 µg/24h",           "5/5 µg/24h",     False, False),
    # LAMA/LABA/GCI
    "beclometasona + formoterol + glicopirronio": ("LAMA/LABA/GCI", "10/18/87 µg/12h",   "10/18/174 µg/12h",  False, True),
    "fluticasona + umeclidinio + vilanterol":      ("LAMA/LABA/GCI", "22/55/92 µg/24h",   "22/55/184 µg/24h",  False, True),
    "formoterol + glicopirronio + budesonida":     ("LAMA/LABA/GCI", "5/7.2/160 µg/12h",  "5/7.2/160 µg/12h",  False, True),
    "indacaterol + glicopirronio + mometasona":    ("LAMA/LABA/GCI", "150/50/160 µg/24h", "150/50/160 µg/24h", False, True),
}

DOSIS_ASMA_GCI_VTM = {
    "beclometasona": ("200-500 µg/24h",  "501-1.000 µg/24h", "1.001-2.000 µg/24h"),
    "budesonida":    ("200-400 µg/24h",  "401-800 µg/24h",   "801-1.600 µg/24h"),
    "fluticasona":   ("100-250 µg/24h",  "251-500 µg/24h",   "501-1.000 µg/24h"),
    "mometasona":    ("200 µg/24h",      "400 µg/24h",       "800 µg/24h"),
    "ciclesonida":   ("80-160 µg/24h",   "161-320 µg/24h",   "321-1.280 µg/24h"),
}

DOSIS_ASMA_COMBO_VTM = {
    "salmeterol + fluticasona":   ("50/100 µg/12h",    "50/250 µg/12h",    "50/500 µg/12h"),
    "budesonida + formoterol":    ("4.5/160 µg/12h",   "9/320 µg/12h",     "2 inh 9/320 µg/12h"),
    "beclometasona + formoterol": ("6/100: 1 inh/12h", "6/100: 2 inh/12h", "6/200: 2 inh/12h"),
    "fluticasona + vilanterol":   ("22/92 µg/24h",     "22/92 µg/24h",     "22/184 µg/24h"),
    "fluticasona + formoterol":   ("5/100 µg/12h",     "5/250 µg/12h",     "5/500 µg/12h"),
}

# ═══════════════════════════════════════════════════════════════════════════════
# FORMES FARMACÈUTIQUES
# ═══════════════════════════════════════════════════════════════════════════════
FORMES_OK = {
    "SUSPENSIÓN PARA INHALACIÓN EN ENVASE A PRESIÓN",
    "SOLUCIÓN PARA INHALACIÓN EN ENVASE A PRESIÓN",
    "POLVO PARA INHALACIÓN",
    "POLVO PARA INHALACIÓN (UNIDOSIS)",
    "POLVO PARA INHALACIÓN (CÁPSULA DURA)",
    "SOLUCIÓN PARA INHALACIÓN",
    "SOLUCIÓN PARA INHALACIÓN DEL VAPOR",
    "LÍQUIDO PARA INHALACIÓN DEL VAPOR",
}

FORMES_NO_INC = {
    "SOLUCIÓN PARA INHALACIÓN POR NEBULIZADOR",
    "SUSPENSIÓN PARA INHALACIÓN POR NEBULIZADOR",
    "POLVO PARA SOLUCIÓN PARA INHALACIÓN POR NEBULIZADOR",
    "POLVO Y DISOLVENTE PARA SOLUCIÓN PARA INHALACIÓN POR NEBULIZADOR",
    "SOLUCIÓN ORAL O CONCENTRADO PARA INHALACIÓN POR NEBULIZADOR",
}

# Dispositius específics amb flux inspiratori correcte per a cada un
DISPOSITIU_MAP = {
    "turbuhaler":  ("IPS-multi","🟢","50-60 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11893"),
    "accuhaler":   ("IPS-multi","🟢","60-90 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11813"),
    "genuair":     ("IPS-multi","🟢","60-90 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11878"),
    "ellipta":     ("IPS-multi","🟢","<50 l/m",  "Ràpida","https://scientiasalut.gencat.cat/handle/11351/11818"),
    "novolizer":   ("IPS-multi","🟢","60-90 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11882"),
    "easyhaler":   ("IPS-multi","🟢","<50 l/m",  "Ràpida","https://scientiasalut.gencat.cat/handle/11351/11817"),
    "nexthaler":   ("IPS-multi","🟢","60-90 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11881"),
    "spiromax":    ("IPS-multi","🟢","60-90 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11892"),
    "forspiro":    ("IPS-multi","🟢","60-90 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11819"),
    "twisthaler":  ("IPS-multi","🟢","<50 l/m",  "Ràpida","https://scientiasalut.gencat.cat/handle/11351/11894"),
    "breezhaler":  ("IPS-uni", "🟢",">90 l/m",  "Ràpida","https://scientiasalut.gencat.cat/handle/11351/11816"),
    "aerolizer":   ("IPS-uni", "🟢",">90 l/m",  "Ràpida","https://scientiasalut.gencat.cat/handle/11351/11815"),
    "handihaler":  ("IPS-uni", "🟢","<50 l/m",  "Ràpida","https://scientiasalut.gencat.cat/handle/11351/11879"),
    "zonda":       ("IPS-uni", "🟢","<50 l/m",  "Ràpida","https://scientiasalut.gencat.cat/handle/11351/11895"),
    "respimat":    ("IVS",     "🟢","20-30 l/m","Lenta", "https://scientiasalut.gencat.cat/handle/11351/11891"),
    "modulite":    ("ICP",     "🔴","20-30 l/m","Lenta", "https://scientiasalut.gencat.cat/handle/11351/11880"),
    "aerosphere":  ("ICP",     "🔴","20-30 l/m","Lenta", "https://scientiasalut.gencat.cat/handle/11351/11880"),
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

def normalitza_vtm(vtm_nom):
    return v(vtm_nom).lower().strip()

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

def classifica_forma(forma_nom):
    """Classifica la forma farmacèutica i retorna el tipus base de dispositiu"""
    fn = forma_nom.upper().strip()
    if fn in FORMES_NO_INC:
        return "NEB","—","—","—","", False, True
    if fn in FORMES_OK:
        if "ENVASE A PRESIÓN" in fn or "ENVASE A PRESION" in fn:
            return "ICP","🔴","20-30 l/m","Lenta","https://scientiasalut.gencat.cat/handle/11351/11880", True, False
        elif "SOLUCIÓN PARA INHALACIÓN" in fn or "LÍQUIDO PARA INHALACIÓN" in fn:
            return "IVS","🟢","20-30 l/m","Lenta","https://scientiasalut.gencat.cat/handle/11351/11891", True, False
        elif "CÁPSULA DURA" in fn or "CAPSULA DURA" in fn or "UNIDOSIS" in fn:
            return "IPS-uni","🟢","<50 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11816", True, False
        else:
            return "IPS-multi","🟢","60-90 l/m","Ràpida","https://scientiasalut.gencat.cat/handle/11351/11881", True, False
    return "?","—","—","—","", False, False

def refina_dispositiu(nom_presentacio, tipus_base, co2_base, flux_base, flux4_base, link_base):
    """
    Refina el tipus, flux i link a partir del nom específic del dispositiu.
    Prioritat: nom dispositiu al nom de la presentació > tipus genèric de la forma.
    """
    nl = nom_presentacio.lower()
    for disp, vals in DISPOSITIU_MAP.items():
        if disp in nl:
            return vals  # (tipus, co2, flux, flux4, link)
    return tipus_base, co2_base, flux_base, flux4_base, link_base

def construeix_dosi_vtm(vtm_nom):
    """Construeix el text de dosi a partir del vtm.nombre"""
    vtm = normalitza_vtm(vtm_nom)
    dosi_text = ""
    phf_val = matma_val = ""
    color_atencio = False

    if vtm in VTM_MAP:
        classe, dosi_mpoc, dmx, starred, matma = VTM_MAP[vtm]
        dosi_text = f"MPOC: {dosi_mpoc}\nD.màx: {dmx}"
        if starred: phf_val   = "★"
        if matma:   matma_val = "MATMA"
        if vtm in DOSIS_ASMA_GCI_VTM:
            b, m, a = DOSIS_ASMA_GCI_VTM[vtm]
            dosi_text += f"\nAsma baixa: {b}\nAsma mitjana: {m}\nAsma alta: {a}"
        elif vtm in DOSIS_ASMA_COMBO_VTM:
            b, m, a = DOSIS_ASMA_COMBO_VTM[vtm]
            dosi_text += f"\nAsma baixa: {b}\nAsma mitjana: {m}\nAsma alta: {a}"
    else:
        dosi_text = "PENDENT — consultar fitxa CIMA"
        color_atencio = True

    return dosi_text, phf_val, matma_val, color_atencio

def get_classe_vtm(vtm_nom):
    vtm = normalitza_vtm(vtm_nom)
    if vtm in VTM_MAP:
        return VTM_MAP[vtm][0]
    return "PENDENT"

# ═══════════════════════════════════════════════════════════════════════════════
# CONSULTA CIMA
# ═══════════════════════════════════════════════════════════════════════════════
def get_tots_inhaladors_cima():
    """
    Consulta única: vias=78 + atc=R03 + comerc=1 → 341 medicaments.
    Classe i dosi des de vtm.nombre. Flux des del nom del dispositiu.
    """
    print("📡 Consultant CIMA — vias=78 + atc=R03...")
    resultats = []
    cns_vistos = set()
    pagina = 1
    total_consultat = 0

    while True:
        data = cima_get("medicamentos", {
            "vias": 78, "atc": "R03", "comerc": 1, "pagina": pagina
        })
        if not data: break
        items = data.get("resultados", [])
        if not items: break
        total = data.get("totalFilas", 0)
        total_consultat += len(items)
        print(f"  → Pàg {pagina}: {len(items)} ({total_consultat}/{total})")

        for med in items:
            nregistro = v(med.get("nregistro",""))
            if not nregistro: continue

            forma_nom = v(med.get("formaFarmaceutica", {}).get("nombre",""))
            vtm_nom   = v(med.get("vtm", {}).get("nombre",""))

            tipus_base, co2_base, flux_base, flux4_base, link_base, ok, es_neb = \
                classifica_forma(forma_nom)

            if not ok and not es_neb:
                continue

            pres_data = cima_get("presentaciones", {
                "nregistro": nregistro, "comerc": 1
            })
            if not pres_data: continue

            for p in pres_data.get("resultados", []):
                cn  = normalitza_cn(p.get("cn",""))
                nom = v(p.get("nombre",""))
                if not cn or cn in cns_vistos:
                    continue
                cns_vistos.add(cn)

                if ok and not es_neb:
                    tipus, co2, flux, flux4, link = refina_dispositiu(
                        nom, tipus_base, co2_base, flux_base, flux4_base, link_base)
                else:
                    tipus, co2, flux, flux4, link = "NEB","—","—","—",""

                p["_nregistro"] = nregistro
                p["_vtm_nom"]   = vtm_nom
                p["_classe"]    = get_classe_vtm(vtm_nom)
                p["_tipus"]     = tipus
                p["_co2"]       = co2
                p["_flux"]      = flux
                p["_flux4"]     = flux4
                p["_link"]      = link
                p["_es_neb"]    = es_neb
                resultats.append(p)

        if total_consultat >= total: break
        pagina += 1
        time.sleep(0.3)

    inc = sum(1 for r in resultats if not r.get("_es_neb"))
    neb = sum(1 for r in resultats if r.get("_es_neb"))
    print(f"  ✅ Total: {len(resultats)} ({inc} incorporables, {neb} nebulitzadors)")
    return resultats

# ═══════════════════════════════════════════════════════════════════════════════
# MODES D'EXECUCIÓ
# ═══════════════════════════════════════════════════════════════════════════════
def mode_detecta():
    print("🔍 Mode: detectar novetats CIMA")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active

    # Lectura CNs — separa múltiples CNs a la mateixa cel·la (ex: "700582 ,710249")
    cns = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        cel = v(row[I_CN])
        if not cel: continue
        for cn in cel.replace(';', ',').split(','):
            cn_net = normalitza_cn(cn.strip())
            if cn_net: cns.add(cn_net)
    print(f"  CNs existents: {len(cns)}")

    presentacions = get_tots_inhaladors_cima()
    novetats = [p for p in presentacions
                if normalitza_cn(p.get("cn","")) not in cns]
    print(f"  Novetats: {len(novetats)}")

    if not novetats:
        with open("novetats_count.txt","w") as f: f.write("0")
        print("✅ Cap novetat")
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
        vtm_nom   = p.get("_vtm_nom","")
        classe    = p.get("_classe","PENDENT")
        tipus     = p.get("_tipus","ICP")
        co2       = p.get("_co2","🔴")
        flux      = p.get("_flux","20-30 l/m")
        flux4     = p.get("_flux4","Lenta")
        link      = p.get("_link","")
        es_neb    = p.get("_es_neb", False)

        print(f"  [{i}] {'[NEB] ' if es_neb else ''}{nom[:55]} | {classe} | {tipus} | {flux}")

        dosi_text, phf_val, matma_val, color_atencio = construeix_dosi_vtm(vtm_nom)

        if es_neb:
            estat = "NO INCORPORAR — nebulitzador"
            color = fill_no_inc
            tipus = "NEB"
        elif color_atencio:
            estat = "NOU — PENDENT VALIDACIÓ CLÍNICA"
            color = fill_atencio
        else:
            estat = "NOU — PENDENT VALIDACIÓ CLÍNICA"
            color = fill_nou

        fila = [
            classe, vtm_nom, nom, cn, nom,
            dosi_text, tipus, co2, flux, flux4, link,
            datetime.now().strftime("%Y-%m-%d"),
            f"https://cima.aemps.es/cima/publico/medicamento.html?nregistro={nregistro}",
            estat, phf_val, matma_val,
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
    with open("novetats_count.txt","w") as f: f.write(str(afegits))
    print(f"✅ {afegits} novetats afegides")

def mode_comprova_pendents():
    print("🔍 Mode: comprovar pendents")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
    pendents = sum(1 for row in ws.iter_rows(min_row=2, values_only=True)
                   if len(row) > I_ESTAT and v(row[I_ESTAT]))
    with open("pendents_count.txt","w") as f: f.write(str(pendents))
    print(f"  Pendents: {pendents}")

def mode_comprova_publicar():
    print("🔍 Mode: comprovar per publicar")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
    with open(HTML_PATH,"r",encoding="utf-8") as f: html = f.read()
    pendents = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) <= I_ESTAT: continue
        nom = v(row[I_NOM]); data = v(row[I_DATA]); estat = v(row[I_ESTAT])
        if not nom: continue
        if data and not estat and nom not in html:
            pendents += 1
    with open("publicar_count.txt","w") as f: f.write(str(pendents))
    print(f"  Per publicar: {pendents}")

def mode_regenera():
    print(f"🔄 Regenerant HTML")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
    inclosos = saltats = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row): continue
        nom = v(row[I_NOM])
        if not nom: continue
        estat = v(row[I_ESTAT])
        if estat:
            print(f"  ⏭  {nom[:40]} ({estat[:25]})")
            saltats += 1
        else:
            print(f"  ✅ {nom[:40]}")
            inclosos += 1
    print(f"✅ {inclosos} inclosos, {saltats} saltats")

# ═══════════════════════════════════════════════════════════════════════════════
# ENTRADA PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    args = sys.argv[1:]
    if "--detecta" in args:              mode_detecta()
    elif "--comprova-pendents" in args:  mode_comprova_pendents()
    elif "--comprova-publicar" in args:  mode_comprova_publicar()
    elif "--regenera" in args:           mode_regenera()
    elif "--tot" in args:
        mode_detecta()
        mode_regenera()
    else:
        print("⚠️  Usa: --detecta | --comprova-pendents | --comprova-publicar | --regenera | --tot")
