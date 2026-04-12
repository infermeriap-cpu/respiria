#!/usr/bin/env python3
"""
RespirIA — Actualitzador automàtic CIMA → Excel → HTML
=======================================================
Autora: Sílvia Álvarez Vega · ICS Atenció Primària Girona
Versió: 3.0 · 2026-04-12

Flux:
  1. Consulta CIMA (sust=4&comerc=1)
  2. Detecta novetats respecte al Excel actual
  3. Infereix dosis des de PHF CatSalut 2018 + GEMA 5.5
  4. Afegeix les files noves directament al Excel original
     → marcades en GROC = pendent validació
     → tu canvies a VERD (o treus el color) = validada
  5. Amb --regenera: genera RespirIA HTML des del Excel

Ús:
    python respiria_cima_updater.py              # afegeix novetats al Excel
    python respiria_cima_updater.py --regenera   # regenera HTML
    python respiria_cima_updater.py --tot        # fa les dues coses

Requisits:
    pip install requests openpyxl beautifulsoup4
"""

import requests, openpyxl, json, re, time, os, shutil, sys
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓ
# ═══════════════════════════════════════════════════════════════════════════════
EXCEL_PATH  = "inhaladors_MPOC_ASMA_Avanzado_FINAL.xlsx"
HTML_PATH   = "RespirIA_v2.html"
OUTPUT_HTML = "RespirIA_v3.html"
CIMA_BASE   = "https://cima.aemps.es/cima/rest"

# Colors
COLOR_NOU      = "FEF3C7"  # groc  = pendent validació tu
COLOR_VALIDAT  = "D1FAE5"  # verd  = validat (canvia tu manualment)
COLOR_ATENCIO  = "FEE2E2"  # vermell = dosi no trobada, consultar fitxa tècnica

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
    "salmeterol/fluticasona":   ("50/100 µg/12h",    "50/250 µg/12h",     "50/500 µg/12h"),
    "formoterol/budesonida":    ("4.5/160 µg/12h",   "9/320 µg/12h",      "2 inh de 9/320 µg/12h"),
    "formoterol/beclometasona": ("6/100: 1 inh/12h", "6/100: 2 inh/12h",  "6/200: 2 inh/12h"),
    "vilanterol/fluticasona":   ("22/92 µg/24h",      "22/92 µg/24h",     "22/184 µg/24h"),
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
    "solucion para inhalacion":      ("ICP","🔴","20-30 l/m","Lenta","https://scientiasalut.gencat.cat/handle/11351/11880"),
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
        print(f"  → Pàg {pagina}: {len(items)} presentacions ({len(resultats)}/{total})")
        if len(resultats) >= total:
            break
        pagina += 1
        time.sleep(0.3)
    print(f"  ✅ Total: {len(resultats)} presentacions")
    return resultats

def get_fitxa_posologia(nregistro):
    try:
        data = cima_get("docSegmentado/contenido/1", {"nregistro": nregistro, "seccion": "4.2"})
        if not data or not data.get("secciones"):
            return ""
        soup = BeautifulSoup(data["secciones"][0].get("contenido",""), "html.parser")
        return soup.get_text(separator=" ", strip=True)[:500]
    except:
        return ""

def normalitza(text):
    t = text.lower()
    for p in ["bromur de ","bromur d'","propionat de ","furoat de ","dipropionat de "]:
        t = t.replace(p, "")
    return t.strip()

def cerca_dosi(principis):
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
    nl = nom.lower()
    for k, v in DISPOSITIU_MAP.items():
        if k in nl:
            return v
    return ("ICP","🔴","20-30 l/m","Lenta","https://scientiasalut.gencat.cat/handle/11351/11880")

# ═══════════════════════════════════════════════════════════════════════════════
# LECTURA DEL EXCEL
# ═══════════════════════════════════════════════════════════════════════════════

def llegeix_excel(path):
    print(f"📊 Llegint Excel: {path}")
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    cns = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        cn = str(row[3]).strip() if row[3] else ""
        for c in cn.replace(" ","").split(","):
            if c and c != "None":
                cns.add(c)
    ultima_fila = ws.max_row
    print(f"  ✅ {ultima_fila-1} files, {len(cns)} CNs registrats")
    return wb, ws, cns, ultima_fila

# ═══════════════════════════════════════════════════════════════════════════════
# AFEGIR NOVETATS AL EXCEL ORIGINAL
# ═══════════════════════════════════════════════════════════════════════════════

def afegeix_novetats_al_excel(novetats, wb, ws, ultima_fila, output_path):
    """
    Afegeix les files noves directament al full principal del Excel.
    Marcades en GROC = pendent la teva validació.
    Quan valides una fila, canvia el color a VERD (o treu el color).
    """
    print(f"\n📝 Afegint {len(novetats)} files noves al Excel...")

    fill_nou     = PatternFill("solid", fgColor=COLOR_NOU)
    fill_atencio = PatternFill("solid", fgColor=COLOR_ATENCIO)
    font_bold    = Font(bold=True)
    alineacio    = Alignment(wrap_text=True, vertical="top")

    # Seprador visual entre catàleg existent i novetats
    ws.append([""])  # fila buida separadora

    for n in novetats:
        # Construeix la fila amb la mateixa estructura que el Excel original
        # Col: Classe, IA, Nom Comercial, CN, Dispositiu, Dosi, Tipus, CO2,
        #      Flux, Flux Pas4, Info dispositiu, Data, Origen CIMA, Estat

        dosi_info  = cerca_dosi(n["principis"])
        asma_gci   = cerca_asma_gci(n["principis"])
        asma_combo = cerca_asma_combo(n["principis"])

        dosi_mpoc = dmx = ""
        starred = matma = False
        color_fila = fill_nou  # groc per defecte

        if dosi_info:
            dosi_mpoc, dmx, starred, matma = dosi_info

        # Construeix text de dosi combinat (MPOC + Asma si escau)
        dosi_text = ""
        if dosi_mpoc:
            dosi_text += f"MPOC: {dosi_mpoc}"
        if asma_gci:
            dosi_text += f"\nAsma baixa: {asma_gci[0]}\nAsma mitjana: {asma_gci[1]}\nAsma alta: {asma_gci[2]}"
        if asma_combo:
            dosi_text += f"\nAsma baixa: {asma_combo[0]}\nAsma mitjana: {asma_combo[1]}\nAsma alta: {asma_combo[2]}"
        if not dosi_text:
            # No trobat a PHF/GEMA → llegim fitxa tècnica
            posologia = get_fitxa_posologia(n["nregistro"])
            if posologia:
                dosi_text = f"FITXA TÈCNICA (revisar): {posologia[:200]}"
            else:
                dosi_text = "PENDENT — consultar fitxa tècnica CIMA"
            color_fila = fill_atencio  # vermell: requereix atenció

        tipus, co2, flux, flux4, link = infereix_dispositiu(n["nom"])

        phf_text = "★ PHF" if starred else ""
        matma_text = "MATMA" if matma else ""
        estat = "NOU — PENDENT VALIDACIÓ CLÍNICA"

        fila = [
            n["classe"],          # A: Classe terapèutica
            n["principis"],       # B: Principi actiu
            n["nom"],             # C: Nom Comercial
            n["cn"],              # D: CN
            n["nom"],             # E: Dispositiu/Presentació
            dosi_text,            # F: Dosi recomanada
            tipus,                # G: Tipus dispositiu
            co2,                  # H: Petjada CO2
            flux,                 # I: Flux inspiratori
            flux4,                # J: Flux Pas 4
            link,                 # K: Informació dispositiu
            datetime.now().strftime("%Y-%m-%d"),  # L: Data incorporació
            f"https://cima.aemps.es/cima/publico/medicamento.html?nregistro={n['nregistro']}",  # M: Origen CIMA
            estat,                # N: Estat validació
        ]

        ws.append(fila)
        nova_fila_num = ws.max_row

        # Aplica color i format a tota la fila
        for col in range(1, len(fila) + 1):
            cel = ws.cell(row=nova_fila_num, column=col)
            cel.fill = color_fila
            cel.alignment = alineacio
            if col in (1, 3):  # Classe i Nom en negreta
                cel.font = font_bold

        print(f"  ✅ Afegit: {n['nom'][:50]} (CN:{n['cn']}) — {color_fila.fgColor.rgb[-6:]}")

    # Afegeix llegenda al final del fitxer
    ws.append([""])
    llegenda_fila = ws.max_row + 1
    ws.append(["LLEGENDA DE COLORS:"])
    ws.append(["🟡 GROC = Nou fàrmac detectat a CIMA — PENDENT validació clínica"])
    ws.append(["🔴 VERMELL = Dosi no trobada a PHF/GEMA — Consultar fitxa tècnica CIMA obligatòriament"])
    ws.append(["🟢 VERD = Validat per professional sanitari"])
    ws.append(["Per validar: comprova la fila, canvia el color a verd (o treu el color) i guarda el fitxer"])

    for r in range(llegenda_fila, ws.max_row + 1):
        ws.cell(row=r, column=1).font = Font(italic=True, color="6B7280")

    # Guarda sobre el Excel original (o una còpia si vols preservar l'original)
    shutil.copy(output_path, output_path.replace(".xlsx", "_BACKUP.xlsx"))
    wb.save(output_path)
    print(f"\n  ✅ Excel actualitzat: {output_path}")
    print(f"  💾 Backup guardat: {output_path.replace('.xlsx', '_BACKUP.xlsx')}")

# ═══════════════════════════════════════════════════════════════════════════════
# REGENERACIÓ HTML
# ═══════════════════════════════════════════════════════════════════════════════

def regenera_html(html_path, output_path, ws):
    print(f"\n🔄 Generant HTML: {output_path}")
    with open(html_path, "r", encoding="utf-8") as f:
        html = f.read()

    entrades = []
    classe_actual = None

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        v = lambda x: str(x).strip() if x and str(x).strip() not in ("None","") else ""

        classe = v(row[0]) or classe_actual
        if classe:
            classe_actual = classe.split("\n")[0].strip()
        else:
            classe = classe_actual or ""

        ia   = v(row[1])
        nom  = v(row[2])
        if not nom or nom.startswith("LLEGENDA") or nom.startswith("🟡") or nom.startswith("🔴"):
            continue

        # Salta files marcades com a PENDENTS (no validades)
        estat = v(row[13]) if len(row) > 13 else ""
        if "PENDENT" in estat.upper() or "NOU —" in estat.upper():
            print(f"  ⏭  Saltat (no validat): {nom[:40]}")
            continue

        disp   = v(row[4])
        dosi_r = v(row[5])
        tipus_r= v(row[6])
        link   = v(row[10]) if len(row) > 10 else ""

        t = tipus_r.lower().replace(" ","").replace("_","-")
        if "multi" in t:              tipus_js = "IPS-multi"
        elif "uni" in t:              tipus_js = "IPS-uni"
        elif "ivs" in t or "respimat" in t: tipus_js = "IVS"
        else:                         tipus_js = "ICP"

        mpoc = asma_b = asma_m = asma_a = dmx = ""
        for l in (dosi_r or "").split("\n"):
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
        if ia:
            obj += f",ia:{json.dumps(ia, ensure_ascii=False)}"
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

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_nou)
    print(f"  ✅ HTML generat: {output_path} ({len(entrades)} entrades al catàleg)")

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    print()
    print("╔══════════════════════════════════════════════════════════╗")
    print("║  RespirIA — Actualitzador CIMA v3.0                      ║")
    print("║  PHF CatSalut 2018 · GEMA 5.5 · Fitxa tècnica CIMA      ║")
    print(f"║  {datetime.now().strftime('%Y-%m-%d %H:%M')}                                        ║")
    print("╚══════════════════════════════════════════════════════════╝\n")

    for f in [EXCEL_PATH, HTML_PATH]:
        if not os.path.exists(f):
            print(f"❌ No trobo: {f}")
            print("   Posa el script a la mateixa carpeta que el Excel i el HTML.")
            return

    mode_regenera = "--regenera" in sys.argv
    mode_tot      = "--tot" in sys.argv

    wb, ws, cns_existents, ultima_fila = llegeix_excel(EXCEL_PATH)

    if not mode_regenera:
        # Consulta CIMA i afegeix novetats
        presentacions = get_tots_inhaladors_cima()
        print()

        novetats_raw = [p for p in presentacions
                        if str(p.get("cn","")).strip() not in cns_existents]
        print(f"🔍 Novetats detectades: {len(novetats_raw)}")

        if novetats_raw:
            print(f"\n🧬 Enriquint {min(len(novetats_raw), 50)} novetats...")
            novetats = []
            for i, p in enumerate(novetats_raw[:50], 1):
                cn = str(p.get("cn","")).strip()
                nom = p.get("nombre","")
                nregistro = p.get("nregistro","")
                print(f"  [{i}/{min(len(novetats_raw),50)}]", end=" ")

                med = cima_get("medicamento", {"cn": cn}) or {}
                atcs = med.get("atcs", [])
                pas  = med.get("principiosActivos", [])
                principis = ", ".join([x.get("nombre","") for x in pas])

                novetats.append({
                    "cn": cn, "nom": nom, "nregistro": nregistro,
                    "classe": infereix_classe(atcs),
                    "principis": principis,
                })
                print(f"{nom[:50]}")
                time.sleep(0.4)

            afegeix_novetats_al_excel(novetats, wb, ws, ultima_fila, EXCEL_PATH)

            print()
            print("━"*62)
            print("✋  QUÈ HAS DE FER ARA:")
            print()
            print(f"  1. Obre: {EXCEL_PATH}")
            print("  2. Baixa fins al final — trobaràs les files noves marcades")
            print("  3. Per cada fila:")
            print("     🟡 GROC   → revisa dosi, PHF, MATMA · canvia a verd si OK")
            print("     🔴 VERMELL → consulta fitxa tècnica CIMA (link a columna M)")
            print("  4. Un cop validades totes, guarda el fitxer")
            print("  5. Executa:")
            print("     python respiria_cima_updater.py --regenera")
            print("━"*62)
        else:
            print("✅ Cap novetat. El catàleg ja està al dia.")
            if not mode_tot:
                print("\nPots regenerar el HTML amb:")
                print("  python respiria_cima_updater.py --regenera")

    if mode_regenera or mode_tot or not novetats_raw:
        # Recarrega el Excel per si ha canviat
        wb2 = openpyxl.load_workbook(EXCEL_PATH)
        ws2 = wb2.active
        regenera_html(HTML_PATH, OUTPUT_HTML, ws2)
        print(f"\n✅ Artefacte llest: {OUTPUT_HTML}")
        print("   Obre'l al navegador per comprovar-ho.")

    print(f"\nFet! {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")

if __name__ == "__main__":
    main()
