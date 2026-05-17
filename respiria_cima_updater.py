#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
respiria_cima_updater.py  —  v6.1
===================================
Autor: Sílvia Álvarez Vega  |  ICS Atenció Primària Girona
Projecte: RespirIA — Prescriptor digital d'inhaladors

MODES
-----
  --detecta            Consulta CIMA, detecta CNs nous, els afegeix a
                       l'Excel com "NOU — PENDENT" i envia correu.

  --comprova-pendents  Compta files amb "NOU — PENDENT" → pendents_count.txt

  --comprova-publicar  Compta novetats validades (col N buida + col M=CIMA
                       + col L té data) → publicar_count.txt

  --regenera           FUSIÓ: llegeix el catalog del HTML actual (fàrmacs
                       ja publicats) + afegeix les novetats validades de
                       l'Excel → desa el nou RespirIA_v2.html.

  --tot                Executa --detecta + --regenera (ús local)

LÒGICA DE FUSIÓ (--regenera)
------------------------------
  El --regenera NO toca els fàrmacs ja publicats al HTML.
  Només afegeix les novetats que compleixen les 3 condicions:
    1. Col N = buida  (la Sílvia ha validat, ha esborrat "NOU — PENDENT")
    2. Col M conté URL de cima.aemps.es  (prové del --detecta, no és manual)
    3. Col L té data  (posada automàticament pel --detecta)

  Així els fàrmacs originals de l'Excel (sense URL CIMA ni data)
  es conserven intactes al HTML i no es re-processen.

ESTRUCTURA EXCEL MESTRE (columnes A..P)
---------------------------------------
  A  Classe terapèutica
  B  Principi actiu
  C  Nom comercial
  D  Codi nacional CN
  E  Dispositiu/Presentació
  F  Dosi recomanada
  G  Tipus dispositiu  (ICP / IVS / IPS-multi / IPS-uni / NEB)
  H  Petjada CO₂
  I  Flux inspiratori
  J  Maniobra
  K  Link instruccions
  L  Data incorporació  ← posada pel --detecta (novetats)
  M  Origen CIMA        ← URL cima.aemps.es (posada pel --detecta)
  N  Estat validació    ← "NOU — PENDENT" / "NO INCORPORAR" / buida=validat
  O  PHF ★
  P  MATMA
"""

import argparse
import json
import os
import re
import sys
import time
from datetime import date

import openpyxl
import requests

# ── Configuració ──────────────────────────────────────────────────────────────
GITHUB_OWNER = "infermeriap-cpu"
GITHUB_REPO  = "respiria"
EXCEL_FNAME  = "inhaladors_MPOC_ASMA_Avanzado_FINAL.xlsx"
HTML_FNAME   = "RespirIA_v2.html"
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")

CIMA_URL    = "https://cima.aemps.es/cima/rest/medicamentos"
CIMA_PARAMS = {"vias": "78", "atc": "R03", "comerc": "1", "pagina": 1, "tamPagina": 100}

# ── Mapeig Tipus dispositiu ───────────────────────────────────────────────────
TIPUS_MAP = [
    (["pMDI", "Modulite", "Aerosphere", "MDI", "aerosol pressuritzat",
      "inhalador pressuritzat"], "ICP"),
    (["Respimat", "SMI", "vapor suau"], "IVS"),
    (["Turbuhaler", "Accuhaler", "Genuair", "Ellipta", "Novolizer", "Easyhaler",
      "Nexthaler", "Spiromax", "Forspiro", "Twisthaler",
      "IPS multidosi", "pols seca multidosi"], "IPS-multi"),
    (["Aerolizer", "Breezhaler", "Handihaler", "Zonda",
      "IPS unidosi", "pols seca unidosi", "càpsula"], "IPS-uni"),
]

def inferir_tipus(disp_text: str) -> str:
    if not disp_text:
        return "ICP"
    d = str(disp_text)
    for keywords, tipus in TIPUS_MAP:
        for kw in keywords:
            if kw.lower() in d.lower():
                return tipus
    return "ICP"

# ── Mapeig Dosi ───────────────────────────────────────────────────────────────
def parsear_dosi(dosi_text: str, cat: str) -> dict:
    if not dosi_text or str(dosi_text).strip() in ("", "nan"):
        return {}
    text = str(dosi_text).strip()
    result = {}

    dmx_match = re.search(r"[Dd]\.?\s*[Mm]à?x[\.:]?\s*(.+?)(?:\n|$)", text)
    if dmx_match:
        result["dmx"] = dmx_match.group(1).strip()

    if "MPOC" in text and ("Asma" in text or "asma" in text):
        mpoc_match = re.search(r"MPOC\s*[:\-]?\s*(.+?)(?=Asma|$)", text, re.IGNORECASE | re.DOTALL)
        if mpoc_match:
            result["mpoc"] = mpoc_match.group(1).strip().split("\n")[0].strip()
        baixa_match = re.search(r"[Bb]aixa\s*[:\-]?\s*(.+?)(?=Mitjana|Alta|D\.?\s*[Mm]à?x|$)", text, re.IGNORECASE | re.DOTALL)
        mitj_match  = re.search(r"Mitjana\s*[:\-]?\s*(.+?)(?=Alta|D\.?\s*[Mm]à?x|$)", text, re.IGNORECASE | re.DOTALL)
        alta_match  = re.search(r"Alta\s*[:\-]?\s*(.+?)(?=D\.?\s*[Mm]à?x|$)", text, re.IGNORECASE | re.DOTALL)
        if baixa_match: result["asma_baixa"]  = baixa_match.group(1).strip().split("\n")[0].strip()
        if mitj_match:  result["asma_mitjana"] = mitj_match.group(1).strip().split("\n")[0].strip()
        if alta_match:  result["asma_alta"]    = alta_match.group(1).strip().split("\n")[0].strip()
    else:
        dosi_neta = re.sub(r"D\.?\s*[Mm]à?x[\.:]?.+", "", text).strip()
        if cat in ("GCI", "LABA/GCI", "LAMA/LABA/GCI", "SABA/GCI"):
            result["mpoc"]       = dosi_neta
            result["asma_baixa"] = dosi_neta
        else:
            result["dose"] = dosi_neta

    return result

# ── Descarregar fitxers de GitHub ─────────────────────────────────────────────
def download_from_github(fname: str, mode: str = "binary") -> str:
    url = f"https://raw.githubusercontent.com/{GITHUB_OWNER}/{GITHUB_REPO}/main/{fname}"
    headers = {}
    if GITHUB_TOKEN:
        headers["Authorization"] = f"token {GITHUB_TOKEN}"
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    local_path = f"/tmp/{fname}"
    if mode == "text":
        with open(local_path, "w", encoding="utf-8") as f:
            f.write(r.text)
    else:
        with open(local_path, "wb") as f:
            f.write(r.content)
    print(f"  Descarregat {fname} ({len(r.content)/1024:.1f} KB)")
    return local_path

# ── Llegir novetats validades de l'Excel ──────────────────────────────────────
def llegir_novetats_validades(excel_path: str) -> list:
    """
    Retorna NOMÉS les files que són novetats acabades de validar:
      - Col N = buida  (validat per la Sílvia)
      - Col M conté URL de cima.aemps.es  (prové del --detecta)
      - Col L té data  (posada pel --detecta)

    Els fàrmacs originals de l'Excel (sense URL CIMA a col M)
    s'ignoren → el HTML ja els té i no cal re-processar-los.
    """
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb.active

    novetats = []
    COL = {
        "cat": 0, "ia": 1, "brand": 2, "cn": 3, "disp": 4,
        "dose": 5, "tipus": 6, "co2": 7, "flux": 8, "man": 9,
        "link": 10, "data": 11, "origen": 12, "estat": 13,
        "phf": 14, "matma": 15,
    }

    skip_row = True
    for row in ws.iter_rows(values_only=True):
        if skip_row:
            skip_row = False
            continue
        if len(row) <= COL["estat"]:
            continue

        estat  = str(row[COL["estat"]]  or "").strip()
        origen = str(row[COL["origen"]] or "").strip()
        data_l = row[COL["data"]]

        # Les 3 condicions per ser "novetat validada"
        if estat not in ("", "None"):
            continue                           # NOU — PENDENT o NO INCORPORAR
        if "cima.aemps.es" not in origen:
            continue                           # fàrmac original manual → HTML ja el té
        if not data_l:
            continue                           # sense data → origen manual

        cat    = str(row[COL["cat"]]   or "").strip()
        ia     = str(row[COL["ia"]]    or "").strip()
        if not cat or not ia:
            continue

        brand  = str(row[COL["brand"]] or "").strip()
        disp   = str(row[COL["disp"]]  or "").strip()
        cn     = str(row[COL["cn"]]    or "").strip()
        link   = str(row[COL["link"]]  or "").strip()
        phf    = row[COL["phf"]]
        matma  = row[COL["matma"]]
        dosi_f = str(row[COL["dose"]]  or "").strip()
        tipus_f= str(row[COL["tipus"]] or "").strip()

        if "NEB" in tipus_f.upper() or "nebulitz" in disp.lower():
            continue

        tipus = tipus_f if tipus_f and tipus_f not in ("", "nan", "None") else inferir_tipus(disp)
        dosi_dict = parsear_dosi(dosi_f, cat)
        starred = bool(phf and str(phf).strip() not in ("", "None", "nan", "0", "False"))
        matma_b = bool(matma and str(matma).strip() not in ("", "None", "nan", "0", "False"))
        cns = [c.strip() for c in re.split(r"[,\n]", cn) if c.strip()]

        farmac = {
            "cat": cat, "ia": ia, "brands": brand, "disp": disp,
            "cn": ", ".join(cns), "tipus": tipus,
            "link": link if link.startswith("http") else "",
            "starred": starred, "matma": matma_b,
        }
        farmac.update(dosi_dict)
        novetats.append(farmac)

    wb.close()
    print(f"  Novetats validades a afegir: {len(novetats)}")
    return novetats

# ── Generar JS per a un fàrmac ────────────────────────────────────────────────
def farmac_a_js(d: dict) -> str:
    parts = []
    parts.append(f"cat:{json.dumps(d['cat'], ensure_ascii=False)}")
    parts.append(f"ia:{json.dumps(d['ia'], ensure_ascii=False)}")
    if d.get("starred"): parts.append("starred:true")
    if d.get("matma"):   parts.append("matma:true")
    parts.append(f"brands:{json.dumps(d['brands'], ensure_ascii=False)}")
    parts.append(f"disp:{json.dumps(d['disp'], ensure_ascii=False)}")
    for camp in ("dose", "mpoc", "asma_baixa", "asma_mitjana", "asma_alta"):
        if d.get(camp):
            parts.append(f"{camp}:{json.dumps(d[camp], ensure_ascii=False)}")
    if d.get("dmx"):
        parts.append(f"dmx:{json.dumps(d['dmx'], ensure_ascii=False)}")
    parts.append(f"tipus:{json.dumps(d['tipus'], ensure_ascii=False)}")
    if d.get("link"):
        parts.append(f"link:{json.dumps(d['link'], ensure_ascii=False)}")
    return "  {" + ",\n   ".join(parts) + "}"

# ── Fusionar catalog HTML + novetats Excel ────────────────────────────────────
def fusionar_catalog(html_path: str, novetats: list) -> str:
    """
    Llegeix el HTML actual, extreu el var catalog=[...] existent,
    li AFEGEIX les novetats al final i retorna el HTML actualitzat.
    """
    with open(html_path, "r", encoding="utf-8") as f:
        html = f.read()

    pattern = re.compile(r"(var catalog=\[)(.*?)(\];)", re.DOTALL)
    match = pattern.search(html)
    if not match:
        raise ValueError("No s'ha trobat 'var catalog=[...];' al HTML.")

    catalog_actual = match.group(2).strip()  # contingut entre [ i ]
    n_actuals = catalog_actual.count("{cat:")
    print(f"  Fàrmacs ja al HTML: {n_actuals}")
    print(f"  Novetats a afegir:  {len(novetats)}")

    # Construir les línies JS de les novetats
    novetats_js = ""
    if novetats:
        novetats_lines = [farmac_a_js(f) for f in novetats]
        novetats_js = ",\n" + ",\n".join(novetats_lines)

    # Nou catalog = contingut actual + novetats
    nou_catalog_bloc = f"var catalog=[{catalog_actual}{novetats_js}];"

    html_nou = html[:match.start()] + nou_catalog_bloc + html[match.end():]
    print(f"  ✅ Total fàrmacs al prescriptor: {n_actuals + len(novetats)}")
    return html_nou

# ── MODE: --regenera ──────────────────────────────────────────────────────────
def mode_regenera():
    """
    Fusiona el catalog existent al HTML + novetats validades de l'Excel.
    No toca els fàrmacs originals.
    """
    print("\n=== MODE REGENERA (fusió) ===")

    print("1. Descarregant Excel de GitHub...")
    excel_path = download_from_github(EXCEL_FNAME, mode="binary")

    print("2. Llegint novetats validades de l'Excel...")
    novetats = llegir_novetats_validades(excel_path)

    if not novetats:
        print("  ℹ️  Cap novetat validada. El HTML no s'actualitzarà.")
        with open("publicar_count.txt", "w") as f:
            f.write("0")
        return

    print("3. Descarregant HTML actual de GitHub...")
    html_path = download_from_github(HTML_FNAME, mode="text")

    print("4. Fusionant catalog HTML + novetats...")
    html_nou = fusionar_catalog(html_path, novetats)

    print("5. Desant HTML actualitzat...")
    with open(HTML_FNAME, "w", encoding="utf-8") as f:
        f.write(html_nou)
    print(f"  ✅ {HTML_FNAME} desat ({len(html_nou)/1024:.1f} KB)")

    with open("publicar_count.txt", "w") as f:
        f.write(str(len(novetats)))

    print(f"\n✅ REGENERA completat: {len(novetats)} novetats afegides al prescriptor.")

# ── MODE: --comprova-publicar ─────────────────────────────────────────────────
def mode_comprova_publicar():
    """Compta novetats validades → publicar_count.txt"""
    print("\n=== MODE COMPROVA-PUBLICAR ===")
    excel_path = download_from_github(EXCEL_FNAME, mode="binary")
    novetats = llegir_novetats_validades(excel_path)
    with open("publicar_count.txt", "w") as f:
        f.write(str(len(novetats)))
    print(f"✅ {len(novetats)} novetats validades llestos per publicar → publicar_count.txt")

# ── MODE: --comprova-pendents ─────────────────────────────────────────────────
def mode_comprova_pendents():
    """Compta files amb 'NOU — PENDENT' → pendents_count.txt"""
    print("\n=== MODE COMPROVA-PENDENTS ===")
    excel_path = download_from_github(EXCEL_FNAME, mode="binary")
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb.active
    count = 0
    skip = True
    for row in ws.iter_rows(values_only=True):
        if skip:
            skip = False
            continue
        estat = str(row[13] or "").strip() if len(row) > 13 else ""
        if "NOU" in estat and "PENDENT" in estat:
            count += 1
    wb.close()
    with open("pendents_count.txt", "w") as f:
        f.write(str(count))
    print(f"✅ {count} fàrmacs pendents de validació → pendents_count.txt")

# ── MODE: --detecta ───────────────────────────────────────────────────────────
# Formes farmacèutiques que NO són inhaladors → descartar
FORMES_EXCLOURE = [
    "comprimido", "comprimits", "tableta", "capsula dura oral",
    "solucion oral", "solució oral", "jarabe", "xarop",
    "inyectable", "injectable", "infusion", "perfusion",
    "nasal", "oftalmico", "topico", "cutaneo", "crema",
    "nebulizador", "nebulitzador", "inhalacion por nebulizador",
    "polvo para reconstitucion", "granulado",
]

def es_inhalador_valid(forma: str) -> bool:
    """Retorna True si la forma farmacèutica és un inhalador vàlid (no oral, no injectable, no nebulitzador)."""
    if not forma:
        return False
    f = forma.lower()
    for excl in FORMES_EXCLOURE:
        if excl in f:
            return False
    # Ha de contenir alguna paraula d'inhalador
    paraules_inhalador = [
        "inhal", "aerosol", "polvo para inh", "pols per a inh",
        "vapor", "respimat", "turbuhaler", "accuhaler", "ellipta",
        "breezhaler", "genuair", "novolizer", "easyhaler",
        "nexthaler", "spiromax", "handihaler", "aerolizer",
        "pressuritzat", "pressurizado",
    ]
    return any(p in f for p in paraules_inhalador)

def inferir_cat_des_atc(atc: str, ia: str) -> str:
    """Retorna la classe terapèutica o None si el codi ATC no és reconegut."""
    ia_l = ia.lower()
    if "R03AC" in atc:
        return "SABA" if any(x in ia_l for x in ["salbutamol", "terbutalin"]) else "LABA"
    if "R03AL" in atc:
        te_gci = any(x in ia_l for x in ["beclometasona", "fluticasona", "budesonida", "mometasona"])
        te_lama = any(x in ia_l for x in ["glicopirroni", "umeclidini", "aclidini"])
        if te_lama and te_gci: return "LAMA/LABA/GCI"
        if te_gci:             return "LABA/GCI"
        return "LABA/LAMA"
    if "R03AK" in atc:
        te_lama = any(x in ia_l for x in ["glicopirroni", "umeclidini", "aclidini", "tiotropi"])
        te_gci = any(x in ia_l for x in ["beclometasona", "fluticasona", "budesonida", "mometasona"])
        if te_lama and te_gci: return "LAMA/LABA/GCI"
        if te_lama:            return "LABA/LAMA"
        return "LABA/GCI"
    if "R03BB" in atc:
        return "SAMA" if any(x in ia_l for x in ["ipratropi", "ipratropium"]) else "LAMA"
    if "R03BA" in atc: return "GCI"
    if "R03AB" in atc: return "SAMA"
    if "R03CC" in atc: return "SABA"
    # Codi ATC no reconegut → descartar (no afegir al Excel)
    return None

def mode_detecta():
    print("\n=== MODE DETECTA ===")

    print("1. Descarregant Excel actual...")
    excel_path = download_from_github(EXCEL_FNAME, mode="binary")

    # Recollir CNs existents
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    cns_existents = set()
    skip = True
    for row in ws.iter_rows(values_only=True):
        if skip:
            skip = False
            continue
        cn_cell = str(row[3] or "").strip()
        for cn in re.split(r"[,\n\s]+", cn_cell):
            cn = cn.strip()
            if cn:
                cns_existents.add(cn)
    print(f"  CNs existents: {len(cns_existents)}")

    # Consultar CIMA
    print("2. Consultant CIMA...")
    medicaments_cima = []
    pagina = 1
    while True:
        params = {**CIMA_PARAMS, "pagina": pagina}
        r = requests.get(CIMA_URL, params=params, timeout=30)
        r.raise_for_status()
        data = r.json()
        items = data.get("resultados", [])
        if not items:
            break
        medicaments_cima.extend(items)
        if len(items) < 100:
            break
        pagina += 1
        time.sleep(0.5)
    print(f"  Total CIMA: {len(medicaments_cima)}")

    # Detectar novetats
    novetats = []
    for med in medicaments_cima:
        cn = str(med.get("cn", "")).strip().zfill(6)
        if cn not in cns_existents:
            novetats.append(med)
    print(f"  Novetats: {len(novetats)}")

    if not novetats:
        with open("novetats_count.txt", "w") as f:
            f.write("0")
        print("✅ Cap novetat.")
        return

    # Afegir novetats a l'Excel
    today = date.today().isoformat()
    afegits = 0
    descartats = 0
    for med in novetats:
        cn   = str(med.get("cn", "")).strip().zfill(6)
        nom  = med.get("nombre", "").strip()
        vtm  = med.get("vtm", {})
        ia   = vtm.get("nombre", nom) if vtm else nom
        nreg = med.get("nregistro", "")
        cima_url = f"https://cima.aemps.es/cima/publico/detalle.html?nregistro={nreg}"
        disp = med.get("formaFarmaFull", "") or med.get("formaFarma", "")

        # Filtre 1: forma farmacèutica ha de ser inhalador vàlid
        if not es_inhalador_valid(disp):
            descartats += 1
            continue

        tipus = inferir_tipus(disp)
        atc_grup = med.get("atc", [{}])[0].get("codigo", "") if med.get("atc") else ""
        cat = inferir_cat_des_atc(atc_grup, ia)

        # Filtre 2: categoria ha de ser reconeguda (no None)
        if cat is None:
            descartats += 1
            continue

        ws.append([
            cat, ia, nom, cn, disp,
            "",      # F dosi (a validar)
            tipus,   # G tipus
            "",      # H CO₂
            "",      # I flux
            "",      # J maniobra
            "",      # K link
            today,   # L data ← posada aquí
            cima_url,# M origen CIMA ← posada aquí
            "NOU — PENDENT",  # N estat
            "", "",  # O PHF, P MATMA
        ])
        afegits += 1

    wb.save(EXCEL_FNAME)
    with open("novetats_count.txt", "w") as f:
        f.write(str(afegits))
    print(f"✅ DETECTA: {afegits} novetats afegides com 'NOU — PENDENT' ({descartats} descartades per forma/ATC incorrecte).")

# ── MAIN ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="RespirIA CIMA Updater v6.1")
    parser.add_argument("--mode", choices=["detecta", "comprova-pendents",
                        "comprova-publicar", "regenera", "tot"], required=True)
    args = parser.parse_args()

    if args.mode == "detecta":
        mode_detecta()
    elif args.mode == "comprova-pendents":
        mode_comprova_pendents()
    elif args.mode == "comprova-publicar":
        mode_comprova_publicar()
    elif args.mode == "regenera":
        mode_regenera()
    elif args.mode == "tot":
        mode_detecta()
        mode_regenera()
