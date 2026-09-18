"""
fetch_mef_data.py
=================
Descarga el "Resumen General de Tenencias de Bonos Soberanos" del MEF (PDF),
lo parsea y actualiza los datos del dashboard.

Por qué este parser es distinto al anterior
-------------------------------------------
El PDF del MEF es un export de Excel con GRÁFICOS DE TORTA: una torta por bono,
dos por fila, y las etiquetas ("Bancos" / "63.78%") flotan alrededor de cada torta.
Al extraer el texto "plano" el orden de lectura se mezcla entre las dos tortas de la
fila, así que las regex del tipo "Bancos\\s+NN.NN%" fallaban o capturaban valores de
otra torta. Este parser trabaja con COORDENADAS:

  1. Extrae cada fragmento de texto con su posición (x0, x1, top).
  2. Detecta los títulos "Bonos Soberanos 12AGO2031" + "14 560 480 Unidades".
  3. Empareja cada etiqueta de inversor con el % que está justo debajo.
  4. Asigna cada par a la torta más cercana (misma fila, centro X más próximo).
  5. Valida que cada torta sume ~100%; si no, re-asigna las etiquetas ambiguas
     (las que quedan entre dos tortas) probando combinaciones.

Las unidades en circulación de cada bono vienen en el mismo PDF (debajo del título),
así que ya NO se necesita el PDF de stock.

Uso
---
    pip install -r requirements.txt
    python scripts/fetch_mef_data.py                 # último reporte publicado
    python scripts/fetch_mef_data.py --force         # re-procesa aunque ya exista
    python scripts/fetch_mef_data.py --backfill 4    # re-procesa los 4 últimos meses en orden
    python scripts/fetch_mef_data.py --pdf archivo.pdf   # usa un PDF local
"""

from __future__ import annotations

import argparse
import io
import itertools
import json
import re
import sys
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
INDEX_HTML = ROOT / "index.html"
HIST_FILE = DATA_DIR / "historico_evolucion.json"
YIELDS_FILE = DATA_DIR / "yields.json"

MEF_BASE = "https://www.mef.gob.pe"
TENENCIAS_URL = MEF_BASE + "/contenidos/deuda_publ/mercado/reportes_tenencia.php"
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; SoberanosDashboard/3.0)"}

# ─────────────────────────────────────────────────────────────────
# Diccionarios
# ─────────────────────────────────────────────────────────────────
# Etiqueta MEF (normalizada) → categoría del dashboard
LABELS = {
    "no residentes": "Offshores",
    "afps": "Pension Funds",
    "bancos": "Banks",
    "seguros": "Insurance",
    "fondos publicos": "Public Funds",
    "fondos privados": "Others",
    "otros": "Others",
    "personas naturales": "Others",
}
# Primeras palabras de etiquetas que Excel a veces parte en 2 líneas
SPLIT_HEADS = {"personas": "naturales", "fondos": ("privados", "publicos")}

CATS = ["Offshores", "Pension Funds", "Banks", "Insurance", "Public Funds", "Others"]
DV01_CATS = ["Offshores", "Pension Funds", "Banks", "Insurance"]

# Cupones (fijos por emisión; fuente: MEF, Stock de Bonos Soberanos)
CUPONES = {
    "12AGO2026": 8.20, "12AGO2028": 6.35, "12FEB2029": 6.00, "12FEB2029E": 5.94,
    "12AGO2031": 6.95, "12AGO2032": 6.15, "12AGO2033": 7.30, "12AGO2034": 5.40,
    "12AGO2035": 6.85, "12AGO2037": 6.90, "12AGO2039": 7.60, "12AGO2040": 5.35,
    "12FEB2042": 6.85, "12FEB2055": 6.7142,
}

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
MES_ABR_ES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Set", "Oct", "Nov", "Dic"]
MES_ABR_EN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MES_BONO = {"ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6, "JUL": 7,
            "AGO": 8, "SET": 9, "SEP": 9, "OCT": 10, "NOV": 11, "DIC": 12}

BONO_RE = re.compile(r"Bonos\s+Soberanos\s+(\d{2}[A-Z]{3}\d{4}E?)\b")
PCT_RE = re.compile(r"^(\d{1,3}(?:[.,]\d+)?)\s*%$")
UNITS_RE = re.compile(r"(\d{1,3}(?:[ ., ]\d{3})+|\d{4,})")


def _norm(s: str) -> str:
    s = s.lower().strip()
    for a, b in (("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u")):
        s = s.replace(a, b)
    return re.sub(r"\s+", " ", s)


def bono_fecha(bono_id: str) -> date:
    m = re.match(r"(\d{2})([A-Z]{3})(\d{4})", bono_id)
    return date(int(m.group(3)), MES_BONO[m.group(2)], int(m.group(1)))


def bono_sob(bono_id: str) -> str:
    return "SOB" + bono_id[7:9]


def bono_tenor(bono_id: str) -> str:
    return bono_id[7:9] + "s"


# ─────────────────────────────────────────────────────────────────
# 1. Descubrir y descargar PDFs
# ─────────────────────────────────────────────────────────────────
def _get(url: str):
    import requests
    r = requests.get(url, headers=HEADERS, timeout=90)
    r.raise_for_status()
    return r


def listar_reportes() -> list[tuple[date, str]]:
    """Lista (fecha_corte, url) de todos los PDFs de tenencias, ordenados por fecha."""
    html = _get(TENENCIAS_URL).text
    out = {}
    for href in re.findall(r'href="([^"]*tenencia_bono_(\d{6})\.pdf)"', html, re.I):
        url, ddmmyy = href
        try:
            d = datetime.strptime(ddmmyy, "%d%m%y").date()
        except ValueError:
            continue
        if not url.startswith("http"):
            url = MEF_BASE + ("" if url.startswith("/") else "/contenidos/deuda_publ/mercado/") + url
        out[d] = url
    if not out:
        raise RuntimeError("No se encontraron PDFs de tenencias en " + TENENCIAS_URL)
    return sorted(out.items())


# ─────────────────────────────────────────────────────────────────
# 2. Extraer fragmentos de texto con coordenadas
# ─────────────────────────────────────────────────────────────────
TOKEN_RE = re.compile(
    r"Bonos\s*Soberanos\s*\d{2}[A-Z]{3}\d{4}E?"      # título de torta
    r"|\d{1,3}(?:[.,]\d+)?\s*%"                         # porcentaje
    r"|\d[\d \u00a0]{2,}\d(?:\s*Unidades)?"          # unidades
    r"|no\s*residentes|afps|bancos|seguros|otros"
    r"|fondos\s*p[úu]blicos|fondos\s*privados|personas\s*naturales"
    r"|personas|naturales|fondos|privados|p[úu]blicos",
    re.I)


def _runs_de_pagina(page, pno):
    """Agrupa los caracteres en fragmentos (misma línea y contiguos) y luego los
    parte en tokens. Se hace a mano (no extract_words) porque Excel exporta algunas
    etiquetas con caracteres ligeramente rotados o pegadas a la vecina
    (ej. 'SegurosAFPs'), y extract_words las deja letra por letra o fusionadas."""
    chars = [c for c in page.chars if c["text"].strip() or c["text"] == " "]
    chars.sort(key=lambda c: (round(c["top"]), c["x0"]))
    lineas = []
    for c in sorted(chars, key=lambda c: c["top"]):
        for ln in lineas:
            if abs(ln["top"] - c["top"]) <= 1.5:
                ln["chars"].append(c)
                break
        else:
            lineas.append({"top": c["top"], "chars": [c]})
    runs = []
    for ln in lineas:
        cs = sorted(ln["chars"], key=lambda c: c["x0"])
        grupos, cur = [], [cs[0]]
        for c in cs[1:]:
            gap = c["x0"] - cur[-1]["x1"]
            if gap > max(2.0, 0.6 * float(c.get("size", 6))):
                grupos.append(cur)
                cur = [c]
            else:
                cur.append(c)
        grupos.append(cur)
        for g in grupos:
            text, xs = "", []
            for k, c in enumerate(g):
                if k and c["x0"] - g[k - 1]["x1"] > 0.25 * float(c.get("size", 6)) and not text.endswith(" "):
                    text += " "
                    xs.append((c["x0"], c["x0"]))
                for ch in c["text"]:
                    text += ch
                    xs.append((c["x0"], c["x1"]))
            top = min(c["top"] for c in g)
            # partir en tokens + resto
            pos = 0
            cortes = []
            for m in TOKEN_RE.finditer(text):
                if m.start() > pos:
                    cortes.append((pos, m.start()))
                cortes.append((m.start(), m.end()))
                pos = m.end()
            if pos < len(text):
                cortes.append((pos, len(text)))
            for i0, i1 in cortes:
                t = text[i0:i1].strip()
                if not t:
                    continue
                runs.append({"page": pno, "x0": float(xs[i0][0]), "x1": float(xs[i1 - 1][1]),
                             "top": float(top), "text": t})
    return runs


def extraer_runs(pdf_bytes: bytes) -> list[dict]:
    """Devuelve [{page, x0, x1, top, text}] – un item por etiqueta/valor del PDF."""
    import pdfplumber

    runs = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for pno, page in enumerate(pdf.pages, start=1):
            runs.extend(_runs_de_pagina(page, pno))
    return runs


def _cx(r):
    return (r["x0"] + r["x1"]) / 2


# ─────────────────────────────────────────────────────────────────
# 3. Parsear el PDF de tenencias
# ─────────────────────────────────────────────────────────────────
def parse_periodo(runs) -> tuple[int, int]:
    txt = " ".join(r["text"] for r in runs if r["page"] == 1)
    m = re.search(r"(" + "|".join(MESES) + r"|setiembre)\s+de\s+(\d{4})", txt, re.I)
    if not m:
        raise RuntimeError("No se encontró el período (ej. 'Julio de 2026') en el PDF")
    mes = m.group(1).lower()
    mes = 9 if mes == "setiembre" else MESES.index(mes) + 1
    return int(m.group(2)), mes


def _etiquetas(runs_page):
    """Encuentra etiquetas de inversor (uniendo las partidas en 2 líneas)."""
    usados, labels = set(), []
    for i, r in enumerate(runs_page):
        if i in usados:
            continue
        n = _norm(r["text"])
        if n in LABELS:
            labels.append({**r, "cat": LABELS[n], "key": n})
            continue
        if n in SPLIT_HEADS:
            tails = SPLIT_HEADS[n]
            tails = (tails,) if isinstance(tails, str) else tails
            for j, s in enumerate(runs_page):
                if j in usados or j == i:
                    continue
                if (_norm(s["text"]) in tails and 0 < s["top"] - r["top"] <= 9
                        and abs(_cx(s) - _cx(r)) <= 12):
                    key = n + " " + _norm(s["text"])
                    labels.append({**r, "x0": min(r["x0"], s["x0"]), "x1": max(r["x1"], s["x1"]),
                                   "top": s["top"], "cat": LABELS[key], "key": key})
                    usados.update({i, j})
                    break
    return labels


def _emparejar_valores(labels, runs_page):
    """Para cada etiqueta, el % inmediatamente debajo (emparejamiento global por distancia)."""
    pcts = [r for r in runs_page if PCT_RE.match(r["text"])]
    cand = []
    for i, lab in enumerate(labels):
        for k, p in enumerate(pcts):
            dy = p["top"] - lab["top"]
            dx = abs(_cx(p) - _cx(lab))
            if 0 < dy <= 12 and dx <= 20:
                cand.append((dy + dx * 0.8, i, k))
    cand.sort()
    li, pk, pares = set(), set(), []
    for _, i, k in cand:
        if i in li or k in pk:
            continue
        li.add(i)
        pk.add(k)
        val = float(PCT_RE.match(pcts[k]["text"]).group(1).replace(",", "."))
        pares.append({**labels[i], "val": val})
    return pares


def _asignar_a_tortas(pares, charts):
    """Asigna cada par (etiqueta, %) a la torta correspondiente."""
    asign = {c["id"]: [] for c in charts}
    ambiguos = []  # (par, [cand1, cand2])
    for p in pares:
        arriba = [c for c in charts if c["top"] < p["top"]]
        if not arriba:
            continue
        fila_top = max(c["top"] for c in arriba)
        fila = [c for c in arriba if fila_top - c["top"] <= 6]
        fila.sort(key=lambda c: abs(c["cx"] - _cx(p)))
        if len(fila) >= 2:
            d1, d2 = abs(fila[0]["cx"] - _cx(p)), abs(fila[1]["cx"] - _cx(p))
            if d2 > 0 and d1 / d2 > 0.55:
                ambiguos.append((p, fila[:2]))
                continue
        asign[fila[0]["id"]].append(p)

    # Resolver ambiguos buscando la combinación cuyas sumas queden más cerca de 100
    if ambiguos:
        if len(ambiguos) > 10:  # no debería pasar; asignar al más cercano
            for p, cands in ambiguos:
                asign[cands[0]["id"]].append(p)
        else:
            best, best_err = None, None
            for combo in itertools.product((0, 1), repeat=len(ambiguos)):
                trial = {k: list(v) for k, v in asign.items()}
                for (p, cands), ch in zip(ambiguos, combo):
                    trial[cands[ch]["id"]].append(p)
                err = 0.0
                for cid, ps in trial.items():
                    cats = [x["key"] for x in ps]
                    dup = len(cats) - len(set(cats))
                    err += abs(sum(x["val"] for x in ps) - 100) + dup * 100
                if best_err is None or err < best_err - 1e-9:
                    best, best_err = trial, err
            asign = best
    return asign


def parse_tenencias(pdf_bytes: bytes = None, runs: list = None) -> dict:
    if runs is None:
        runs = extraer_runs(pdf_bytes)
    anio, mes = parse_periodo(runs)

    bonos = {}
    seccion = "nominal"
    total_mn_nominal = None
    for pno in sorted({r["page"] for r in runs}):
        rp = [r for r in runs if r["page"] == pno]
        ptxt = _norm(" ".join(r["text"] for r in rp))
        if "moneda nacional indexada" in ptxt and "nominal e indexada" not in ptxt:
            seccion = "indexada"
        elif "tenencias por bonos en moneda nacional nominal" in ptxt and "e indexada" not in ptxt:
            seccion = "nominal"

        # Total MN nominal (pág. resumen) para control cruzado
        for r in rp:
            if total_mn_nominal is None and "moneda nacional nominal" in _norm(r["text"]) \
                    and r["text"].lower().startswith("bonos soberanos"):
                below = [s for s in rp if 0 < s["top"] - r["top"] <= 12 and UNITS_RE.search(s["text"])]
                below.sort(key=lambda s: abs(s["x0"] - r["x0"]))
                if below:
                    total_mn_nominal = int(re.sub(r"\D", "", UNITS_RE.search(below[0]["text"]).group(1)))

        # Títulos de torta por bono
        charts = []
        for r in rp:
            m = BONO_RE.search(r["text"])
            if not m:
                continue
            bid = m.group(1)
            units = None
            below = [s for s in rp if 0 < s["top"] - r["top"] <= 12 and UNITS_RE.search(s["text"])
                     and not BONO_RE.search(s["text"])]
            below.sort(key=lambda s: abs(_cx(s) - _cx(r)))
            if below:
                units = int(re.sub(r"\D", "", UNITS_RE.search(below[0]["text"]).group(1)))
            charts.append({"id": bid, "top": r["top"], "cx": _cx(r), "units": units})
        if not charts:
            continue

        pares = _emparejar_valores(_etiquetas(rp), rp)
        asign = _asignar_a_tortas(pares, charts)
        for c in charts:
            pct = {k: 0.0 for k in CATS}
            detalle = {}
            for p in asign[c["id"]]:
                pct[p["cat"]] += p["val"]
                detalle[p["key"]] = p["val"]
            bonos[c["id"]] = {
                "seccion": seccion, "unidades": c["units"],
                "pct": {k: round(v, 4) for k, v in pct.items()},
                "detalle": detalle, "suma": round(sum(detalle.values()), 2),
            }

    return {"anio": anio, "mes": mes, "bonos": bonos, "total_mn_nominal_unidades": total_mn_nominal}


def validar(parsed: dict) -> list[str]:
    """Devuelve lista de errores. Lista vacía = OK."""
    errs = []
    nominales = {k: v for k, v in parsed["bonos"].items() if v["seccion"] == "nominal"}
    if len(nominales) < 8:
        errs.append(f"Solo se detectaron {len(nominales)} bonos nominales (esperado ≥ 8)")
    for bid, b in nominales.items():
        if b["unidades"] is None:
            errs.append(f"{bid}: no se encontraron las unidades en circulación")
        if abs(b["suma"] - 100) > 1.0:
            errs.append(f"{bid}: los % suman {b['suma']} (esperado ~100): {b['detalle']}")
    tot = parsed.get("total_mn_nominal_unidades")
    s = sum((b["unidades"] or 0) for b in nominales.values())
    if tot and abs(s - tot) / tot > 0.001:
        errs.append(f"Suma de unidades por bono ({s:,}) ≠ total MN nominal del PDF ({tot:,})")
    return errs


# ─────────────────────────────────────────────────────────────────
# 4. Duración modificada (para DV01)
# ─────────────────────────────────────────────────────────────────
def _add_months(d: date, n: int) -> date:
    y, m = divmod(d.month - 1 + n, 12)
    return date(d.year + y, m + 1, d.day)


def duracion_modificada(bono_id: str, cupon: float, ytm: float, settle: date) -> float:
    """Duración modificada (años) de un bono bullet con cupón semestral."""
    mat = bono_fecha(bono_id)
    if mat <= settle:
        return 0.0
    flujos = []
    d = mat
    while d > settle:
        flujos.append(d)
        d = _add_months(d, -6)
    flujos.reverse()
    c, y = cupon / 2, ytm / 100 / 2
    pv = pvt = 0.0
    for i, fd in enumerate(flujos):
        t = (fd - settle).days / 365.25 * 2  # en semestres
        cf = c + (100 if i == len(flujos) - 1 else 0)
        df = (1 + y) ** (-t)
        pv += cf * df
        pvt += t * cf * df
    mac = pvt / pv / 2
    return mac / (1 + y)


# ─────────────────────────────────────────────────────────────────
# 5. Construir el JSON del dashboard
# ─────────────────────────────────────────────────────────────────
def periodo_key(anio, mes):
    return f"{MES_ABR_EN[mes - 1]}-{anio}"


def construir(parsed: dict, anterior: dict | None, url: str) -> dict:
    anio, mes = parsed["anio"], parsed["mes"]
    nominales = {k: v for k, v in parsed["bonos"].items() if v["seccion"] == "nominal"}

    # Agrupar por SOB (12FEB2029 + 12FEB2029E → SOB29), ordenados por vencimiento
    grupos: dict[str, list[str]] = {}
    for bid in sorted(nominales, key=bono_fecha):
        grupos.setdefault(bono_sob(bid), []).append(bid)

    outstanding, own_pct, tenores = [], {}, []
    for sob, ids in grupos.items():
        tenor = bono_tenor(ids[0])
        tenores.append(tenor)
        mm = {k: 0.0 for k in CATS}
        units = 0
        for bid in ids:
            b = nominales[bid]
            u = b["unidades"] or 0
            units += u
            for k in CATS:
                mm[k] += b["pct"][k] / 100 * u / 1000  # 1 unidad = S/ 1,000 → MM PEN
        total = units / 1000
        row = {"bono": sob, "tenor": tenor, "vencimiento": " / ".join(ids)}
        for k in CATS:
            row[k] = round(mm[k])
        row["PFs"] = row["Pension Funds"]
        row["TOTAL"] = round(total)
        row["_mm"] = mm  # precisión completa (se elimina antes de guardar)
        outstanding.append(row)
        own_pct[tenor] = {k: round(mm[k] / total * 100, 1) if total else 0 for k in CATS}
        own_pct[tenor]["PFs"] = own_pct[tenor]["Pension Funds"]

    total_mn = sum(r["_mm"][k] for r in outstanding for k in CATS)
    total_units = sum(r["TOTAL"] for r in outstanding)
    tipo = {k: round(sum(r["_mm"][k] for r in outstanding) / total_mn * 100, 2) for k in CATS}

    # MoM y DV01
    settle = _add_months(date(anio, mes, 1), 1)  # ~ fin de mes
    yields = {}
    if YIELDS_FILE.exists():
        yields = json.loads(YIELDS_FILE.read_text(encoding="utf-8"))
    prev_rows = {r["bono"]: r for r in (anterior or {}).get("outstanding_por_bono", [])}
    prev_per = (anterior or {}).get("meta", {}).get("periodo")
    all_tenors = list(dict.fromkeys(
        [r.get("tenor") or bono_tenor(r["vencimiento"].split(" / ")[0]) for r in prev_rows.values()] + tenores))
    all_tenors.sort()
    por_inv, dv01 = {}, {"nota": ""}
    mom_label = {"Offshores": "Offshores", "Pension Funds": "PFs", "Banks": "Banks",
                 "Insurance": "Insurance", "Public Funds": "Public Funds", "Others": "Others"}
    act_rows = {r["bono"]: r for r in outstanding}
    for k in CATS:
        fila, filad = {}, {}
        for sob in dict.fromkeys(list(prev_rows) + list(act_rows)):
            a = act_rows.get(sob, {}).get(k, 0)
            p = prev_rows.get(sob, {}).get(k, 0)
            ten = (act_rows.get(sob) or prev_rows.get(sob)).get("tenor") or ("" + sob[3:] + "s")
            fila[ten] = a - p
            if k in DV01_CATS and sob in act_rows:
                bid = act_rows[sob]["vencimiento"].split(" / ")[-1]
                cup = CUPONES.get(bid, 6.5)
                ytm = float(yields.get(sob, yields.get(bid, cup)))
                md = duracion_modificada(bid, cup, ytm, settle)
                filad[ten] = round((a - p) * md * 0.1)  # K PEN por pb
        fila["Total"] = sum(v for t, v in fila.items())
        por_inv[mom_label[k]] = fila
        if k in DV01_CATS:
            dv01[k] = filad
    fuente_y = "yields de data/yields.json" if yields else "yield = cupón (aprox. a la par)"
    dv01["nota"] = f"K PEN por pb = ΔMM PEN × duración modificada × 0.1 ({fuente_y})"

    mom_nota = f"Cambios en MM PEN MoM ({periodo_key(anio, mes)} vs {prev_per or 's/d'})"
    if not anterior:
        por_inv = {}

    data = {
        "meta": {
            "periodo": periodo_key(anio, mes),
            "periodo_label": f"{MESES[mes - 1].capitalize()} {anio}",
            "periodo_anterior": prev_per,
            "fecha_reporte": datetime.now().strftime("%Y-%m-%d"),
            "fuente": "Ministerio de Economía y Finanzas del Perú (MEF)",
            "url_tenencias": url,
            "total_nominal_mn": total_units,
            "parser": "v3-coordenadas",
        },
        "tenores": tenores,
        "tenencias_por_tipo": tipo,
        "outstanding_por_bono": outstanding,
        "cambios_mom": {
            "nota": mom_nota,
            "nota2": "Calculado a partir de las tenencias MEF (% por bono × unidades en circulación). "
                     "Afectado por REPOs AFP-BCRP si los hubiera.",
            "por_inversor": por_inv,
        },
        "ownership_pct_por_tenor": own_pct,
        "evolucion_ownership": [],  # se completa en construir_evolucion()
        "dv01_mom": dv01,
        "detalle_pdf": {bid: {"unidades": b["unidades"], "pct": b["detalle"], "suma": b["suma"]}
                        for bid, b in parsed["bonos"].items()},
    }
    for r in outstanding:
        r.pop("_mm")
    return data


def _punto_evolucion(d: dict) -> dict:
    anio, mes = int(d["meta"]["periodo"][-4:]), MES_ABR_EN.index(d["meta"]["periodo"][:3]) + 1
    t = d["tenencias_por_tipo"]
    return {
        "fecha": f"{MES_ABR_EN[mes - 1]}-{str(anio)[2:]}", "_orden": anio * 100 + mes,
        "Offshores": round(t["Offshores"], 2),
        "Pension Funds": round(t["Pension Funds"], 2), "PFs": round(t["Pension Funds"], 2),
        "Banks": round(t["Banks"], 2),
        "Insurance": round(t["Insurance"], 2),
        "Others": round(t["Public Funds"] + t["Others"], 2),  # igual que la serie histórica
    }


def construir_evolucion(actual: dict) -> list:
    """Serie = histórico manual (hasta el primer mes parseado) + un punto por cada data_*.json."""
    puntos = {}
    for f in sorted(DATA_DIR.glob("data_*.json")):
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
            if d.get("meta", {}).get("parser", "").startswith("v3"):
                p = _punto_evolucion(d)
                puntos[p["_orden"]] = p
        except Exception:
            pass
    p = _punto_evolucion(actual)
    puntos[p["_orden"]] = p
    primero = min(puntos)
    hist = []
    if HIST_FILE.exists():
        for h in json.loads(HIST_FILE.read_text(encoding="utf-8")):
            mm, yy = h["fecha"].split("-")
            nmes = (MES_ABR_EN.index(mm) if mm in MES_ABR_EN else MES_ABR_ES.index(mm)) + 1
            orden = (2000 + int(yy)) * 100 + nmes
            h = {**h, "fecha": f"{MES_ABR_EN[nmes - 1]}-{yy}"}
            if orden < primero:
                h = dict(h)
                h.setdefault("Pension Funds", h.get("PFs"))
                h["_orden"] = orden
                hist.append(h)
    serie = sorted(hist + list(puntos.values()), key=lambda x: x["_orden"])
    for s in serie:
        s.pop("_orden", None)
    return serie


# ─────────────────────────────────────────────────────────────────
# 6. Guardar
# ─────────────────────────────────────────────────────────────────
def slug(periodo: str) -> str:
    return periodo.replace("-", "_").lower()


def cargar_periodo(periodo: str) -> dict | None:
    p = DATA_DIR / f"data_{slug(periodo)}.json"
    if p.exists():
        d = json.loads(p.read_text(encoding="utf-8"))
        if d.get("meta", {}).get("parser", "").startswith("v3"):
            return d
    return None


def periodo_previo(anio, mes) -> str:
    return periodo_key(anio - 1, 12) if mes == 1 else periodo_key(anio, mes - 1)


def guardar(data: dict, es_ultimo: bool):
    DATA_DIR.mkdir(exist_ok=True)
    txt = json.dumps(data, ensure_ascii=False, indent=2)
    (DATA_DIR / f"data_{slug(data['meta']['periodo'])}.json").write_text(txt, encoding="utf-8")
    if es_ultimo:
        (DATA_DIR / "latest.json").write_text(txt, encoding="utf-8")
        actualizar_html(data)
    print(f"  ✓ Guardado data_{slug(data['meta']['periodo'])}.json" + (" + latest.json + index.html" if es_ultimo else ""))


def actualizar_html(data: dict):
    if not INDEX_HTML.exists():
        return
    html = INDEX_HTML.read_text(encoding="utf-8")
    embedded = {k: v for k, v in data.items() if k != "detalle_pdf"}
    nuevo = json.dumps(embedded, ensure_ascii=False, indent=2)
    pat = re.compile(r"(const EMBEDDED_DATA\s*=\s*)(\{.*?\n\});", re.S)
    if not pat.search(html):
        raise RuntimeError("No se encontró 'const EMBEDDED_DATA = {...};' en index.html")
    INDEX_HTML.write_text(pat.sub(lambda m: m.group(1) + nuevo + ";", html, count=1), encoding="utf-8")


# ─────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────
def procesar(pdf_bytes: bytes, url: str, es_ultimo: bool, anterior: dict | None = None) -> dict:
    parsed = parse_tenencias(pdf_bytes)
    per = periodo_key(parsed["anio"], parsed["mes"])
    print(f"  Período: {per} — {len(parsed['bonos'])} tortas detectadas")
    errs = validar(parsed)
    if errs:
        print("  ✗ Validación fallida:")
        for e in errs:
            print("    -", e)
        raise SystemExit(2)
    if anterior is None:
        anterior = cargar_periodo(periodo_previo(parsed["anio"], parsed["mes"]))
    if anterior is None:
        print("  [WARN] No hay datos v3 del mes anterior → MoM/DV01 vacíos")
    data = construir(parsed, anterior, url)
    data["evolucion_ownership"] = construir_evolucion(data)
    guardar(data, es_ultimo)
    return data


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--force", action="store_true", help="re-procesar aunque el período ya exista")
    ap.add_argument("--backfill", type=int, default=0, help="re-procesar los N últimos reportes en orden")
    ap.add_argument("--pdf", help="usar un PDF local en vez de descargar")
    args = ap.parse_args()

    print("=" * 60)
    print("Soberanos Dashboard — actualización", datetime.now().strftime("%Y-%m-%d %H:%M"))
    print("=" * 60)

    if args.pdf:
        procesar(Path(args.pdf).read_bytes(), args.pdf, es_ultimo=True)
        return

    reportes = listar_reportes()
    print(f"  {len(reportes)} reportes en el MEF; último: {reportes[-1][1]}")

    if args.backfill:
        sel = reportes[-args.backfill:]
        anterior = None
        for i, (d, url) in enumerate(sel):
            print(f"\n→ {url}")
            anterior = procesar(_get(url).content, url, es_ultimo=(i == len(sel) - 1),
                                anterior=anterior)
        return

    d, url = reportes[-1]
    latest = DATA_DIR / "latest.json"
    if latest.exists() and not args.force:
        meta = json.loads(latest.read_text(encoding="utf-8")).get("meta", {})
        if meta.get("url_tenencias") == url and meta.get("parser", "").startswith("v3"):
            print("  Sin reporte nuevo — nada que hacer.")
            return
    print(f"\n→ {url}")
    procesar(_get(url).content, url, es_ultimo=True)


if __name__ == "__main__":
    main()
