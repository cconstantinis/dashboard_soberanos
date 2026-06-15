"""
fetch_mef_data.py
=================
Descarga y parsea automáticamente los dos PDFs del MEF:
  1. Tenencias de Bonos Soberanos  → % por inversor por bono + totales
  2. Stock de Bonos Soberanos      → unidades en circulación por bono

Cruza ambas fuentes para calcular MM PEN por bono × inversor,
calcula cambios MoM respecto al mes anterior, y reescribe
EMBEDDED_DATA en index.html.

Uso:
    pip install pdfplumber requests beautifulsoup4
    python scripts/fetch_mef_data.py

GitHub Actions lo corre el día 5 de cada mes automáticamente.
"""

import json
import re
import sys
import io
import copy
from datetime import datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup

# ─────────────────────────────────────────────────────────────────
# Configuración
# ─────────────────────────────────────────────────────────────────
ROOT        = Path(__file__).parent.parent
DATA_DIR    = ROOT / "data"
INDEX_HTML  = ROOT / "index.html"
DATA_DIR.mkdir(exist_ok=True)

TENENCIAS_URL = "https://www.mef.gob.pe/contenidos/deuda_publ/mercado/reportes_tenencia.php"
STOCK_URL     = "https://www.mef.gob.pe/contenidos/deuda_publ/bonos/internos/stock_bonos_soberanos.php"

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; SoberanosDashboard/2.0)"}

# Mapeo de nombres del MEF → claves del dashboard
# Nota: "Personas naturales" puede aparecer dividido en 2 líneas en el PDF.
# El parser usa un patrón especial para este caso (ver abajo).
INVERSOR_MAP = {
    "No residentes":    "Offshores",
    "AFPs":             "Pension Funds",
    "Bancos":           "Banks",
    "Seguros":          "Insurance",
    "Fondos públicos":  "Public Funds",
    "Fondos privados":  "Fondos Privados",
    "Otros":            "Others",
    "Personas naturales": "Personas Naturales",
}

# Patrones de regex por inversor (para manejar saltos de línea internos)
INVERSOR_PATTERNS = {
    "No residentes":    r"No residentes",
    "AFPs":             r"AFPs",
    "Bancos":           r"Bancos",
    "Seguros":          r"Seguros",
    "Fondos públicos":  r"Fondos p[úu]blicos",
    "Fondos privados":  r"Fondos privados",
    "Otros":            r"Otros",
    "Personas naturales": r"Personas\s*\n?\s*naturales",
}

# Bonos MN nominal que nos interesan (en orden)
BONOS_OBJETIVO = [
    "12AGO2026", "12AGO2028", "12FEB2029", "12FEB2029E",
    "12AGO2031", "12AGO2032", "12AGO2033", "12AGO2034",
    "12AGO2035", "12AGO2037", "12AGO2039", "12AGO2040",
    "12FEB2042", "12FEB2055",
]

BONO_A_SOB = {
    "12AGO2026": "SOB26", "12AGO2028": "SOB28",
    "12FEB2029": "SOB29", "12FEB2029E": "SOB29",
    "12AGO2031": "SOB31", "12AGO2032": "SOB32",
    "12AGO2033": "SOB33", "12AGO2034": "SOB34",
    "12AGO2035": "SOB35", "12AGO2037": "SOB37",
    "12AGO2039": "SOB39", "12AGO2040": "SOB40",
    "12FEB2042": "SOB42", "12FEB2055": "SOB55",
}

TENOR_MAP = {
    "SOB26":"26s","SOB28":"28s","SOB29":"29s","SOB31":"31s",
    "SOB32":"32s","SOB33":"33s","SOB34":"34s","SOB35":"35s",
    "SOB37":"37s","SOB39":"39s","SOB40":"40s","SOB42":"42s","SOB55":"55s",
}

MESES_ES = {
    "enero":"Jan","febrero":"Feb","marzo":"Mar","abril":"Apr",
    "mayo":"May","junio":"Jun","julio":"Jul","agosto":"Aug",
    "septiembre":"Sep","octubre":"Oct","noviembre":"Nov","diciembre":"Dec",
}

# ─────────────────────────────────────────────────────────────────
# 1. Utilidades HTTP
# ─────────────────────────────────────────────────────────────────

def get(url, **kwargs):
    r = requests.get(url, headers=HEADERS, timeout=60, **kwargs)
    r.raise_for_status()
    return r


# ─────────────────────────────────────────────────────────────────
# 2. Encontrar PDFs más recientes
# ─────────────────────────────────────────────────────────────────

def latest_tenencias_url():
    """Devuelve URL del PDF de tenencias más reciente del año en curso."""
    soup = BeautifulSoup(get(TENENCIAS_URL).text, "html.parser")
    links = soup.find_all("a", href=re.compile(r"tenencia_bono_\d+\.pdf", re.I))
    if not links:
        raise RuntimeError("No se encontraron PDFs de tenencias")
    href = links[-1]["href"]
    if not href.startswith("http"):
        href = "https://www.mef.gob.pe" + href
    print(f"  Tenencias PDF: {href}")
    return href


def latest_stock_url():
    """Devuelve URL del PDF de stock más reciente."""
    soup = BeautifulSoup(get(STOCK_URL).text, "html.parser")
    links = soup.find_all("a", href=re.compile(r"Stock_bonos_soberanos_\d+\.pdf", re.I))
    if not links:
        raise RuntimeError("No se encontraron PDFs de stock")
    href = links[-1]["href"]
    if not href.startswith("http"):
        href = "https://www.mef.gob.pe" + href
    print(f"  Stock PDF:     {href}")
    return href


# ─────────────────────────────────────────────────────────────────
# 3. Parsear PDF de TENENCIAS
# ─────────────────────────────────────────────────────────────────

def _words_to_column_text(words, x_split):
    """
    Dado una lista de palabras con coordenadas pdfplumber,
    separa en columna izquierda (x1 < x_split) y derecha (x1 >= x_split).
    Dentro de cada columna reconstruye el texto ordenado por Y luego X.
    Retorna (texto_izq, texto_der).
    """
    left  = [w for w in words if w["x1"] <= x_split]
    right = [w for w in words if w["x1"] >  x_split]

    def reconstruct(ws):
        if not ws:
            return ""
        ws_sorted = sorted(ws, key=lambda w: (round(w["top"] / 5) * 5, w["x0"]))
        lines = []
        cur_y, cur_line = None, []
        for w in ws_sorted:
            y = round(w["top"] / 5) * 5
            if cur_y is None or abs(y - cur_y) > 4:
                if cur_line:
                    lines.append(" ".join(cur_line))
                cur_line = [w["text"]]
                cur_y = y
            else:
                cur_line.append(w["text"])
        if cur_line:
            lines.append(" ".join(cur_line))
        return "\n".join(lines)

    return reconstruct(left), reconstruct(right)


def _extract_pct_from_column(col_text):
    """
    Dado el texto de una columna de bono, extrae {key: pct} usando
    los nombres del INVERSOR_MAP. Acepta "Nombre\nN.NN%" o "Nombre N.NN%".
    """
    inv_pct = {}
    for nombre_mef, key in INVERSOR_MAP.items():
        pat = INVERSOR_PATTERNS[nombre_mef] + r"[\s\n]+([\d.]+)%"
        m = re.search(pat, col_text, re.I)
        if m:
            inv_pct[key] = float(m.group(1))
    return inv_pct


def parse_tenencias(pdf_bytes: bytes) -> dict:
    """
    Extrae del PDF de tenencias del MEF usando coordenadas de palabras
    para manejar el layout de 2 columnas por página.
    """
    import pdfplumber

    result = {}
    pct_por_bono   = {}
    tenencias_glob = {}
    periodo_found  = False
    total_units    = 0

    # Nombre de bono → regex para reconocerlo en el texto
    bono_re = re.compile(r"(1[12][A-Z]+\d{4}E?)")

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        full_text = "\n".join(p.extract_text() or "" for p in pdf.pages)

        for page in pdf.pages:
            page_text = page.extract_text() or ""
            words     = page.extract_words(keep_blank_chars=False,
                                           extra_attrs=["x0","x1","top","bottom"])

            # ── Período ────────────────────────────────────────────
            if not periodo_found:
                m = re.search(
                    r"(enero|febrero|marzo|abril|mayo|junio|julio|agosto|"
                    r"septiembre|octubre|noviembre|diciembre)\s+de\s+(\d{4})",
                    page_text, re.I
                )
                if m:
                    mes_en = MESES_ES.get(m.group(1).lower(), m.group(1)[:3].capitalize())
                    result["periodo"]       = f"{mes_en}-{m.group(2)}"
                    result["periodo_label"] = f"{m.group(1).capitalize()} {m.group(2)}"
                    periodo_found = True

            # ── Total MN Nominal ───────────────────────────────────
            if total_units == 0:
                # Buscar "XXX XXX XXX Unidades" con número de 8-9 dígitos
                for pat in [
                    r"Nacional Nominal\s*\n?([\d][\d ]{7,11})\s*Unidades",
                    r"(1[5-9]\d[\d ]{6,9})\s*Unidades",
                ]:
                    m2 = re.search(pat, page_text)
                    if m2:
                        raw = m2.group(1).replace(" ", "")
                        if raw.isdigit() and len(raw) >= 8:
                            total_units = int(raw)
                            break

            # ── % globales (página resumen, solo primera página con "MN Nominal") ──
            if not tenencias_glob and "Moneda Nacional Nominal" in page_text:
                # Ancho de página → columna izquierda es toda la página (layout circular/radial)
                # Extraer todos los % de la página y mapear con los nombres
                for nombre_mef, key in INVERSOR_MAP.items():
                    pat = INVERSOR_PATTERNS[nombre_mef] + r"[\s\n ]+([\d.]+)%"
                    m3 = re.search(pat, page_text, re.I)
                    if m3:
                        tenencias_glob[key] = float(m3.group(1))

            # ── % por bono (páginas de detalle con 2 columnas) ────
            # Detectar si hay nombres de bonos en esta página
            bonos_en_pagina = bono_re.findall(page_text)
            bonos_en_pagina = [b for b in bonos_en_pagina if b in BONOS_OBJETIVO]
            if not bonos_en_pagina:
                continue

            # Ancho de página para dividir columnas
            page_w = float(page.width)
            mid_x  = page_w / 2

            if len(bonos_en_pagina) == 1:
                # Una sola columna (o bono único en página)
                col_text = page_text
                inv_pct  = _extract_pct_from_column(col_text)
                if inv_pct:
                    pct_por_bono[bonos_en_pagina[0]] = inv_pct
            else:
                # 2 columnas: dividir palabras por X
                col_left, col_right = _words_to_column_text(words, mid_x)

                # Averiguar qué bono va en cada columna buscando en cada texto
                left_bonos  = [b for b in bonos_en_pagina if b in col_left]
                right_bonos = [b for b in bonos_en_pagina if b in col_right]

                # Si la separación no funciona, buscar el bono cuyo nombre
                # aparece primero en el texto plano de cada columna
                if not left_bonos and not right_bonos:
                    left_bonos  = bonos_en_pagina[:1]
                    right_bonos = bonos_en_pagina[1:]

                for bono in left_bonos:
                    inv_pct = _extract_pct_from_column(col_left)
                    if inv_pct:
                        pct_por_bono[bono] = inv_pct

                for bono in right_bonos:
                    inv_pct = _extract_pct_from_column(col_right)
                    if inv_pct:
                        pct_por_bono[bono] = inv_pct

    # ── Fallbacks ─────────────────────────────────────────────────
    if not periodo_found:
        result["periodo"]       = datetime.now().strftime("%b-%Y")
        result["periodo_label"] = result["periodo"]

    result["total_nominal_mn"] = round(total_units / 1000) if total_units else 0
    result["tenencias_por_tipo"] = tenencias_glob

    if tenencias_glob:
        print(f"  tenencias_por_tipo: {tenencias_glob}")
    else:
        print(f"  [WARN] tenencias_por_tipo vacío")

    result["pct_por_bono"] = pct_por_bono
    print(f"  Bonos parseados con %: {list(pct_por_bono.keys())}")

    # ── Evolución histórica ───────────────────────────────────────
    result["evolucion"] = _parse_evolucion(full_text)

    return result


def _parse_evolucion(text: str) -> list:
    """
    Extrae la serie histórica de ownership del gráfico de evolución.
    El PDF siempre muestra los últimos 12 meses.
    Devuelve lista de dicts {fecha, Offshores, PFs, Banks, Insurance, Others}.
    """
    # Buscar bloque con fechas tipo "May-25 Jun-25 ..."
    fecha_pattern = r"((?:(?:Ene|Feb|Mar|Abr|May|Jun|Jul|Ago|Set|Oct|Nov|Dic)-\d{2}\s*)+)"
    m = re.search(fecha_pattern, text)
    if not m:
        return []

    fechas_raw = m.group(1).strip().split()
    # Normalizar fechas
    mes_map = {"Ene":"Jan","Feb":"Feb","Mar":"Mar","Abr":"Apr","May":"May",
               "Jun":"Jun","Jul":"Jul","Ago":"Aug","Set":"Sep","Oct":"Oct",
               "Nov":"Nov","Dic":"Dec"}
    fechas = []
    for f in fechas_raw:
        parts = f.split("-")
        if len(parts) == 2:
            mes_en = mes_map.get(parts[0], parts[0])
            fechas.append(f"{mes_en}-{parts[1]}")

    if not fechas:
        return []

    n = len(fechas)

    # Buscar los bloques de % flotantes antes del bloque de fechas
    # El texto antes del bloque de fechas tiene 4 series de n valores:
    # AFPs%, Bancos%, No residentes%, y quizás Seguros%
    before = text[:m.start()]

    # Extraer todos los números flotantes del bloque previo (últimos n*4 aprox)
    nums = re.findall(r"(\d{1,2}\.\d{1,2})%", before)
    nums = [float(x) for x in nums]

    # Las últimas 3*n son AFPs, Bancos, No residentes (en ese orden en el gráfico)
    # Validamos que tengamos suficientes
    if len(nums) < n * 3:
        return []

    afps    = nums[-3*n : -2*n]
    bancos  = nums[-2*n : -n]
    nores   = nums[-n:]

    evol = []
    for i, fecha in enumerate(fechas):
        evol.append({
            "fecha":    fecha,
            "Offshores": round(nores[i], 2)  if i < len(nores)  else 0,
            "PFs":       round(afps[i], 2)   if i < len(afps)   else 0,
            "Banks":     round(bancos[i], 2) if i < len(bancos) else 0,
            "Insurance": 0,  # No siempre está en la serie extraíble
            "Others":    0,
        })

    return evol


# ─────────────────────────────────────────────────────────────────
# 4. Parsear PDF de STOCK
# ─────────────────────────────────────────────────────────────────

def parse_stock(pdf_bytes: bytes) -> dict:
    """
    Extrae del PDF de stock:
      - stock_por_bono: { "12AGO2026": 1445972, ... }  (unidades)
      - coupon_por_bono: { "12AGO2026": 8.20, ... }
      - venc_por_bono: { "12AGO2026": "2026-08-12", ... }
    """
    import pdfplumber

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages)

    stock = {}
    cupones = {}
    vencimientos = {}

    # Patrón: "Bonos Soberanos 12AGO2026 ... PEP... SB... 1,445,972 1,000.0 ... 8.2000%"
    # El PDF tiene líneas como:
    # "Bonos Soberanos 12AGO2026 3/ 5/ 7/ 9/ PEP01000C0J9 SB12AGO26 1,445,972 ..."
    bono_line = re.compile(
        r"Bonos Soberanos\s+(1\d[A-Z]+\d{4}E?)"  # nombre bono
        r"(?:\s+[\d/]+)*"                          # notas opcionales
        r"\s+PEP\w+"                               # ISIN
        r"\s+SB\w+"                                # nemónico
        r"\s+([\d,]+)"                             # unidades en circulación
        r"\s+[\d,.]+\s+[\d,.]+\s+[\d,.]+"         # valor nominal x3
        r"\s+[\d.]+\s+"                            # plazo
        r"([\d.]+)%"                               # tasa cupón
    )

    for m in bono_line.finditer(text):
        bono_id = m.group(1)
        units   = int(m.group(2).replace(",", ""))
        cupon   = float(m.group(3))
        stock[bono_id]   = units
        cupones[bono_id] = cupon

    # Fechas de vencimiento: "DD/MM/YYYY" o similar al final de cada línea
    # Más simple: buscar el patrón fecha al final de línea de bono
    date_line = re.compile(
        r"Bonos Soberanos\s+(1\d[A-Z]+\d{4}E?)"
        r".*?"
        r"(\d{1,2}/\d{2}/\d{4})\s+"   # fecha emisión
        r"(\d{1,2}/\d{2}/\d{4})",      # fecha vencimiento
        re.S
    )
    for m in date_line.finditer(text):
        bono_id = m.group(1)
        venc_raw = m.group(3)  # DD/MM/YYYY
        try:
            dt = datetime.strptime(venc_raw, "%d/%m/%Y")
            vencimientos[bono_id] = dt.strftime("%Y-%m-%d")
        except ValueError:
            pass

    return {
        "stock_por_bono": stock,
        "coupon_por_bono": cupones,
        "venc_por_bono": vencimientos,
    }


# ─────────────────────────────────────────────────────────────────
# 5. Cruzar tenencias × stock → MM PEN
# ─────────────────────────────────────────────────────────────────

def calcular_outstanding(pct_por_bono: dict, stock_por_bono: dict) -> list:
    """
    Para cada bono: MM PEN = (% / 100) × stock × 1000 / 1_000_000
    Retorna lista de dicts con columnas del dashboard.
    """
    rows = []
    cols = ["Offshores", "Pension Funds", "Banks", "Insurance", "Public Funds", "Others"]

    for bono_id in BONOS_OBJETIVO:
        sob = BONO_A_SOB.get(bono_id)
        if not sob:
            continue

        pct  = pct_por_bono.get(bono_id, {})
        stk  = stock_por_bono.get(bono_id, 0)  # unidades

        row = {"bono": sob, "vencimiento": bono_id}
        total = 0
        for col in cols:
            pf = pct.get(col, 0)
            # 1 unidad = S/1,000 → stk × 1000 / 1e6 = stk / 1000 (en MM PEN)
            mm = round(pf / 100 * stk / 1000)
            row[col] = mm
            total += mm
        # PFs = Pension Funds alias
        row["PFs"] = row["Pension Funds"]
        row["TOTAL"] = round(stk / 1000)  # total real del bono
        rows.append(row)

    return rows


# ─────────────────────────────────────────────────────────────────
# 6. Calcular cambios MoM
# ─────────────────────────────────────────────────────────────────

def calcular_mom(outstanding_actual: list, outstanding_anterior: list) -> dict:
    """Resta outstanding anterior del actual, agrupado por inversor y tenor."""

    def indexar(rows):
        return {r["bono"]: r for r in rows}

    actual_idx   = indexar(outstanding_actual)
    anterior_idx = indexar(outstanding_anterior)

    cols = ["Offshores", "Pension Funds", "Banks", "Insurance", "Public Funds", "Others"]
    col_labels = {
        "Offshores": "Offshores", "Pension Funds": "PFs",
        "Banks": "Banks", "Insurance": "Insurance",
        "Public Funds": "Public Funds", "Others": "Others",
    }
    tenors = ["SOB26","SOB28","SOB29","SOB31","SOB32","SOB33","SOB34",
              "SOB35","SOB37","SOB39","SOB40","SOB42","SOB55"]

    por_inversor = {}
    for col in cols:
        label = col_labels[col]
        row = {}
        total = 0
        for sob in tenors:
            t = TENOR_MAP.get(sob)
            if not t:
                continue
            v_act = actual_idx.get(sob, {}).get(col, 0)
            v_ant = anterior_idx.get(sob, {}).get(col, 0)
            diff = v_act - v_ant
            row[t] = diff
            total += diff
        row["Total"] = total
        por_inversor[label] = row

    return {
        "nota":  "Cambios en MM PEN MoM",
        "nota2": "Calculado automáticamente desde datos MEF",
        "por_inversor": por_inversor,
    }


# ─────────────────────────────────────────────────────────────────
# 7. Construir ownership_pct_por_tenor
# ─────────────────────────────────────────────────────────────────

def build_ownership_pct(pct_por_bono: dict) -> dict:
    """Reformatea pct_por_bono de {bono_id: {inv: pct}} a {tenor: {inv: pct}}."""
    result = {}
    for bono_id, pct in pct_por_bono.items():
        sob = BONO_A_SOB.get(bono_id)
        tenor = TENOR_MAP.get(sob)
        if not tenor:
            continue
        result[tenor] = {
            "Offshores":    round(pct.get("Offshores", 0)),
            "PFs":          round(pct.get("Pension Funds", 0)),
            "Banks":        round(pct.get("Banks", 0)),
            "Insurance":    round(pct.get("Insurance", 0)),
            "Public Funds": round(pct.get("Public Funds", 0)),
            "Others":       round(pct.get("Others", 0)),
        }
    return result


# ─────────────────────────────────────────────────────────────────
# 8. Cargar datos del mes anterior
# ─────────────────────────────────────────────────────────────────

def cargar_anterior() -> dict:
    """Carga latest.json como referencia del mes anterior."""
    p = DATA_DIR / "latest.json"
    if p.exists():
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    return {}


# ─────────────────────────────────────────────────────────────────
# 9. Armar JSON final
# ─────────────────────────────────────────────────────────────────

def build_json(tenencias: dict, stock: dict,
               anterior: dict, url_ten: str, url_stk: str) -> dict:

    pct_por_bono   = tenencias["pct_por_bono"]
    stock_por_bono = stock["stock_por_bono"]

    outstanding      = calcular_outstanding(pct_por_bono, stock_por_bono)
    outstanding_prev = anterior.get("outstanding_por_bono", outstanding)
    mom              = calcular_mom(outstanding, outstanding_prev)
    ownership_tenor  = build_ownership_pct(pct_por_bono)

    # Si tenencias_por_tipo está vacío o incompleto, derivar del outstanding
    t_tipo = tenencias.get("tenencias_por_tipo", {})
    cols = ["Offshores", "Pension Funds", "Banks", "Insurance", "Public Funds", "Others"]
    if not t_tipo or all(t_tipo.get(c, 0) == 0 for c in cols):
        total_all = sum(r.get("TOTAL", 0) for r in outstanding) or 1
        t_tipo = {}
        for col in cols:
            s = sum(r.get(col, 0) for r in outstanding)
            t_tipo[col] = round(s / total_all * 100, 2)
        print(f"  [INFO] tenencias_por_tipo derivado del outstanding: {t_tipo}")
    tenencias["tenencias_por_tipo"] = t_tipo

    # Evolución: añadir nuevo punto a la serie anterior
    evol_prev = anterior.get("evolucion_ownership", [])
    evol_new  = tenencias.get("evolucion", [])

    # Usar la serie del PDF si la tenemos, sino extender la anterior
    if evol_new:
        evolucion = evol_new
    else:
        # Agregar punto actual a la serie anterior
        t_global = tenencias["tenencias_por_tipo"]
        nuevo_punto = {
            "fecha":     tenencias["periodo"],
            "Offshores": t_global.get("Offshores", 0),
            "PFs":       t_global.get("Pension Funds", 0),
            "Banks":     t_global.get("Banks", 0),
            "Insurance": t_global.get("Insurance", 0),
            "Others":    t_global.get("Others", 0),
        }
        evolucion = evol_prev.copy()
        if not evolucion or evolucion[-1]["fecha"] != nuevo_punto["fecha"]:
            evolucion.append(nuevo_punto)

    return {
        "meta": {
            "periodo":          tenencias["periodo"],
            "periodo_label":    tenencias.get("periodo_label", tenencias["periodo"]),
            "fecha_reporte":    datetime.now().strftime("%Y-%m-%d"),
            "fuente":           "Ministerio de Economía y Finanzas del Perú (MEF)",
            "url_tenencias":    url_ten,
            "url_stock":        url_stk,
            "total_nominal_mn": tenencias["total_nominal_mn"],
        },
        "tenencias_por_tipo":      tenencias["tenencias_por_tipo"],
        "outstanding_por_bono":    outstanding,
        "cambios_mom":             mom,
        "ownership_pct_por_tenor": ownership_tenor,
        "evolucion_ownership":     evolucion,
        "dv01_mom": {
            "nota": "Pendiente calibración con yields reales"
        },
    }


# ─────────────────────────────────────────────────────────────────
# 10. Guardar JSON y actualizar index.html
# ─────────────────────────────────────────────────────────────────

def guardar_json(data: dict) -> Path:
    periodo_slug = data["meta"]["periodo"].replace("-","_").lower()
    out = DATA_DIR / f"data_{periodo_slug}.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    # latest.json
    with open(DATA_DIR / "latest.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"  JSON guardado: {out.name}")
    return out


def actualizar_html(data: dict):
    """Reemplaza el bloque EMBEDDED_DATA en index.html con los datos nuevos."""
    if not INDEX_HTML.exists():
        print("  WARN: index.html no encontrado, saltando actualización HTML")
        return

    html = INDEX_HTML.read_text(encoding="utf-8")

    # Serializar solo los campos que van en EMBEDDED_DATA
    embedded = {k: data[k] for k in [
        "meta", "tenencias_por_tipo", "outstanding_por_bono",
        "cambios_mom", "ownership_pct_por_tenor", "evolucion_ownership",
        "dv01_mom",
    ] if k in data}

    new_data_str = json.dumps(embedded, ensure_ascii=False, indent=2)

    # Reemplazar entre marcadores
    pattern = re.compile(
        r"(const EMBEDDED_DATA\s*=\s*)(\{.*?\});",
        re.S
    )
    if not pattern.search(html):
        print("  WARN: No se encontró EMBEDDED_DATA en index.html")
        return

    new_html = pattern.sub(
        lambda m: m.group(1) + new_data_str + ";",
        html
    )

    INDEX_HTML.write_text(new_html, encoding="utf-8")
    print(f"  index.html actualizado con datos de {data['meta']['periodo']}")


# ─────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────

def main():
    try:
        import pdfplumber  # noqa
    except ImportError:
        print("ERROR: Instala dependencias: pip install pdfplumber requests beautifulsoup4")
        sys.exit(1)

    print("=" * 60)
    print("Soberanos Dashboard — Actualización automática")
    print(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 60)

    # 1. URLs
    print("\n[1/5] Buscando PDFs más recientes...")
    url_ten = latest_tenencias_url()
    url_stk = latest_stock_url()

    # 2. Descargar
    print("\n[2/5] Descargando PDFs...")
    pdf_ten = get(url_ten).content
    pdf_stk = get(url_stk).content
    print(f"  Tenencias: {len(pdf_ten)/1024:.0f} KB")
    print(f"  Stock:     {len(pdf_stk)/1024:.0f} KB")

    # 3. Parsear
    print("\n[3/5] Parseando PDFs...")
    tenencias = parse_tenencias(pdf_ten)
    stock     = parse_stock(pdf_stk)
    print(f"  Período: {tenencias['periodo']}")
    print(f"  Bonos con % parseados: {list(tenencias['pct_por_bono'].keys())}")
    print(f"  Bonos con stock: {list(stock['stock_por_bono'].keys())}")

    # 4. Cargar anterior y construir JSON
    print("\n[4/5] Calculando outstanding y MoM...")
    anterior = cargar_anterior()
    data = build_json(tenencias, stock, anterior, url_ten, url_stk)

    # 5. Guardar
    print("\n[5/5] Guardando archivos...")
    guardar_json(data)
    actualizar_html(data)

    print("\n✓ Listo.")
    print(f"  Período actualizado: {data['meta']['periodo']}")
    print(f"  Total MN Nominal:    {data['meta']['total_nominal_mn']:,} MM PEN")
    bonos_ok = len([b for b in data["outstanding_por_bono"] if b["TOTAL"] > 0])
    print(f"  Bonos procesados:    {bonos_ok}")


if __name__ == "__main__":
    main()
