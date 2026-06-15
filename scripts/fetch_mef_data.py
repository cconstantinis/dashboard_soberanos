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

def parse_tenencias(pdf_bytes: bytes) -> dict:
    """
    Extrae del PDF de tenencias del MEF:
      - periodo, total_nominal_mn, tenencias_por_tipo
      - pct_por_bono: { "12AGO2026": {"Offshores": 7.0, "Banks": 63.42, ...} }
      - evolucion histórica
    """
    import pdfplumber

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        pages = [p.extract_text() or "" for p in pdf.pages]
    text = "\n".join(pages)

    result = {}

    # ── Período ──────────────────────────────────────────────────
    m = re.search(
        r"(enero|febrero|marzo|abril|mayo|junio|julio|agosto|"
        r"septiembre|octubre|noviembre|diciembre)\s+de\s+(\d{4})",
        text, re.I
    )
    if m:
        mes_en = MESES_ES.get(m.group(1).lower(), m.group(1)[:3].capitalize())
        result["periodo"] = f"{mes_en}-{m.group(2)}"
        result["periodo_label"] = f"{m.group(1).capitalize()} {m.group(2)}"
    else:
        result["periodo"] = datetime.now().strftime("%b-%Y")
        result["periodo_label"] = result["periodo"]

    # ── Total MN Nominal ─────────────────────────────────────────
    # El PDF tiene: "184 500 018 Unidades" en varias formas
    # Buscar el número más grande de unidades (total general MN nominal)
    # Aparece como "Bonos Soberanos en Moneda Nacional Nominal\nXXX XXX XXX Unidades"
    for pat in [
        r"Nacional Nominal\s*\n([\d\s]{5,})\s*Unidades",
        r"Nacional Nominal\s+([\d\s]{5,})\s*Unidades",
        r"(1[5-9]\d[\s\d]{5,10})\s*Unidades",   # 1xx,xxx,xxx formato
        r"([\d][\d ]{8,12})\s*Unidades",
    ]:
        m = re.search(pat, text)
        if m:
            raw = m.group(1).replace(" ", "").replace("\n", "")
            if raw.isdigit() and len(raw) >= 8:
                units = int(raw)
                result["total_nominal_mn"] = round(units / 1000)
                break
    else:
        result["total_nominal_mn"] = 0

    # ── % Globales por tipo de inversor (bloque resumen MN Nominal) ──
    # El PDF página 1 tiene el resumen general. Buscamos el bloque
    # entre "Moneda Nacional Nominal" y el siguiente "Moneda Nacional"
    block_mn = re.search(
        r"Moneda Nacional Nominal(.*?)(?:Moneda Nacional Indexada|Corto Plazo)",
        text, re.S
    )
    block_text = block_mn.group(1) if block_mn else text[:3000]

    tenencias = {}
    for nombre_mef, key in INVERSOR_MAP.items():
        pat = INVERSOR_PATTERNS[nombre_mef] + r"[\s\n]+([\d.]+)%"
        m2 = re.search(pat, block_text, re.I)
        if m2:
            tenencias[key] = float(m2.group(1))
    result["tenencias_por_tipo"] = tenencias
    if tenencias:
        print(f"  tenencias_por_tipo: { {k: v for k, v in tenencias.items()} }")
    else:
        print(f"  [WARN] tenencias_por_tipo vacío — block_text snippet: {repr(block_text[:200])}")

    # ── % por inversor por bono ───────────────────────────────────
    # Estrategia robusta: para cada bono, encontrar su posición en el texto,
    # tomar los ~800 chars anteriores y buscar los % de cada inversor ahí.
    # El PDF tiene el texto ANTES del nombre del bono.
    # Ejemplo real:
    #   "Bancos\n63.42%\nFondos públicos\n5.83%\nNo residentes\n7.00%\n
    #    Otros\n17.69%\nPersonas \nnaturales\n0.15%\nSeguros\n5.91%\n
    #    Bonos Soberanos 12AGO2026\n1 445 972 Unidades"

    # Encontrar todas las posiciones de "Bonos Soberanos XXXXX" en el texto
    # Usamos esto para delimitar el bloque de % de cada bono al segmento
    # que va desde el fin del bono anterior hasta el inicio del bono actual.
    all_bono_positions = [(m.start(), m.group(1))
                         for m in re.finditer(r"Bonos Soberanos\s+(1\d[A-Z]+\d{4}E?)", text)]

    pct_por_bono = {}

    for bono in BONOS_OBJETIVO:
        # Encontrar TODAS las ocurrencias del bono en el texto
        positions = [m.start() for m in re.finditer(re.escape(bono), text)]
        if not positions:
            continue

        best = {}
        for pos in positions:
            # Encontrar el "Bonos Soberanos" inmediatamente anterior a esta posición
            # para delimitar el inicio del chunk (evitar contaminar con % del bono anterior)
            prev_bono_end = 0
            for bpos, bname in all_bono_positions:
                if bpos < pos - 5:  # -5 para evitar matchear el bono actual
                    prev_bono_end = bpos
                else:
                    break
            # El chunk va desde después del bono anterior hasta esta posición
            # pero máximo 1200 chars para no ir demasiado lejos
            start = max(prev_bono_end, pos - 1200)
            chunk = text[start:pos]

            inv_pct = {}
            for nombre_mef, key in INVERSOR_MAP.items():
                pat = INVERSOR_PATTERNS[nombre_mef] + r"[\s\n]+([\d.]+)%"
                # Buscar la ÚLTIMA ocurrencia en el chunk (la más cercana al bono)
                matches = list(re.finditer(pat, chunk, re.I))
                if matches:
                    inv_pct[key] = float(matches[-1].group(1))

            # Quedarse con el bloque que tenga más inversores parseados
            if len(inv_pct) > len(best):
                best = inv_pct

        if best:
            pct_por_bono[bono] = best
            # Verificar que los % sumen ~100 (sanity check)
            total_pct = sum(best.values())
            if total_pct < 50 or total_pct > 115:
                print(f"  [WARN] {bono}: suma de % = {total_pct:.1f}% (sospechoso)")
        else:
            if positions:
                pos = positions[0]
                start = max(0, pos - 400)
                snippet = repr(text[start:pos + 50])
                print(f"  [DEBUG] No se parsearon % para {bono}:")
                print(f"    {snippet[:250]}")

    result["pct_por_bono"] = pct_por_bono
    print(f"  Bonos parseados con %: {list(pct_por_bono.keys())}")

    # ── Evolución histórica ───────────────────────────────────────
    evol = _parse_evolucion(text)
    result["evolucion"] = evol

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
