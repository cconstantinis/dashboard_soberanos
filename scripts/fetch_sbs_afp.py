#!/usr/bin/env python3
"""
Cartera administrada de las AFPs (SBS - Boletín Estadístico Semanal del SPP).

Fuente: XLS semanal de la SBS, hojas 'Total Fondo' (Cuadro 8) y 'Fondo0'..'Fondo3'
(Cuadros 9-12). La página índice está protegida por un anti-bot, pero los XLS tienen
una URL predecible y se descargan directo:

  https://intranet2.sbs.gob.pe/estadistica/financiera/{AAAA}/{Mes}/FP-1336-{NN}-{mm}{AAAA}.XLS

Flujo:
  1. Descubre los XLS de los meses que faltan (prueba NN = 01..05).
  2. Parsea cada XLS por ETIQUETA dentro de su sección (no por número de fila) y guarda
     una copia compacta en data/afp/raw/AAAA-MM-DD.json.
  3. Valida: líneas = subtotales, I+II+III = TOTAL, AFPs = Total SPP, F0+F1+F2+F3 = Total.
     Si algo no cuadra el script termina con error y no publica nada.
  4. Agrega en 3 vistas (clase de activo / emisor SBS / categorías del Excel) y escribe
     data/afp/hist_{FONDO}.json y data/afp/latest.json.

Uso:
  python scripts/fetch_sbs_afp.py                 # solo semanas nuevas
  python scripts/fetch_sbs_afp.py --desde 2021-01 # re-descarga desde ese mes (de publicación)
  python scripts/fetch_sbs_afp.py --rebuild       # re-agrega desde data/afp/raw sin descargar
  python scripts/fetch_sbs_afp.py --xls archivo.XLS
"""
import argparse, datetime as dt, glob, io, json, os, re, sys, time, unicodedata

import openpyxl
import requests
import xlrd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "data", "afp")
RAW = os.path.join(OUT, "raw")
BASE = "https://intranet2.sbs.gob.pe/estadistica/financiera"
MESES = [("Enero", "en"), ("Febrero", "fe"), ("Marzo", "ma"), ("Abril", "ab"), ("Mayo", "my"),
         ("Junio", "jn"), ("Julio", "jl"), ("Agosto", "ag"), ("Setiembre", "se"),
         ("Octubre", "oc"), ("Noviembre", "no"), ("Diciembre", "di")]
INICIO = "2021-01"   # primer mes de publicación (corte 31-dic-2020, igual que el Excel)
FONDOS = {"TOTAL": "Total Fondo", "F0": "Fondo0", "F1": "Fondo1", "F2": "Fondo2", "F3": "Fondo3"}
AFPS = ["Habitat", "Integra", "Prima", "Profuturo", "Total SPP"]
TOL = 0.5  # miles de S/ (tolerancia de redondeo de la SBS)
UA = {"User-Agent": "Mozilla/5.0 (dashboard_soberanos; +github.com/cconstantinis/dashboard_soberanos)"}


# ----------------------------------------------------------------------------- etiquetas
def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    s = re.sub(r"\(\d\)|\b\d/", " ", s)          # notas: (5)  5/  2/ 4/
    s = re.sub(r"[^a-z0-9.\- ]", " ", s)
    return re.sub(r"\s+", " ", s).strip(" .-")


SECCION = re.compile(r"^(i{1,3})\. ")            # I. / II. / III.
SUBSEC = re.compile(r"^([1-5])\. ")             # 1. Gobierno ...

# ------------------------------------------------------------- vista 1: clase de activo
GRUPOS_CLASE = [
    ("soberanos", "Soberanos PEN"),
    ("liquidez", "Liquidez"),
    ("rf_local", "Renta fija corporativa local"),
    ("rf_ext", "Renta fija exterior directa"),
    ("rv_local", "Renta variable local"),
    ("fondos_ext", "Fondos del exterior"),
    ("alternativos", "Alternativos"),
    ("otros", "Otros / sin clasificar"),
    ("transito", "Operaciones en tránsito"),
]
# (sección, subsección, prefijo de etiqueta normalizada) -> grupo. Se usa el primer match.
REGLAS_CLASE = [
    ("i", "1", "bonos del gobierno central", "soberanos"),
    ("i", "1", "bonos brady", "soberanos"),
    ("i", "1", "certificados y depositos a plazo del bcrp", "liquidez"),
    ("i", "1", "letras del tesoro", "liquidez"),
    ("i", "2", "certificados y depositos a plazo", "liquidez"),
    ("i", "2", "acciones", "rv_local"),
    ("i", "2", "", "rf_local"),
    ("i", "3", "acciones", "rv_local"),
    ("i", "3", "certificados de suscripcion", "rv_local"),
    ("i", "3", "certificado provisionales", "rv_local"),
    ("i", "3", "otros instrumentos autorizados", "rv_local"),
    ("i", "3", "fondo de inversiones alternativos", "alternativos"),
    ("i", "3", "", "rf_local"),
    ("i", "4", "fondos mutuos del extranjero", "fondos_ext"),
    ("i", "4", "fondo mutuo local", "fondos_ext"),
    ("i", "4", "", "alternativos"),
    ("i", "5", "titulos con derecho", "alternativos"),
    ("i", "5", "certificados de suscripcion", "alternativos"),
    ("i", "5", "", "rf_local"),
    ("ii", "1", "", "rf_ext"),
    ("ii", "2", "certificados y depositos", "liquidez"),
    ("ii", "2", "fondos mutuos", "fondos_ext"),
    ("ii", "2", "acciones", "fondos_ext"),
    ("ii", "2", "", "rf_ext"),
    ("ii", "3", "acciones", "fondos_ext"),
    ("ii", "3", "", "rf_ext"),
    ("ii", "4", "fondos mutuos alternativos", "alternativos"),
    ("ii", "4", "fondos de inversion", "alternativos"),
    ("ii", "4", "fondos mutuos", "fondos_ext"),
    ("ii", "4", "", "fondos_ext"),
    ("ii", "5", "", "rf_ext"),
]

# --------------------------------------------------------- vista 2: emisor (formato SBS)
SUBNOMBRE = {"1": "Gobierno", "2": "Sistema Financiero", "3": "Empresas no Financieras",
             "4": "Administradoras de Fondos", "5": "Sociedades Titulizadoras"}
GRUPOS_EMISOR = ([("loc_" + k, "Local · " + v) for k, v in SUBNOMBRE.items()] +
                 [("ext_" + k, "Exterior · " + v) for k, v in SUBNOMBRE.items()] +
                 [("transito", "Operaciones en tránsito")])

# ------------------------------------------- vista 3: categorías del Excel (Resumen semanal)
GRUPOS_EXCEL = [
    ("x_bgc", "Bonos del Gobierno Central"),
    ("x_cds", "Certificados y Depósitos a Plazo"),
    ("x_corp", "Bonos Corporativos Locales"),
    ("x_rv", "Renta Variable Local"),
    ("x_admf", "Administradoras de Fondos (local)"),
    ("x_resto_loc", "Resto local"),
    ("x_ext_gob", "Exterior · Gobierno"),
    ("x_ext_sf", "Exterior · Sistema Financiero"),
    ("x_ext_enf", "Exterior · Empresas no Financieras"),
    ("x_ext_admf", "Exterior · Administradoras de Fondos"),
    ("x_resto_ext", "Resto exterior"),
    ("transito", "Operaciones en tránsito"),
]
EXCEL_CORP = {("2", "papeles comerciales"), ("2", "bonos de arrendamiento financiero"),
              ("2", "bonos subordinados"), ("2", "otros bonos sector financiero"),
              ("2", "bonos hipotecarios"), ("3", "papeles comerciales"),
              ("3", "bonos de empresas no financieras"), ("3", "bonos para nuevos proyectos"),
              ("3", "bonos estructurados"), ("5", "bonos de titulizacion")}
EXCEL_RV = {("2", "acciones y valores representativos sobre acciones"), ("2", "acciones preferentes"),
            ("3", "acciones de capital por privatizacion"),
            ("3", "acciones y valores representativos sobre acciones")}


def grupo_clase(sec, sub, lab):
    if sec == "iii":
        return "transito"
    if lab.startswith("sin desglose"):
        return "otros"
    for s, u, pref, g in REGLAS_CLASE:
        if s == sec and u == sub and lab.startswith(pref):
            return g
    return "otros"


def grupo_excel(sec, sub, lab):
    if sec == "iii":
        return "transito"
    if lab.startswith("sin desglose"):
        return "x_resto_loc" if sec == "i" else "x_resto_ext"
    if sec == "ii":
        return {"1": "x_ext_gob", "2": "x_ext_sf", "3": "x_ext_enf", "4": "x_ext_admf"}.get(sub, "x_resto_ext")
    if sub == "1" and lab == "bonos del gobierno central":
        return "x_bgc"
    if sub == "2" and lab == "certificados y depositos a plazo":
        return "x_cds"
    if (sub, lab) in EXCEL_CORP:
        return "x_corp"
    if (sub, lab) in EXCEL_RV:
        return "x_rv"
    if sub == "4":
        return "x_admf"
    return "x_resto_loc"


def grupo_emisor(sec, sub, lab):
    if sec == "iii":
        return "transito"
    return ("loc_" if sec == "i" else "ext_") + sub


VISTAS = {"clase": (GRUPOS_CLASE, grupo_clase),
          "emisor": (GRUPOS_EMISOR, grupo_emisor),
          "excel": (GRUPOS_EXCEL, grupo_excel)}


# ----------------------------------------------------------------------------- parseo
class DatosMalos(Exception):
    pass


class HojaXlsx:
    """Adaptador mínimo de openpyxl con la interfaz de xlrd usada aquí (la SBS a veces
    publica un .xlsx con extensión .XLS)."""
    def __init__(self, ws):
        self.name = ws.title
        self._rows = [list(r) for r in ws.iter_rows(values_only=True)]
        self.nrows = len(self._rows)
        self.ncols = max((len(r) for r in self._rows), default=0)

    def cell_value(self, r, c):
        fila = self._rows[r]
        v = fila[c] if c < len(fila) else None
        if v is None:
            return ""
        if isinstance(v, dt.datetime):
            return (v - dt.datetime(1899, 12, 30)).days * 1.0
        return v


class LibroXlsx:
    datemode = 0

    def __init__(self, contenido):
        wb = openpyxl.load_workbook(io.BytesIO(contenido), read_only=True, data_only=True)
        self._h = {ws.title: HojaXlsx(ws) for ws in wb.worksheets}

    def sheets(self):
        return list(self._h.values())

    def sheet_by_name(self, n):
        return self._h[n]


def abrir_libro(contenido):
    if contenido[:2] == b"PK":
        return LibroXlsx(contenido)
    return xlrd.open_workbook(file_contents=contenido)


def parse_hoja(book, nombre):
    s = book.sheet_by_name(nombre)
    # fecha de corte (A3)
    v = s.cell_value(2, 0)
    if isinstance(v, float):
        fecha = xlrd.xldate_as_datetime(v, book.datemode).date()
    else:
        m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", str(v))
        if not m:
            raise DatosMalos(f"{nombre}: no se pudo leer la fecha '{v}'")
        fecha = dt.date(int(m[3]), int(m[2]), int(m[1]))
    # columnas de cada AFP (fila de encabezados con 'Habitat')
    cols, fila_hdr = {}, None
    for r in range(3, 12):
        fila = [str(s.cell_value(r, c)).strip() for c in range(s.ncols)]
        if any(norm(x) == "integra" for x in fila):
            fila_hdr = r
            for c, x in enumerate(fila):
                for a in AFPS:
                    if norm(x) == norm(a) or (a == "Total SPP" and norm(x) == "total"):
                        cols[a] = c
            break
    if fila_hdr is None or set(cols) != set(AFPS):
        raise DatosMalos(f"{nombre}: encabezados de AFP no encontrados ({cols})")

    def num(r, c):
        x = s.cell_value(r, c)
        if isinstance(x, (int, float)):
            return float(x)
        x = str(x).strip().replace(",", "")
        return float(x) if re.fullmatch(r"-?\d+(\.\d+)?", x) else 0.0

    sec = sub = None
    lineas, subtot, sectot, total = [], {}, {}, None
    for r in range(fila_hdr + 1, s.nrows):
        etiqueta = str(s.cell_value(r, 0)).strip()
        if not etiqueta:
            continue
        lab = norm(etiqueta)
        vals = {a: num(r, cols[a]) for a in AFPS}
        m = SECCION.match(lab)
        if m:
            sec, sub = m.group(1), None
            sectot[sec] = vals
            if sec == "iii":
                lineas.append({"sec": "iii", "sub": "", "label": etiqueta, "key": lab, "v": vals})
            continue
        if lab == "total":
            total = vals
            break
        m = SUBSEC.match(lab)
        if m and sec in ("i", "ii"):
            sub = m.group(1)
            subtot[(sec, sub)] = vals
            continue
        if sec in ("i", "ii") and sub:
            lineas.append({"sec": sec, "sub": sub, "label": re.sub(r"\s+", " ", etiqueta),
                           "key": lab, "v": vals})
    if total is None:
        raise DatosMalos(f"{nombre}: no se encontró la fila TOTAL")

    # subsecciones sin detalle -> la subsección es la línea
    for (sc, sb), vals in subtot.items():
        if not any(l["sec"] == sc and l["sub"] == sb for l in lineas) and any(abs(x) > TOL for x in vals.values()):
            lineas.append({"sec": sc, "sub": sb, "label": SUBNOMBRE[sb], "key": norm(SUBNOMBRE[sb]), "v": vals})

    # ---- validaciones
    # Descuadres internos de la SBS (líneas != subtotal): si son chicos se registran como
    # "Sin desglose (diferencia SBS)" -> grupo 'otros', con aviso. Si son grandes, error.
    errores, avisos = [], []
    for (sc, sb), vals in subtot.items():
        dif = {a: vals[a] - sum(l["v"][a] for l in lineas if l["sec"] == sc and l["sub"] == sb) for a in AFPS}
        if any(abs(dif[a]) > max(TOL * 4, abs(vals[a]) * 1e-6) for a in AFPS):
            lineas.append({"sec": sc, "sub": sb, "label": "Sin desglose (diferencia SBS)",
                           "key": "sin desglose diferencia sbs", "v": dif})
            avisos.append(f"{nombre} {fecha} {sc}.{sb}: diferencia SBS {dif['Total SPP']:,.0f} miles")
    for a in AFPS:
        for sc in ("i", "ii"):
            if sc in sectot:
                suma = sum(v[a] for (x, _), v in subtot.items() if x == sc)
                if abs(suma - sectot[sc][a]) > TOL * 6:
                    errores.append(f"{a} sección {sc}: {suma:,.0f} vs {sectot[sc][a]:,.0f}")
        suma = sum(sectot.get(sc, {}).get(a, 0.0) for sc in ("i", "ii", "iii"))
        if abs(suma - total[a]) > TOL * 6:
            errores.append(f"{a}: I+II+III {suma:,.0f} vs TOTAL {total[a]:,.0f}")
    afps_sum = sum(total[a] for a in AFPS[:-1])
    if abs(afps_sum - total["Total SPP"]) > TOL * 8:
        errores.append(f"Suma AFPs {afps_sum:,.0f} vs Total SPP {total['Total SPP']:,.0f}")
    for l in lineas:
        if l["key"] == "sin desglose diferencia sbs":
            for a in AFPS:
                if abs(l["v"][a]) > 0.02 * max(abs(total[a]), 1.0) and abs(l["v"][a]) > 1000:
                    errores.append(f"{a} {l['sec']}.{l['sub']}: descuadre SBS demasiado grande ({l['v'][a]:,.0f})")
    if errores:
        raise DatosMalos(f"{nombre} ({fecha}): " + "; ".join(errores[:6]))
    for x in avisos:
        print("  AVISO", x)

    return fecha, {"total": total, "lineas": lineas, "avisos": avisos}


def parse_xls(contenido, url=""):
    book = abrir_libro(contenido)
    nombres = {norm(s.name): s.name for s in book.sheets()}
    out, fecha = {"url": url, "fondos": {}}, None
    for fk, hoja in FONDOS.items():
        real = nombres.get(norm(hoja))
        if not real:
            if fk == "TOTAL":
                raise DatosMalos(f"no existe la hoja '{hoja}'")
            continue
        f, d = parse_hoja(book, real)
        if fecha and f != fecha:
            raise DatosMalos(f"fechas distintas entre hojas: {fecha} vs {f} ({real})")
        fecha = f
        out["fondos"][fk] = d
    # F0..F3 = Total
    fs = [k for k in ("F0", "F1", "F2", "F3") if k in out["fondos"]]
    if len(fs) == 4:
        for a in AFPS:
            suma = sum(out["fondos"][k]["total"][a] for k in fs)
            tot = out["fondos"]["TOTAL"]["total"][a]
            if abs(suma - tot) > max(50.0, tot * 1e-5):
                raise DatosMalos(f"{fecha} {a}: F0+F1+F2+F3 {suma:,.0f} vs Total {tot:,.0f}")
    out["fecha"] = fecha.isoformat()
    return out


# ------------------------------------------------------------------------- descarga
def urls_mes(anio, mes):
    nom, ab = MESES[mes - 1]
    for nn in range(1, 6):
        for ext in ("XLS", "xls"):
            yield nn, f"{BASE}/{anio}/{nom}/FP-1336-{nn:02d}-{ab}{anio}.{ext}"


def descargar(url, ses):
    for intento in range(3):
        try:
            r = ses.get(url, timeout=60, headers=UA)
            if r.status_code == 404:
                return None
            r.raise_for_status()
            if r.content[:4] != b"\xd0\xcf\x11\xe0" and r.content[:2] != b"PK":
                if b"Incapsula" in r.content[:2000]:
                    raise RuntimeError("bloqueado por el anti-bot de la SBS")
                return None
            return r.content
        except (requests.RequestException, RuntimeError) as e:
            if intento == 2:
                raise
            print(f"  reintento {url}: {e}")
            time.sleep(5 * (intento + 1))


def meses(desde, hasta):
    y, m = desde
    while (y, m) <= hasta:
        yield y, m
        m += 1
        if m == 13:
            y, m = y + 1, 1


def guardar_raw(d):
    os.makedirs(RAW, exist_ok=True)
    comp = {"fecha": d["fecha"], "url": d["url"], "fondos": {}}
    for fk, f in d["fondos"].items():
        comp["fondos"][fk] = {
            "total": [round(f["total"][a], 1) for a in AFPS],
            "lineas": [[l["sec"], l["sub"], l["label"], [round(l["v"][a], 1) for a in AFPS]]
                       for l in f["lineas"]],
        }
    with open(os.path.join(RAW, d["fecha"] + ".json"), "w", encoding="utf-8") as fh:
        json.dump(comp, fh, ensure_ascii=False, separators=(",", ":"))


def cargar_raw():
    res = []
    for p in sorted(glob.glob(os.path.join(RAW, "*.json"))):
        with open(p, encoding="utf-8") as fh:
            res.append(json.load(fh))
    return res


def ultimo_mes_raw():
    """Mes de publicación del último corte guardado (por la URL)."""
    ult = None
    for p in sorted(glob.glob(os.path.join(RAW, "*.json"))):
        with open(p, encoding="utf-8") as fh:
            u = json.load(fh).get("url", "")
        m = re.search(r"/financiera/(\d{4})/", u)
        ab = re.search(r"-\d{2}-([a-z]{2})\d{4}\.", u, re.I)
        if m and ab:
            mes = [x[1] for x in MESES].index(ab.group(1).lower()) + 1
            ult = max(ult or (0, 0), (int(m.group(1)), mes))
    return ult


def actualizar(desde=None, hasta=None):
    hoy = dt.date.today()
    if desde:
        ini = tuple(int(x) for x in desde.split("-"))
    else:
        ini = ultimo_mes_raw() or tuple(int(x) for x in INICIO.split("-"))
    existentes = {os.path.basename(p)[:10] for p in glob.glob(os.path.join(RAW, "*.json"))}
    ses = requests.Session()
    nuevos = 0
    fin = tuple(int(x) for x in hasta.split("-")) if hasta else (hoy.year, hoy.month)
    for y, m in meses(ini, fin):
        vistos = set()
        for nn, url in urls_mes(y, m):
            if nn in vistos:
                continue
            cont = descargar(url, ses)
            if cont is None:
                continue
            vistos.add(nn)
            d = parse_xls(cont, url)
            if d["fecha"] in existentes and not desde:
                continue
            guardar_raw(d)
            existentes.add(d["fecha"])
            nuevos += 1
            print(f"  {d['fecha']}  <- {url.rsplit('/', 1)[-1]}")
            time.sleep(0.3)
    return nuevos


# ------------------------------------------------------------------------- agregación
def agregar(raws):
    raws = sorted(raws, key=lambda r: r["fecha"])
    fechas = [r["fecha"] for r in raws]
    hist = {}
    for fk in FONDOS:
        h = {"fondo": fk, "fechas": [], "afps": AFPS, "total": {a: [] for a in AFPS},
             "vistas": {v: {a: {g: [] for g, _ in grupos} for a in AFPS} for v, (grupos, _) in VISTAS.items()}}
        for r in raws:
            f = r["fondos"].get(fk)
            if not f:
                continue
            h["fechas"].append(r["fecha"])
            for i, a in enumerate(AFPS):
                h["total"][a].append(round(f["total"][i] / 1000, 2))
            for v, (grupos, fn) in VISTAS.items():
                acc = {a: {g: 0.0 for g, _ in grupos} for a in AFPS}
                for sec, sub, label, vals in f["lineas"]:
                    g = fn(sec, sub, norm(label))
                    for i, a in enumerate(AFPS):
                        acc[a][g] += vals[i]
                for a in AFPS:
                    for g, _ in grupos:
                        h["vistas"][v][a][g].append(round(acc[a][g] / 1000, 2))
        hist[fk] = h
    return fechas, hist


def detalle(raw, fk):
    """Líneas de detalle (S/ MM) de un corte, con sus grupos en cada vista."""
    f = raw["fondos"].get(fk)
    if not f:
        return None
    out = []
    for sec, sub, label, vals in f["lineas"]:
        lab = norm(label)
        out.append({"sec": sec, "sub": sub, "label": label,
                    "g": {v: fn(sec, sub, lab) for v, (_, fn) in VISTAS.items()},
                    "v": [round(x / 1000, 2) for x in vals]})
    return {"fecha": raw["fecha"], "total": [round(x / 1000, 2) for x in f["total"]], "lineas": out}



def cruce_mef(hist_total):
    """Tenencia de soberanos de las AFPs según el MEF (mensual, nominal y a precio sucio)
    vs. 'Bonos del Gobierno Central' de la SBS en el corte semanal más cercano al cierre."""
    import calendar
    MES = {m: i + 1 for i, m in enumerate(["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                                             "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])}
    fechas = [dt.date.fromisoformat(f) for f in hist_total["fechas"]]
    sob = hist_total["vistas"]["clase"]["Total SPP"]["soberanos"]
    out = []
    for p in glob.glob(os.path.join(ROOT, "data", "data_*.json")):
        try:
            with open(p, encoding="utf-8") as fh:
                d = json.load(fh)
            mm, yy = d["meta"]["periodo"].split("-")
            fin = dt.date(int(yy), MES[mm], calendar.monthrange(int(yy), MES[mm])[1])
        except Exception:
            continue
        an = d.get("analitica_bonos", {})
        nom = mv = 0.0
        for b in d.get("outstanding_por_bono", []):
            x = b.get("Pension Funds", 0) or 0
            nom += x
            a = an.get(b.get("tenor"), {})
            mv += x * ((a.get("precio") or 100) + (a.get("interes_corrido") or 0)) / 100
        cands = [(abs((f - fin).days), i) for i, f in enumerate(fechas)]
        dist, i = min(cands) if cands else (99, None)
        out.append({"mes": fin.strftime("%Y-%m"), "mef_nominal": round(nom), "mef_mercado": round(mv),
                    "sbs_fecha": fechas[i].isoformat() if i is not None and dist <= 5 else None,
                    "sbs_bgc": round(sob[i]) if i is not None and dist <= 5 else None})
    return sorted(out, key=lambda r: r["mes"])


def publicar(raws):
    if not raws:
        raise SystemExit("No hay datos en data/afp/raw")
    raws = sorted(raws, key=lambda r: r["fecha"])
    fechas, hist = agregar(raws)
    otros = [(fk, a, fechas[i]) for fk, h in hist.items() for a in AFPS
             for i, x in enumerate(h["vistas"]["clase"][a]["otros"]) if abs(x) > 0.5]
    for fk, h in hist.items():
        with open(os.path.join(OUT, f"hist_{fk}.json"), "w", encoding="utf-8") as fh:
            json.dump(h, fh, ensure_ascii=False, separators=(",", ":"))

    ult = raws[-1]
    fu = dt.date.fromisoformat(ult["fecha"])

    def cerca(objetivo):
        cands = [r for r in raws if r["fecha"] <= objetivo.isoformat()]
        return cands[-1] if cands else raws[0]

    comps = {
        "ultimo": ult,
        "semana": raws[-2] if len(raws) > 1 else ult,
        "4sem": cerca(fu - dt.timedelta(days=28)),
        "ytd": cerca(dt.date(fu.year - 1, 12, 31)),
        "12m": cerca(fu - dt.timedelta(days=364)),
    }
    latest = {
        "meta": {
            "fuente": "SBS — Boletín Estadístico Semanal del SPP (Cuadros 8 a 12)",
            "url_indice": "https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaBoletinSPPSemanal.aspx",
            "url_xls": ult["url"], "fecha": ult["fecha"], "n_cortes": len(raws),
            "desde": raws[0]["fecha"], "unidad": "S/ millones (valor de mercado)",
            "generado": dt.datetime.utcnow().strftime("%Y-%m-%dT%H:%MZ"),
            "avisos": [f"Líneas sin clasificar en {fk}/{a} ({f})" for fk, a, f in otros[-5:]],
        },
        "afps": AFPS,
        "fondos": {fk: {"nombre": {"TOTAL": "Total", "F0": "Fondo 0", "F1": "Fondo 1",
                                     "F2": "Fondo 2", "F3": "Fondo 3"}[fk]} for fk in FONDOS},
        "vistas": {v: [{"id": g, "nombre": n} for g, n in grupos] for v, (grupos, _) in VISTAS.items()},
        "comparativos": {k: r["fecha"] for k, r in comps.items()},
        "detalle": {fk: {k: detalle(r, fk) for k, r in comps.items()} for fk in FONDOS},
        "cruce_mef": cruce_mef(hist["TOTAL"]),
    }
    with open(os.path.join(OUT, "latest.json"), "w", encoding="utf-8") as fh:
        json.dump(latest, fh, ensure_ascii=False, separators=(",", ":"))
    print(f"Publicado: {len(raws)} cortes, {raws[0]['fecha']} -> {ult['fecha']}")
    if otros:
        print("AVISO: líneas sin clasificar:", otros[-5:])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--desde", help="AAAA-MM (mes de publicación) para re-descargar")
    ap.add_argument("--hasta", help="AAAA-MM (mes de publicación) final; por defecto el actual")
    ap.add_argument("--solo-descarga", action="store_true", help="no re-agregar (backfill por tramos)")
    ap.add_argument("--rebuild", action="store_true", help="solo re-agregar desde data/afp/raw")
    ap.add_argument("--xls", help="parsear un XLS local")
    a = ap.parse_args()
    os.makedirs(RAW, exist_ok=True)
    try:
        if a.xls:
            with open(a.xls, "rb") as fh:
                guardar_raw(parse_xls(fh.read(), os.path.basename(a.xls)))
        elif not a.rebuild:
            n = actualizar(a.desde, a.hasta)
            if a.solo_descarga:
                print(f"Cortes descargados: {n}")
                return
            print(f"Cortes nuevos: {n}")
            if n == 0 and os.path.exists(os.path.join(OUT, "latest.json")):
                return
        publicar(cargar_raw())
    except DatosMalos as e:
        print("ERROR de validación — no se publica:", e)
        sys.exit(1)


if __name__ == "__main__":
    main()
