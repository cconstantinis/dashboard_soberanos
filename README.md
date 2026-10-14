# Soberanos Bondholders Dashboard 🇵🇪

Dashboard interactivo de tenencias de Bonos Soberanos del Perú, con datos del **Ministerio de Economía y Finanzas (MEF)**.

## Live Dashboard

👉 **[Ver dashboard](https://cconstantinis.github.io/dashboard_soberanos/)**

## Contenido

- **Holdings por tipo de inversor** — KPIs con % de tenencia (Offshores, AFPs, Bancos, Seguros, Fondos Públicos)
- **Evolución histórica de ownership** — line chart desde Mar-23
- **Estructura por tenor** — stacked bar chart por bono (26s → 55s)
- **Outstanding por bono e institución** — tabla completa en MM PEN
- **Cambios MoM** — tabla de variaciones mensuales por inversor y tenor
- **Movimientos DV01 MoM** — horizontal bar charts por tipo de inversor
- **Flujos históricos (pág. 2 del reporte)** — Offshores, AFPs y Bancos: cambio MoM nominal + DV01 y evolución acumulada desde Abr-23

## Datos

Los datos provienen del [Reporte de Tenencias de Bonos Soberanos](https://www.mef.gob.pe/contenidos/deuda_publ/mercado/reportes_tenencia.php) del MEF, publicado mensualmente.

El archivo `data/latest.json` es el que usa el dashboard. Se actualiza automáticamente via GitHub Actions.

## Actualización automática

`update-data.yml` corre lunes y jueves. Si el MEF publicó un reporte nuevo:
1. Descarga el PDF de tenencias más reciente (ordenado por fecha del nombre de archivo)
2. Lo parsea por **coordenadas** (el PDF son tortas de Excel: una por bono, dos por fila)
3. **Valida** que cada torta sume ~100% y que las unidades cuadren con el total MN nominal.
   Si algo no cuadra, el job falla y **no** publica datos malos.
4. Recalcula outstanding (MM PEN), MoM vs el mes anterior, DV01 y la serie de evolución
5. Hace commit de `data/` e `index.html`

Desde la pestaña **Actions → Run workflow** se puede re-procesar con `backfill = N` o con
`desde = 2023-03` (reconstruye toda la historia mensual que usan los gráficos de flujos).

## Actualización manual

```bash
pip install -r requirements.txt
python scripts/fetch_mef_data.py                 # último reporte
python scripts/fetch_mef_data.py --backfill 7    # re-procesa los 7 últimos meses
python scripts/fetch_mef_data.py --pdf tenencia_bono_310726.pdf   # PDF local
```

Notas de datos:
- Las unidades por bono salen del mismo PDF de tenencias (ya no se usa el PDF de stock).
- `Others` = Otros + Fondos privados + Personas naturales. SOB29 = 12FEB2029 + 12FEB2029E.
- DV01 (K PEN por pb) = ΔMM nominal × precio sucio/100 × duración modificada × 0.1.
  Precio, intereses corridos, yield y duración modificada salen del **Reporte Diario del MEF**
  del último día hábil publicado del mes (`/contenidos/english/report/AAAA/Daily_MM_DD_AA.pdf`).
  Si no se encuentra, se usa una duración estimada con yield = cupón (o `data/yields.json`).
- `data/historico_evolucion.json` guarda la serie manual previa a 2026 (Mar-23 → Nov-25).

## Tab AFPs (SBS semanal)

Segundo tab del dashboard (`#afps`): cartera administrada de las AFPs por instrumento, con datos del
[Boletín Estadístico Semanal del SPP](https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaBoletinSPPSemanal.aspx)
de la SBS (Cuadro 8 — Total, y Cuadros 9-12 — Fondos 0 a 3). Reemplaza el Excel *Resumen AFPs*.

- **Fuente:** el índice de la SBS tiene anti-bot, pero los XLS se descargan directo con una URL predecible:
  `intranet2.sbs.gob.pe/estadistica/financiera/{AAAA}/{Mes}/FP-1336-{NN}-{mm}{AAAA}.XLS` (a veces es un .xlsx con extensión .XLS).
- **Parser:** `scripts/fetch_sbs_afp.py` lee cada línea por su etiqueta dentro de su sección (no por número de fila),
  así resiste filas nuevas y cambios de notas ("(5)" → "5/").
- **Validación** (si falla, no publica): líneas = subtotales, I+II+III = TOTAL, AFPs = Total SPP, F0+F1+F2+F3 = Total.
  Descuadres chicos de la propia SBS entre líneas y subtotal quedan como "Sin desglose (diferencia SBS)" con aviso.
- **Agrupaciones:** *Clase de activo* (Soberanos PEN, Liquidez, RF corporativa local, RF exterior directa, RV local,
  Fondos del exterior, Alternativos), *Emisor (SBS)* y *Categorías Excel* (las 11 del Excel original;
  validadas contra sus 89 cortes: diferencia máx. 0.0002 pp).
- **Datos:** `data/afp/raw/AAAA-MM-DD.json` (un corte por semana, desde 31-dic-2020), `data/afp/hist_{TOTAL,F0..F3}.json`
  (series por grupo) y `data/afp/latest.json` (detalle de líneas, comparativos y cruce MEF vs SBS).
- **Workflow:** `update-afp.yml` corre sábados y martes. *Run workflow* con `desde = AAAA-MM` re-descarga desde ese mes.

```bash
python scripts/fetch_sbs_afp.py                       # semanas nuevas
python scripts/fetch_sbs_afp.py --desde 2026-01       # re-descarga desde ese mes de publicación
python scripts/fetch_sbs_afp.py --rebuild             # re-agrega desde data/afp/raw
```

Para verlo localmente hay que servirlo (`python -m http.server`), porque el tab AFPs lee los JSON con `fetch`.

## Publicar en GitHub Pages

1. Sube este repositorio a GitHub
2. Ve a **Settings → Pages**
3. Source: **Deploy from a branch** → `main` → `/ (root)`
4. El dashboard queda en `https://cconstantinis.github.io/dashboard_soberanos/`

## Estructura

```
soberanos-dashboard/
├── index.html                  # Dashboard (standalone HTML)
├── data/
│   ├── latest.json             # Datos más recientes (lo lee el dashboard)
│   └── data_marzo_2026.json    # Histórico por período
├── scripts/
│   └── fetch_mef_data.py       # Script de actualización de datos
└── .github/
    └── workflows/
        └── update-data.yml     # GitHub Action mensual
```

## Fuente

Ministerio de Economía y Finanzas del Perú — Dirección General del Tesoro Público  
https://www.mef.gob.pe/contenidos/deuda_publ/mercado/reportes_tenencia.php
