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

Desde la pestaña **Actions → Run workflow** se puede re-procesar con `backfill = N`.

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
- DV01 (K PEN por pb) = ΔMM PEN × duración modificada × 0.1. Por defecto yield = cupón;
  se puede poner yields reales en `data/yields.json`, ej. `{"SOB35": 6.45, "SOB40": 6.90}`.
- `data/historico_evolucion.json` guarda la serie manual previa a 2026 (Mar-23 → Nov-25).

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
