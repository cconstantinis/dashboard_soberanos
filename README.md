# Soberanos Bondholders Dashboard 🇵🇪

Dashboard interactivo de tenencias de Bonos Soberanos del Perú, con datos del **Ministerio de Economía y Finanzas (MEF)**.

## Live Dashboard

👉 **[Ver dashboard](https://TU-USUARIO.github.io/soberanos-dashboard/)**

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

Un GitHub Action (`update-data.yml`) corre el día 5 de cada mes:
1. Descarga el PDF más reciente del MEF
2. Parsea los datos clave
3. Actualiza `data/latest.json`
4. Hace commit automático

También puedes correrlo manualmente desde la pestaña **Actions** de GitHub.

## Actualización manual

```bash
pip install pdfplumber requests beautifulsoup4
python scripts/fetch_mef_data.py
```

## Publicar en GitHub Pages

1. Sube este repositorio a GitHub
2. Ve a **Settings → Pages**
3. Source: **Deploy from a branch** → `main` → `/ (root)`
4. El dashboard queda en `https://TU-USUARIO.github.io/soberanos-dashboard/`

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
