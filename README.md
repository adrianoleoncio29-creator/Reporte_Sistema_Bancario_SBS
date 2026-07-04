# SBS Scraper — Descarga de Estadísticas Banca Múltiple

Automatiza la descarga de archivos Excel del portal de estadísticas de la SBS:
https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaBoletinEstadistico.aspx?p=1#

## Instalación

```bash
pip install -r requirements.txt
playwright install chromium
```

## Uso

```bash
python sbs_downloader.py
```

Los archivos se guardan en `descargas_sbs/` replicando la jerarquía del portal:

```
descargas_sbs/
└── Información de la Banca Múltiple/
      └── Créditos directos por sector económico/
            ├── Enero 2024.xlsx
            ├── Febrero 2024.xlsx
            └── ...
```

## Configuración

Edita las constantes al inicio de `sbs_downloader.py`:

| Constante              | Descripción                                              |
|------------------------|----------------------------------------------------------|
| `BASE_DOWNLOAD_DIR`    | Carpeta raíz de descarga (por defecto `descargas_sbs`)   |
| `ROOT_NODE_TEXT`       | Texto exacto del nodo raíz a expandir en el menú         |
| `SUBCARPETAS_OBJETIVO` | Lista de subcarpetas a procesar                          |
| `TIMEOUT_ELEMENTO`     | Tiempo máximo de espera por elemento (ms)                |
| `TIMEOUT_DESCARGA`     | Tiempo máximo de espera por descarga (ms)                |
| `headless`             | `False` = muestra el navegador, `True` = modo silencioso |
