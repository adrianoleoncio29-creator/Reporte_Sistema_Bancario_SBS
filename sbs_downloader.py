# -*- coding: utf-8 -*-
"""
sbs_downloader.py
=================
Script de web scraping con Playwright para descargar archivos Excel del portal
de estadísticas de la SBS (Superintendencia de Banca, Seguros y AFP del Perú).

URL objetivo:
    https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaBoletinEstadistico.aspx?p=1#

Arquitectura real del portal (verificada por inspección del DOM):
    - El menú lateral usa divs con onclick="cambiarDisplay()" para expandir/colapsar.
    - Cada subcarpeta tiene un enlace que carga una página de resultados con URL:
        EstadisticaSistemaFinancieroResultados.aspx?c=<CODIGO>
    - En esa página, cada periodo mensual es un enlace <a href="...XLS"> directo
      al archivo en https://intranet2.sbs.gob.pe/estadistica/financiera/<año>/<mes>/
    - El año está codificado en el href, no en el texto del enlace.

Flujo:
    1. Abrir el portal y expandir "Información de la Banca Múltiple".
    2. Para cada subcarpeta configurada, navegar a su página de resultados.
    3. Extraer todos los href de enlaces .XLS/.XLSX de la página.
    4. Descargar cada archivo con Playwright (interceptando la descarga).
    5. Guardar replicando la jerarquía: <raíz>/<subcarpeta>/<año>/<mes>.xls

Requisitos:
    pip install playwright
    python -m playwright install chromium
"""

import asyncio
import re
import sys
import io
from pathlib import Path
from urllib.parse import urlparse

# Fuerza UTF-8 en stdout/stderr para que los textos con tildes
# se muestren correctamente en la terminal de Windows.
if sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

from playwright.async_api import (
    async_playwright,
    Page,
    Download,
    TimeoutError as PWTimeoutError,
)

# ---------------------------------------------------------------------------
# CONFIGURACIÓN GLOBAL
# ---------------------------------------------------------------------------

PORTAL_URL = (
    "https://www.sbs.gob.pe/app/stats_net/stats/"
    "EstadisticaBoletinEstadistico.aspx?p=1#"
)

# Directorio raíz de descarga (se crea automáticamente)
BASE_DOWNLOAD_DIR = Path("descargas_sbs")

# Texto exacto del nodo raíz en el menú lateral
ROOT_NODE_TEXT = "Información de la Banca Múltiple"

# Subcarpetas de interés.
# Cada entrada es un dict con:
#   'texto' : texto exacto del enlace en el menú (para hacer clic)
#   'codigo': código de la URL (c=<CODIGO>) — confirmado por inspección del DOM
#
# Todos los códigos fueron verificados en el portal el 25/05/2026.
# Comenta las líneas que NO quieras descargar.
SUBCARPETAS_OBJETIVO = [

    # ── Información de la Banca Múltiple ─────────────────────────────────
    {"texto": "Créditos Directos por Sector Económico",                                          "codigo": "B-2311"},
    {"texto": "Créditos Directos y Depósitos por Zona Geográfica",                               "codigo": "B-2314"},
    {"texto": "Depósitos según Escala de Montos",                                                "codigo": "B-2321"},
    {"texto": "Contratos de Arrendamiento Financiero por Sector Económico y Tipo de Bien",       "codigo": "B-2323"},
    {"texto": "Número de Personal",                                                              "codigo": "B-2305"},

    # ── Estados Financieros por Empresa Bancaria ──────────────────────────
    {"texto": "Balance General y Estado de Ganancias y Pérdidas",                                "codigo": "B-2201"},

    # ── Indicadores de las Empresas Bancarias ─────────────────────────────
    {"texto": "Indicadores Financieros",                                                         "codigo": "B-2401"},
    {"texto": "Requerimiento de Patrimonio Efectivo y Ratio de Capital Global",                  "codigo": "B-2402"},

    # ── Riesgo Crediticio ─────────────────────────────────────────────────
    {"texto": "Activos y Contingentes Ponderados por Riesgo de Crédito",                         "codigo": "B-2306"},
    {"texto": "Créditos Directos según Situación",                                               "codigo": "B-2315"},
    {"texto": "Créditos Directos según Tipo de Crédito y Situación",                             "codigo": "B-2334"},
    {"texto": "Estructura de Créditos Directos e Indirectos según Categoría de Riesgo del Deudor","codigo": "B-2309"},
    {"texto": "Estructura de Créditos Directos e Indirectos por Tipo de Crédito y Categoría de Riesgo del Deudor", "codigo": "B-230802"},
    {"texto": "Créditos Directos por Tipo, Modalidad y Moneda",                                  "codigo": "B-2359"},
    {"texto": "Morosidad por tipo de crédito y modalidad",                                       "codigo": "B-2362"},
    {"texto": "Ratios de Morosidad según días de incumplimiento",                                "codigo": "B-220512"},
    {"texto": "Créditos por Tipo de Garantía",                                                   "codigo": "B-2366"},
    {"texto": "Créditos a Actividades Empresariales por Sector Económico",                       "codigo": "B-2336"},
    {"texto": "Flujo de Créditos Castigados por Tipo de Crédito",                                "codigo": "B-2369"},

    # ── Riesgo de Liquidez ────────────────────────────────────────────────
    {"texto": "Ratios de Liquidez",                                                              "codigo": "B-2340"},
    {"texto": "Movimiento de los Depósitos",                                                     "codigo": "B-2318"},
    {"texto": "Depósitos del Público por Tipo de Depósito y Plazo",                              "codigo": "B-220513"},
    {"texto": "Ratio de Cobertura de Liquidez",                                                  "codigo": "B-230809"},
    {"texto": "Ratio de Financiación Neta Estable",                                              "codigo": "B-234021"},

    # ── Riesgo de Mercado ─────────────────────────────────────────────────
    {"texto": "Requerimiento de Patrimonio Efectivo por Riesgo de Mercado",                      "codigo": "B-2337"},
    {"texto": "Posición Global en Moneda Extranjera",                                            "codigo": "B-2368"},
    {"texto": "Operaciones Forward en Moneda Extranjera",                                        "codigo": "B-2339"},
    {"texto": "Operaciones Swaps",                                                               "codigo": "B-2371"},
    {"texto": "Ganancias en Riesgo y Valor Patrimonial en Riesgo",                               "codigo": "B-2379"},

    # ── Riesgo Operacional ────────────────────────────────────────────────
    {"texto": "Requerimiento de Patrimonio Efectivo por Riesgo Operacional",                     "codigo": "B-2405"},

    # ── Estructura de las Principales Cuentas ─────────────────────────────
    {"texto": "Estructura del Activo",                                                           "codigo": "B-2341"},
    {"texto": "Estructura de los Créditos Directos por Modalidad",                               "codigo": "B-2343"},
    {"texto": "Estructura de los Créditos Directos por Tipo y Modalidad",                        "codigo": "B-2389"},
    {"texto": "Estructura de los Créditos Indirectos",                                           "codigo": "B-2322"},
    {"texto": "Estructura del Pasivo",                                                           "codigo": "B-2342"},
    {"texto": "Estructura de los Depósitos por Tipo",                                            "codigo": "B-2344"},
    {"texto": "Estructura de los Adeudos y Obligaciones Financieras",                            "codigo": "B-2310"},
    {"texto": "Estructura del Patrimonio Efectivo",                                              "codigo": "B-2370"},
    {"texto": "Estructura de los Ingresos Financieros",                                          "codigo": "B-2347"},
    {"texto": "Estructura de los Gastos Financieros",                                            "codigo": "B-2390"},
    {"texto": "Estructura de los Gastos de Administración",                                      "codigo": "B-2348"},
    {"texto": "Estructura de Fideicomisos y Comisiones de Confianza",                            "codigo": "B-2365"},
    {"texto": "Estructura de Inversiones",                                                       "codigo": "B-2381"},

    # ── Alcance y Participación de Mercado ────────────────────────────────
    {"texto": "Ranking de Créditos, Depósitos y Patrimonio",                                     "codigo": "B-2332"},
    {"texto": "Ranking de Créditos Directos por Tipo",                                           "codigo": "B-2333"},
    {"texto": "Ranking de Créditos Directos por Modalidad de Operación",                         "codigo": "B-2313"},
    {"texto": "Ranking de Depósitos por Tipo",                                                   "codigo": "B-2320"},
    {"texto": "Distribución de Oficinas por Zona Geográfica",                                    "codigo": "B-2303"},
    {"texto": "Cajeros corresponsales y automáticos",                                            "codigo": "B-2364"},
    {"texto": "Créditos Directos y Depósitos por Oficinas",                                      "codigo": "B-2358"},
    {"texto": "Estructura de los Créditos Directos por Departamento",                            "codigo": "B-2349"},
    {"texto": "Estructura de los Depósitos por Departamento",                                    "codigo": "B-2350"},
    {"texto": "Depósitos por Tipo y Persona",                                                    "codigo": "B-2372"},
    {"texto": "Número de Depositantes por Tipo de Depósito",                                     "codigo": "B-2373"},
    {"texto": "Número de Tarjetas de Débito",                                                    "codigo": "B-2391"},
    {"texto": "Número de Deudores según Tipo de Crédito",                                        "codigo": "B-230803"},
    {"texto": "Número de Tarjetas de Crédito por Tipo",                                          "codigo": "B-2363"},
    {"texto": "Nuevos Créditos Hipotecarios para Vivienda",                                      "codigo": "B-2367"},
    {"texto": "Nuevos Créditos a Principales Sectores Económicos",                               "codigo": "B-2392"},
    {"texto": "Operaciones con dinero electrónico",                                              "codigo": "B-230804"},
    {"texto": "Número de cuentas y monto de dinero electrónico",                                 "codigo": "B-230805"},
]

# URL base de resultados (se le agrega ?c=<CODIGO>)
BASE_RESULTADOS_URL = (
    "https://www.sbs.gob.pe/app/stats_net/stats/"
    "EstadisticaSistemaFinancieroResultados.aspx?c="
)

# Tiempos de espera en milisegundos
TIMEOUT_ELEMENTO = 25_000   # 25 s para que un elemento aparezca
TIMEOUT_DESCARGA  = 90_000  # 90 s por archivo

# Pausa entre descargas (segundos) para no saturar el servidor
PAUSA_ENTRE_DESCARGAS = 1.0

# Si True, omite archivos que ya existen en disco (reanuda descargas interrumpidas)
OMITIR_EXISTENTES = True

# Descarga desde este periodo en adelante (inclusive).
# Los archivos que ya existen en disco se saltean automáticamente gracias a
# OMITIR_EXISTENTES = True, por lo que solo se descargan los meses nuevos.
DESDE_PERIODO = (2024, 1)   # Enero 2024 en adelante


# ---------------------------------------------------------------------------
# UTILIDADES
# ---------------------------------------------------------------------------

def sanitizar_nombre(nombre: str) -> str:
    """Convierte texto en nombre de carpeta/archivo válido para Windows/Linux."""
    nombre = nombre.strip()
    nombre = re.sub(r'[\\/:*?"<>|]', "_", nombre)
    nombre = re.sub(r"\s+", " ", nombre)
    return nombre


# Mapeo de nombre de mes (en español, como aparece en las URLs del portal) → número
MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "setiembre": 9, "septiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}


def periodo_es_valido(año: str, mes: str) -> bool:
    """
    Retorna True si el periodo (año, mes) es igual o posterior a DESDE_PERIODO.
    Si DESDE_PERIODO es None, siempre retorna True (sin filtro).
    """
    if DESDE_PERIODO is None:
        return True
    try:
        año_int = int(año)
        mes_int = MESES.get(mes.lower(), 0)
        desde_año, desde_mes = DESDE_PERIODO
        return (año_int, mes_int) >= (desde_año, desde_mes)
    except (ValueError, TypeError):
        # Si no se puede parsear, incluye el archivo por precaución
        return True


def extraer_año_mes_de_url(url: str) -> tuple[str, str]:
    """
    Extrae el año y el mes desde la URL del archivo.
    Ejemplo:
        .../estadistica/financiera/2025/Enero/B-2311-en2025.XLS
        → ('2025', 'Enero')
    """
    partes = urlparse(url).path.split("/")
    # La estructura es: .../financiera/<año>/<mes>/<archivo>
    try:
        idx = partes.index("financiera")
        año = partes[idx + 1]   # ej. "2025"
        mes = partes[idx + 2]   # ej. "Enero"
        return año, mes
    except (ValueError, IndexError):
        # Fallback: extrae año con regex
        match_año = re.search(r"/(\d{4})/", url)
        match_mes = re.search(
            r"/(enero|febrero|marzo|abril|mayo|junio|julio|agosto"
            r"|setiembre|septiembre|octubre|noviembre|diciembre)/",
            url, re.IGNORECASE,
        )
        año = match_año.group(1) if match_año else "año_desconocido"
        mes = match_mes.group(1).capitalize() if match_mes else "mes_desconocido"
        return año, mes


# ---------------------------------------------------------------------------
# PASO 1 — EXPANDIR EL NODO RAÍZ
# ---------------------------------------------------------------------------

async def expandir_nodo_raiz(page: Page) -> None:
    """
    Hace clic en "Información de la Banca Múltiple" para expandir sus hijos.
    El menú usa onclick="cambiarDisplay()" que cambia display:none → block.
    """
    print(f"\n[>>] Expandiendo: '{ROOT_NODE_TEXT}'...")
    try:
        # Espera a que el nodo raíz esté presente (attached), no necesariamente visible
        await page.wait_for_selector(
            f"a:has-text('{ROOT_NODE_TEXT}')",
            state="attached",
            timeout=TIMEOUT_ELEMENTO,
        )
        await page.get_by_text(ROOT_NODE_TEXT, exact=True).first.click()
        # Pausa generosa para que el JS del portal procese el clic
        await asyncio.sleep(2.5)

        # Verifica que los hijos están en el DOM (attached es suficiente —
        # el portal los tiene en DOM aunque con display:none al inicio)
        count = await page.locator("a:has-text('Créditos Directos')").count()
        if count == 0:
            raise RuntimeError("No se encontraron hijos tras expandir el nodo raíz.")

        print(f"  [OK] Nodo raiz expandido. ({count} hijos detectados en DOM)")
    except PWTimeoutError:
        raise RuntimeError(
            f"No se pudo encontrar '{ROOT_NODE_TEXT}' en la página. "
            "Verifica que la pagina cargo correctamente."
        )


# ---------------------------------------------------------------------------
# PASO 2 — OBTENER TODOS LOS ENLACES XLS DE UNA SUBCARPETA
# ---------------------------------------------------------------------------

async def obtener_enlaces_xls(page: Page, codigo: str) -> list[dict]:
    """
    Navega a la página de resultados de la subcarpeta y extrae todos los
    enlaces que apuntan a archivos .XLS/.XLSX.

    Retorna lista de dicts:
        url  : URL completa del archivo
        texto: texto del enlace (ej. "Enero")
        año  : año extraído de la URL (ej. "2025")
        mes  : mes extraído de la URL (ej. "Enero")
    """
    url_resultados = BASE_RESULTADOS_URL + codigo
    print(f"  [WEB] Cargando resultados: {url_resultados}")

    await page.goto(url_resultados, wait_until="networkidle", timeout=60_000)
    await asyncio.sleep(1.5)

    # Extrae todos los <a> cuyo href termina en .XLS o .XLSX
    enlaces_raw = await page.evaluate("""
        () => {
            const links = Array.from(document.querySelectorAll('a[href]'));
            return links
                .filter(a => /\\.xls[x]?$/i.test(a.getAttribute('href')))
                .map(a => ({
                    url: a.href,
                    texto: a.innerText.trim()
                }));
        }
    """)

    # Enriquece cada enlace con año y mes extraídos de la URL
    resultado = []
    for item in enlaces_raw:
        año, mes = extraer_año_mes_de_url(item["url"])
        resultado.append({
            "url": item["url"],
            "texto": item["texto"],
            "año": año,
            "mes": mes,
        })

    print(f"  [INFO] Archivos encontrados: {len(resultado)}")
    return resultado


# ---------------------------------------------------------------------------
# PASO 3 — DESCARGAR UN ARCHIVO XLS
# ---------------------------------------------------------------------------

async def descargar_archivo(
    page: Page,
    url_archivo: str,
    ruta_destino: Path,
) -> None:
    """
    Descarga un archivo XLS desde su URL directa usando Playwright.
    Intercepta el evento de descarga del navegador para guardar el archivo
    en la ruta indicada.

    Parámetros:
        page          : instancia activa de Playwright Page
        url_archivo   : URL directa del archivo .XLS
        ruta_destino  : ruta completa donde se guardará el archivo
    """
    # Crea el directorio de destino si no existe
    ruta_destino.parent.mkdir(parents=True, exist_ok=True)

    # Omite si ya existe y la opción está activada
    if OMITIR_EXISTENTES and ruta_destino.exists():
        print(f"    [SKIP] Ya existe, omitiendo: {ruta_destino.name}")
        return

    # Navega a la URL del archivo; el navegador dispara el evento de descarga
    async with page.expect_download(timeout=TIMEOUT_DESCARGA) as descarga_info:
        # Abre la URL en una nueva pestaña para no perder la página actual
        await page.evaluate(f"window.open('{url_archivo}', '_blank')")

    descarga: Download = await descarga_info.value
    await descarga.save_as(str(ruta_destino))
    print(f"    [OK] {ruta_destino.parent.name}/{ruta_destino.name}")


# ---------------------------------------------------------------------------
# ORQUESTADOR DE SUBCARPETA
# ---------------------------------------------------------------------------

async def procesar_subcarpeta(
    page: Page,
    subcarpeta: dict,
    dir_raiz_sub: Path,
) -> None:
    """
    Procesa una subcarpeta completa:
        1. Obtiene todos los enlaces XLS de su página de resultados.
        2. Descarga cada archivo organizándolo por año/mes.

    Estructura de carpetas generada:
        descargas_sbs/
          └── Información de la Banca Múltiple/
                └── Créditos Directos por Sector Económico/
                      ├── 2026/
                      │     ├── Enero.xls
                      │     ├── Febrero.xls
                      │     └── Marzo.xls
                      ├── 2025/
                      │     ├── Enero.xls
                      │     └── ...
                      └── ...
    """
    nombre = subcarpeta["texto"]
    codigo = subcarpeta["codigo"]

    print(f"\n{'='*55}")
    print(f"  [DIR] Subcarpeta: {nombre}")
    print(f"{'='*55}")

    # Obtiene la lista de archivos disponibles
    try:
        enlaces = await obtener_enlaces_xls(page, codigo)
    except Exception as e:
        print(f"  [ERROR] Error obteniendo enlaces: {e}")
        return

    if not enlaces:
        print(f"  [WARN] No se encontraron archivos XLS.")
        return

    # Aplica el filtro de periodo si está configurado
    if DESDE_PERIODO is not None:
        total_antes = len(enlaces)
        enlaces = [e for e in enlaces if periodo_es_valido(e["año"], e["mes"])]
        descartados = total_antes - len(enlaces)
        desde_str = f"{DESDE_PERIODO[1]:02d}/{DESDE_PERIODO[0]}"
        print(f"  [FILTER] Desde {desde_str} "
              f"-> {len(enlaces)} archivos ({descartados} anteriores omitidos)")

    if not enlaces:
        print(f"  [WARN] Ningun archivo cumple el filtro de periodo.")
        return

    # Descarga cada archivo
    errores = 0
    for idx, item in enumerate(enlaces):
        año   = item["año"]
        mes   = item["mes"]
        url   = item["url"]
        texto = item["texto"]

        # Nombre del archivo: usa el texto del enlace + extensión original
        ext = Path(urlparse(url).path).suffix.lower() or ".xls"
        nombre_archivo = sanitizar_nombre(f"{mes}{ext}")

        # Ruta: <raíz_sub>/<año>/<mes>.xls
        ruta_destino = dir_raiz_sub / año / nombre_archivo

        print(f"  [DOWN] [{idx + 1}/{len(enlaces)}] {año}/{mes} -> {nombre_archivo}")

        try:
            await descargar_archivo(page, url, ruta_destino)
            await asyncio.sleep(PAUSA_ENTRE_DESCARGAS)
        except Exception as e:
            print(f"    [ERROR] {e}")
            errores += 1
            continue

    print(f"\n  [OK] Subcarpeta completada. "
          f"Descargados: {len(enlaces) - errores}/{len(enlaces)}")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

async def main() -> None:
    """
    Punto de entrada. Orquesta la apertura del navegador, la expansión
    del menú y la descarga de todos los archivos configurados.
    """
    print("=" * 55)
    print("  SBS Scraper - Estadisticas Banca Multiple")
    print("=" * 55)

    async with async_playwright() as pw:

        browser = await pw.chromium.launch(
            headless=False,   # Cambia a True para ejecución silenciosa
            slow_mo=200,
        )

        context = await browser.new_context(
            accept_downloads=True,
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1280, "height": 900},
            locale="es-PE",
        )

        page = await context.new_page()

        try:
            # ── 1. Navegar al portal y expandir el menú ───────────────────
            print(f"\n[WEB] Abriendo portal SBS...")
            await page.goto(PORTAL_URL, wait_until="networkidle", timeout=60_000)
            await asyncio.sleep(2)
            await expandir_nodo_raiz(page)

            # ── 2. Procesar cada subcarpeta configurada ────────────────────
            for subcarpeta in SUBCARPETAS_OBJETIVO:

                # Ruta local que replica la jerarquía del portal:
                # descargas_sbs / <nodo raíz> / <subcarpeta> / <año> / <mes>.xls
                dir_sub = (
                    BASE_DOWNLOAD_DIR
                    / sanitizar_nombre(ROOT_NODE_TEXT)
                    / sanitizar_nombre(subcarpeta["texto"])
                )

                try:
                    await procesar_subcarpeta(page, subcarpeta, dir_sub)
                except Exception as e:
                    print(f"\n  [ERROR] Error en '{subcarpeta['texto']}': {e}")
                    continue

            # ── 3. Resumen final ───────────────────────────────────────────
            print(f"\n{'='*55}")
            print("  [DONE] Proceso completado.")
            print(f"  [DIR]  Archivos en: {BASE_DOWNLOAD_DIR.resolve()}")
            print(f"{'='*55}")

        except Exception as e:
            print(f"\n[ERROR] Error critico: {e}")
            BASE_DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
            screenshot = BASE_DOWNLOAD_DIR / "error_screenshot.png"
            await page.screenshot(path=str(screenshot), full_page=True)
            print(f"   Screenshot: {screenshot}")
            raise

        finally:
            await context.close()
            await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
