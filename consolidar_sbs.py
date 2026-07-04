"""
consolidar_sbs.py
=================
Script de consolidación de archivos Excel de la SBS (Banca Múltiple).

Estructura de los archivos (Balance General y similares):
  - Filas 1-4  : encabezado del reporte (título, fecha, unidad)
  - Fila 5     : sección/subtítulo (ej. "Activo", "Pasivo") — bold, no UPPER
  - Fila 6     : nombres de bancos (cabecera nivel 1, celdas combinadas)
  - Fila 7     : MN / ME / TOTAL (cabecera nivel 2)
  - Fila 8     : vacía
  - Fila 9+    : datos
  - Pie        : "Tipo de Cambio", "1/", "Nota", "*", etc.

  NOTA: Algunas hojas tienen TABLAS APILADAS (ej. hoja 1 tiene "Activo" y
  luego "Pasivo" con su propio bloque de cabeceras de bancos).

Cambios aplicados vs versión anterior:
  1. Excluir bancos: "Total Banca Múltiple", "Banco de Crédito con Sucursales
     en el Exterior", "Total Banca Múltiple Incluye Sucursales en el Exterior"
  2. Periodo en formato YYYYMM (ej. 202401)
  3. Para la categoría "Balance General y Estado de Ganancias y Pérdidas":
     solo conservar filas con Bold=True en la columna Concepto_Cuenta.
     Para las demás categorías: comportamiento estándar (todas las filas).

Salida: un CSV por Categoría en Output_SBS/.
"""

import re
import shutil
import tempfile
import logging
import warnings
from pathlib import Path

import pandas as pd
import openpyxl

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
BASE_DIR   = Path(__file__).parent / "Data_SBS"
OUTPUT_DIR = Path(__file__).parent / "Output_SBS"
OUTPUT_DIR.mkdir(exist_ok=True)

# Bancos a excluir (comparación case-insensitive, strip)
BANCOS_EXCLUIR = {
    "total banca múltiple",
    "total banca multiple",
    "banco de crédito con sucursales en el exterior",
    "banco de credito con sucursales en el exterior",
    "total banca múltiple incluye sucursales en el exterior",
    "total banca multiple incluye sucursales en el exterior",
}

# Categorías que requieren filtro por negrita en Concepto_Cuenta
CATEGORIAS_FILTRO_BOLD = {
    "Balance General y Estado de Ganancias y Pérdidas",
    "Balance General y Estado de Ganancias y Perdidas",  # versión sin tilde (Data_SBS)
}

# Palabras clave que marcan el inicio del pie de página (case-insensitive)
FOOTER_KEYWORDS = [
    "tipo de cambio",
    "nota",
    "fuente",
    "1/",
    "2/",
    "3/",
    "4/",
    "5/",
    "incluye",
    "considera",
    "corresponde",
    "cifras",
    "elaboración",
    "elaboracion",
]

# Meses en español → número de mes
MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "setiembre": 9, "septiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def slugify(text: str) -> str:
    """Convierte un texto en un nombre de archivo seguro."""
    text = text.strip()
    text = re.sub(r"[áàäâÁÀÄÂ]", "a", text)
    text = re.sub(r"[éèëêÉÈËÊ]", "e", text)
    text = re.sub(r"[íìïîÍÌÏÎ]", "i", text)
    text = re.sub(r"[óòöôÓÒÖÔ]", "o", text)
    text = re.sub(r"[úùüûÚÙÜÛ]", "u", text)
    text = re.sub(r"[ñÑ]", "n", text)
    text = re.sub(r"[^a-zA-Z0-9\s_]", "", text)
    text = re.sub(r"\s+", "_", text)
    return text


def periodo_desde_path(xls_path: Path) -> str:
    """
    Construye el periodo en formato YYYYMM a partir de la ruta del archivo.
    Ejemplo: .../2024/Enero.xls → '202401'
    """
    year  = xls_path.parent.name      # '2024'
    month = xls_path.stem.lower()     # 'enero'
    mes_num = MESES.get(month)
    if mes_num and year.isdigit():
        return f"{year}{mes_num:02d}"
    return f"{year}{xls_path.stem}"


def normalizar_banco(nombre_banco: str) -> str:
    """
    Normaliza el nombre de un banco:
    - Quita asteriscos (notas al pie) del nombre
    - Normaliza espacios múltiples y saltos de línea
    Ej: 'Compartamos Banco*' → 'Compartamos Banco'
    """
    nombre = re.sub(r"\s+", " ", nombre_banco).strip()
    nombre = nombre.rstrip("*").strip()
    return nombre


def es_banco_excluido(nombre_banco: str) -> bool:
    """Devuelve True si el banco debe ser excluido del resultado."""
    # Normalizar primero (quita asteriscos, espacios, saltos de línea)
    normalizado = normalizar_banco(nombre_banco).lower()
    return normalizado in BANCOS_EXCLUIR


def es_fila_footer(valor_col0) -> bool:
    """Devuelve True si la fila pertenece al pie de página."""
    if pd.isna(valor_col0):
        return False
    texto = str(valor_col0).strip().lower()
    if not texto:
        return False
    if texto.startswith("*"):
        return True
    for kw in FOOTER_KEYWORDS:
        if texto.startswith(kw):
            return True
    return False


def abrir_xls_como_xlsx(xls_path: Path) -> tuple[pd.ExcelFile, Path]:
    """
    Los archivos de la SBS pueden ser:
      a) Archivos .xls con formato XLSX dentro (los nuevos) → openpyxl
      b) Archivos .xls con formato BIFF/binario real (los viejos) → xlrd

    Estrategia:
      1. Intenta copiar a .xlsx y abrir con openpyxl (más completo: lee estilos bold).
      2. Si falla, abre directamente con xlrd (sin estilos).

    Devuelve (ExcelFile, ruta_temporal_o_None).
    """
    # ── Intento 1: copiar a .xlsx y abrir con openpyxl ───────────────────
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.close()
    tmp_path = Path(tmp.name)
    try:
        shutil.copy(xls_path, tmp_path)
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            xl = pd.ExcelFile(tmp_path, engine="openpyxl")
        # Verificación rápida: intentar listar hojas
        _ = xl.sheet_names
        return xl, tmp_path
    except Exception:
        try:
            tmp_path.unlink()
        except Exception:
            pass

    # ── Intento 2: abrir directamente con xlrd (XLS binario real) ────────
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        xl = pd.ExcelFile(xls_path, engine="xlrd")
    return xl, None


# ---------------------------------------------------------------------------
# Extracción de mapa de negrita por hoja
# ---------------------------------------------------------------------------

def obtener_bold_map(tmp_xlsx: Path | None, nombre_hoja: str) -> dict[int, bool]:
    """
    Lee los estilos de celda de la columna A de una hoja con openpyxl.
    Si tmp_xlsx es None (archivo leído con xlrd), devuelve dict vacío
    y el parseador usará el fallback de indentación.
    """
    bold_map: dict[int, bool] = {}
    if tmp_xlsx is None:
        return bold_map
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(tmp_xlsx, data_only=True)
        ws = wb[nombre_hoja]
        for row in ws.iter_rows(min_col=1, max_col=1):
            cell = row[0]
            bold_map[cell.row] = bool(cell.font and cell.font.bold)
        wb.close()
    except Exception as e:
        log.warning("    No se pudo leer estilos de '%s': %s", nombre_hoja, e)
    return bold_map


# ---------------------------------------------------------------------------
# Parseo de un bloque de tabla dentro de una hoja
# ---------------------------------------------------------------------------

def _extraer_col_total_map(fila_bancos_raw, fila_moneda_raw) -> dict[int, str]:
    """
    Dado los valores de la fila de bancos y la fila de moneda,
    construye el mapa {col_idx: nombre_banco_normalizado} para columnas TOTAL,
    excluyendo los bancos no deseados.
    Los nombres de banco se normalizan (sin asteriscos, sin espacios extra).
    """
    banco_actual = None
    bancos_ff = []
    for v in fila_bancos_raw:
        if pd.notna(v) and str(v).strip():
            banco_actual = str(v).strip()
        bancos_ff.append(banco_actual)

    col_total_map: dict[int, str] = {}
    for col_idx, (banco, moneda) in enumerate(zip(bancos_ff, fila_moneda_raw)):
        if col_idx == 0:
            continue
        moneda_str = str(moneda).strip().upper() if pd.notna(moneda) else ""
        if moneda_str == "TOTAL" and banco and not es_banco_excluido(banco):
            col_total_map[col_idx] = normalizar_banco(banco)  # ← normalizar aquí

    return col_total_map


def _extraer_nombre_tabla(fila_bancos_raw) -> str:
    """
    Extrae el nombre de la tabla/sección desde col 0 de la fila de bancos.
    En la hoja 1: col 0 contiene 'Activo' o 'Pasivo'.
    En la hoja 2 (Estado de Ganancias): col 0 está vacía → 'Estado de Ganancias y Pérdidas'.
    """
    val = fila_bancos_raw.iloc[0]
    if pd.notna(val) and str(val).strip():
        return str(val).strip()
    return "Estado de Ganancias y Pérdidas"


def parsear_hoja(
    df_raw: pd.DataFrame,
    tmp_xlsx: Path,
    nombre_hoja: str,
    periodo: str,
    categoria: str,
    aplicar_filtro_bold: bool,
) -> pd.DataFrame | None:
    """
    Procesa una hoja cruda (sin cabeceras) y devuelve un DataFrame largo con:
    Periodo | Categoría | Hoja | Concepto_Cuenta | Banco | Monto_Total

    Maneja tablas apiladas verticalmente dentro de la misma hoja:
    cada bloque tiene su propia fila de bancos (fila con texto no vacío en
    posiciones de banco) y fila de MN/ME/TOTAL.

    Estructura de cada bloque:
      fila N   : sección (ej. "Activo") — bold, col 0 no vacía, resto vacío
      fila N+1 : nombres de bancos (col 1+ con nombres de banco)
      fila N+2 : MN / ME / TOTAL
      fila N+3 : vacía (separador)
      fila N+4+: datos
    """
    if df_raw.shape[1] < 4:
        return None

    # Obtener mapa de negrita si se necesita
    bold_map: dict[int, bool] = {}
    if aplicar_filtro_bold:
        bold_map = obtener_bold_map(tmp_xlsx, nombre_hoja)

    # ------------------------------------------------------------------
    # Detectar bloques de tablas apiladas
    # ------------------------------------------------------------------
    # Un bloque nuevo comienza cuando encontramos una fila que:
    #   - tiene texto en col 1+ que parece nombre de banco (no MN/ME/TOTAL)
    # La estructura fija es:
    #   idx 5 → bancos, idx 6 → MN/ME/TOTAL  (primer bloque)
    #   luego puede haber otro bloque más adelante con la misma estructura
    #
    # Estrategia: buscar todas las filas donde la fila siguiente contiene
    # "MN" o "ME" o "TOTAL" en las columnas de datos.

    MONEDAS = {"MN", "ME", "TOTAL"}

    def es_fila_cabecera_moneda(row_series) -> bool:
        """True si la fila contiene MN/ME/TOTAL en columnas de datos."""
        vals = [str(v).strip().upper() for v in row_series.iloc[1:] if pd.notna(v) and str(v).strip()]
        return bool(vals) and all(v in MONEDAS for v in vals)

    # Encontrar índices de filas de moneda (MN/ME/TOTAL)
    indices_moneda = []
    for i in range(len(df_raw)):
        if es_fila_cabecera_moneda(df_raw.iloc[i]):
            indices_moneda.append(i)

    if not indices_moneda:
        log.debug("  Hoja '%s': no se encontró fila MN/ME/TOTAL", nombre_hoja)
        return None

    # Cada bloque: fila_bancos = idx_moneda - 1, datos desde idx_moneda + 2
    # El bloque termina justo antes del siguiente bloque (o al final del df)
    bloques = []
    for k, idx_moneda in enumerate(indices_moneda):
        idx_bancos = idx_moneda - 1
        idx_datos_inicio = idx_moneda + 2  # saltar fila vacía separadora

        # Fin del bloque: inicio del siguiente bloque de bancos - 3 filas
        if k + 1 < len(indices_moneda):
            idx_datos_fin = indices_moneda[k + 1] - 2  # antes de la fila de sección
        else:
            idx_datos_fin = len(df_raw)

        bloques.append((idx_bancos, idx_moneda, idx_datos_inicio, idx_datos_fin))

    # ------------------------------------------------------------------
    # Procesar cada bloque
    # ------------------------------------------------------------------
    resultados_bloques = []

    for idx_bancos, idx_moneda, idx_datos_inicio, idx_datos_fin in bloques:
        if idx_bancos < 0 or idx_datos_inicio >= len(df_raw):
            continue

        fila_bancos_raw = df_raw.iloc[idx_bancos]
        fila_moneda_raw = df_raw.iloc[idx_moneda]

        col_total_map = _extraer_col_total_map(fila_bancos_raw, fila_moneda_raw)
        if not col_total_map:
            continue

        # Nombre de la tabla/sección (Activo, Pasivo, Estado de Ganancias...)
        nombre_tabla = _extraer_nombre_tabla(fila_bancos_raw)

        # Extraer filas de datos del bloque
        df_bloque = df_raw.iloc[idx_datos_inicio:idx_datos_fin].copy()

        # Limpiar pie de página dentro del bloque
        corte = len(df_bloque)
        for i, val in enumerate(df_bloque.iloc[:, 0]):
            if es_fila_footer(val):
                corte = i
                break
        df_bloque = df_bloque.iloc[:corte]

        # Eliminar filas completamente nulas
        df_bloque = df_bloque.dropna(how="all")

        if df_bloque.empty:
            continue

        # Construir tabla ancha: concepto + columnas TOTAL
        cols_keep = [0] + list(col_total_map.keys())
        # Asegurarse de que los índices de columna existen
        cols_keep = [c for c in cols_keep if c < df_bloque.shape[1]]
        df_wide = df_bloque.iloc[:, cols_keep].copy()

        new_cols = ["Concepto_Cuenta"] + [
            col_total_map[c] for c in cols_keep if c != 0
        ]
        df_wide.columns = new_cols

        # Limpiar concepto
        df_wide["Concepto_Cuenta"] = df_wide["Concepto_Cuenta"].astype(str).str.strip()
        df_wide = df_wide[df_wide["Concepto_Cuenta"].str.lower().ne("nan")]
        df_wide = df_wide[df_wide["Concepto_Cuenta"].str.strip().ne("")]

        if df_wide.empty:
            continue

        # ------------------------------------------------------------------
        # Filtro por negrita (solo para categorías específicas)
        # ------------------------------------------------------------------
        if aplicar_filtro_bold:
            if bold_map:
                # Usar estilos reales de openpyxl
                # df_bloque índices son 0-based del df_raw original
                # openpyxl usa 1-based → fila_openpyxl = idx_original + 1
                # Reconstruir máscara solo para las filas que quedaron en df_wide
                idx_wide = df_wide.index.tolist()
                mascara_bold_wide = [
                    bold_map.get(idx + 1, False) for idx in idx_wide
                ]
                df_wide = df_wide[mascara_bold_wide].copy()
            else:
                # Fallback: los conceptos con espacios iniciales son nivel 3 (no bold)
                # Los sin espacios iniciales son nivel 1 o 2 (bold)
                log.debug(
                    "    Usando fallback espacios para filtro bold en hoja '%s'",
                    nombre_hoja,
                )
                mascara_bold_wide = [
                    not str(c).startswith("   ")
                    for c in df_wide["Concepto_Cuenta"]
                ]
                df_wide = df_wide[mascara_bold_wide].copy()

            # (ya aplicada arriba en cada rama del if/else)

            if df_wide.empty:
                continue

        # ------------------------------------------------------------------
        # Melt → formato largo
        # ------------------------------------------------------------------
        bancos_cols = [c for c in df_wide.columns if c != "Concepto_Cuenta"]

        df_long = pd.melt(
            df_wide,
            id_vars=["Concepto_Cuenta"],
            value_vars=bancos_cols,
            var_name="Banco",
            value_name="Monto_Total",
        )

        df_long.insert(0, "Periodo",   periodo)
        df_long.insert(1, "Categoría", categoria)
        df_long.insert(2, "Hoja",      nombre_hoja)
        df_long.insert(3, "Tabla",     nombre_tabla)

        df_long["Monto_Total"] = pd.to_numeric(df_long["Monto_Total"], errors="coerce")
        df_long.dropna(subset=["Monto_Total"], inplace=True)

        if not df_long.empty:
            resultados_bloques.append(df_long)

    if not resultados_bloques:
        return None

    return pd.concat(resultados_bloques, ignore_index=True)


# ---------------------------------------------------------------------------
# Procesamiento de un archivo XLS
# ---------------------------------------------------------------------------

def procesar_archivo(xls_path: Path, categoria: str) -> pd.DataFrame | None:
    """
    Abre un archivo XLS, itera por todas sus hojas y devuelve
    un DataFrame consolidado de todas las hojas procesadas.
    """
    periodo = periodo_desde_path(xls_path)
    aplicar_filtro_bold = categoria in CATEGORIAS_FILTRO_BOLD
    tmp_path = None
    resultados = []

    try:
        xl_file, tmp_path = abrir_xls_como_xlsx(xls_path)

        for nombre_hoja in xl_file.sheet_names:
            try:
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    # No pasar engine si xl_file ya lo tiene configurado
                    df_raw = pd.read_excel(
                        xl_file,
                        sheet_name=nombre_hoja,
                        header=None,
                    )
                df_hoja = parsear_hoja(
                    df_raw,
                    tmp_path,
                    nombre_hoja,
                    periodo,
                    categoria,
                    aplicar_filtro_bold,
                )
                if df_hoja is not None:
                    resultados.append(df_hoja)
                    log.debug(
                        "    ✓ Hoja '%s': %d filas procesadas",
                        nombre_hoja, len(df_hoja),
                    )
            except Exception as e:
                log.warning(
                    "  ⚠ Error en hoja '%s' de '%s': %s",
                    nombre_hoja, xls_path.name, e,
                )

        xl_file.close()

    except Exception as e:
        log.error("✗ No se pudo abrir '%s': %s", xls_path, e)
        return None

    finally:
        if tmp_path and tmp_path.exists():
            try:
                tmp_path.unlink()
            except Exception:
                pass

    if not resultados:
        return None

    return pd.concat(resultados, ignore_index=True)


# ---------------------------------------------------------------------------
# Función principal
# ---------------------------------------------------------------------------

def main():
    if not BASE_DIR.exists():
        log.error(
            "El directorio base '%s' no existe. "
            "Crea la carpeta 'Data_SBS' junto a este script y coloca "
            "los archivos descargados de la SBS dentro.",
            BASE_DIR,
        )
        return

    log.info("Iniciando consolidación desde: %s", BASE_DIR)

    # Agrupar archivos por categoría
    # Estructura: Data_SBS/<Categoría>/<Año>/<Mes>.xls
    # NOTA: Algunas carpetas pueden tener variantes con/sin tilde en el nombre.
    # Normalizamos el nombre de categoría para agruparlas bajo el mismo CSV.
    NORMALIZAR_CATEGORIA = {
        "Balance General y Estado de Ganancias y Perdidas":
            "Balance General y Estado de Ganancias y Pérdidas",
    }

    archivos_por_categoria: dict[str, list[Path]] = {}

    for xls_path in sorted(BASE_DIR.rglob("*.xls")):
        partes = xls_path.relative_to(BASE_DIR).parts
        categoria_raw = partes[0] if len(partes) >= 2 else "General"
        # Normalizar variantes con/sin tilde al nombre canónico
        categoria = NORMALIZAR_CATEGORIA.get(categoria_raw, categoria_raw)
        archivos_por_categoria.setdefault(categoria, []).append(xls_path)

    if not archivos_por_categoria:
        log.warning("No se encontraron archivos .xls en '%s'", BASE_DIR)
        return

    log.info("Categorías encontradas: %d", len(archivos_por_categoria))

    for categoria, archivos in sorted(archivos_por_categoria.items()):
        log.info("─" * 60)
        log.info("Categoría: %s  (%d archivos)", categoria, len(archivos))

        frames_categoria = []

        for xls_path in archivos:
            log.info("  Procesando: %s", xls_path.relative_to(BASE_DIR))
            df = procesar_archivo(xls_path, categoria)
            if df is not None:
                frames_categoria.append(df)
                log.info("    → %d registros", len(df))
            else:
                log.warning("    → Sin datos extraídos")

        if not frames_categoria:
            log.warning("  Sin datos para '%s'. Se omite.", categoria)
            continue

        df_cat = pd.concat(frames_categoria, ignore_index=True)

        # Deduplicar: si dos carpetas tenían los mismos periodos (ej. con y sin tilde),
        # conservar solo la primera ocurrencia por clave completa.
        df_cat.drop_duplicates(
            subset=["Periodo", "Hoja", "Tabla", "Banco", "Concepto_Cuenta"],
            keep="first",
            inplace=True,
        )

        df_cat.sort_values(
            ["Periodo", "Hoja", "Tabla", "Banco", "Concepto_Cuenta"],
            inplace=True,
        )
        df_cat.reset_index(drop=True, inplace=True)

        nombre_csv = slugify(categoria) + ".csv"
        ruta_csv   = OUTPUT_DIR / nombre_csv
        df_cat.to_csv(ruta_csv, index=False, encoding="utf-8-sig")

        log.info(
            "  ✓ Exportado: %s  (%d filas, %d columnas)",
            ruta_csv.name, len(df_cat), len(df_cat.columns),
        )

    log.info("═" * 60)
    log.info("Consolidación completada. Archivos en: %s", OUTPUT_DIR)


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    main()
