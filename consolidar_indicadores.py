"""
consolidar_indicadores.py
==========================
Consolida los archivos Excel de "Indicadores Financieros" de la SBS
en un único CSV con las columnas:

    Periodo | Indicador | Sub_Indicador | Banco | Valor

Estructura del Excel (una sola hoja, XLSX):
  - Fila 1   : título
  - Fila 2   : fecha del periodo (bold)
  - Fila 3   : "(En porcentaje)"
  - Fila 4   : a veces "Actualizado el..." (ignorar)
  - Fila 6   : cabecera de bancos (col 0 = vacío, col 1+ = nombres)
  - Fila 9+  : datos en dos niveles:
                 bold=True  → Indicador  (SOLVENCIA, CALIDAD DE ACTIVOS, etc.)
                 bold=False → Sub_Indicador
  - Pie      : "Nota:", "*", "**", "Los valores..."

Normalización de sub-indicadores:
  - Se quitan asteriscos, notas al pie (1/, 2/, *, **) y fechas entre paréntesis
  - Se aplica un mapa de sinónimos para unificar nombres que varían entre archivos

Bancos:
  - Se quita el sufijo " (con sucursales en el exterior)"
  - Se excluye "Total Banca Múltiple ..."
  - Se normaliza "B. China Perú" → "Bank of China"
  - Se quitan asteriscos del nombre
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

CARPETA = "Indicadores Financieros"

# Bancos a excluir completamente
BANCOS_EXCLUIR_PREFIJOS = [
    "total banca",
]

# Sufijos a quitar del nombre del banco
SUFIJOS_BANCO = [
    " (con sucursales en el exterior)",
]

# Normalización de nombres de banco (después de limpiar sufijos/asteriscos)
BANCO_SINONIMOS = {
    "b. china perú":  "Bank of China",
    "b. china peru":  "Bank of China",
}

# Palabras clave de pie de página
FOOTER_STARTS = ["nota", "los valores", "* para", "** en el", "** inform",
                 "*** mediante", "a los 30", "hayan estado"]

# Meses en español → número de mes
MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "setiembre": 9, "septiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}

# ---------------------------------------------------------------------------
# Mapa de uniformización de sub-indicadores
# Clave: texto normalizado (lower, sin asteriscos, sin fechas)
# Valor: nombre canónico
# ---------------------------------------------------------------------------
SUBINDICADOR_CANON = {
    # SOLVENCIA
    "ratio de capital global":
        "Ratio de Capital Global",
    "pasivo total / capital social y reservas ( n° de veces )":
        "Pasivo Total / Capital Social y Reservas (N° de veces)",
    "pasivo total / capital social y reservas ( nº de veces )":
        "Pasivo Total / Capital Social y Reservas (N° de veces)",

    # CALIDAD DE ACTIVOS
    "créditos atrasados (criterio sbs) / créditos directos":
        "Créditos Atrasados (criterio SBS) / Créditos Directos",
    "créditos atrasados con más de 90 días de atraso / créditos directos":
        "Créditos Atrasados > 90 días / Créditos Directos",
    "créditos atrasados con mas de 90 dias de atraso / creditos directos":
        "Créditos Atrasados > 90 días / Créditos Directos",
    "créditos refinanciados y reestructurados / créditos directos":
        "Créditos Refinanciados y Reestructurados / Créditos Directos",
    "créditos atrasados mn (criterio sbs) / créditos directos mn":
        "Créditos Atrasados MN (criterio SBS) / Créditos Directos MN",
    "créditos atrasados me (criterio sbs) / créditos directos me":
        "Créditos Atrasados ME (criterio SBS) / Créditos Directos ME",
    "provisiones / créditos atrasados":
        "Provisiones / Créditos Atrasados",
    "cartera atrasada ajustada":
        "Cartera Atrasada Ajustada",
    "cartera de alto riesgo ajustada":
        "Cartera de Alto Riesgo Ajustada",

    # EFICIENCIA Y GESTIÓN
    "gastos de administración anualizados / activo productivo promedio":
        "Gastos de Administración Anualizados / Activo Productivo Promedio",
    "gastos de operación / margen financiero total":
        "Gastos de Operación / Margen Financiero Total",
    "ingresos financieros / ingresos totales":
        "Ingresos Financieros / Ingresos Totales",
    "ingresos financieros anualizados / activo productivo promedio":
        "Ingresos Financieros Anualizados / Activo Productivo Promedio",
    "créditos directos / personal ( s/ miles )":
        "Créditos Directos / Personal (S/ Miles)",
    "depósitos / número de oficinas ( s/ miles )":
        "Depósitos / Número de Oficinas (S/ Miles)",

    # RENTABILIDAD
    "utilidad neta anualizada / patrimonio promedio":
        "Utilidad Neta Anualizada / Patrimonio Promedio",
    "utilidad neta anualizada / activo promedio":
        "Utilidad Neta Anualizada / Activo Promedio",

    # LIQUIDEZ
    "ratio de liquidez mn (promedio de saldos del mes)":
        "Ratio de Liquidez MN",
    "ratio de liquidez me (promedio de saldos del mes)":
        "Ratio de Liquidez ME",
    "caja y bancos mn / obligaciones a la vista mn ( n° de veces )":
        "Caja y Bancos MN / Obligaciones a la Vista MN (N° de veces)",
    "caja y bancos mn / obligaciones a la vista mn ( nº de veces )":
        "Caja y Bancos MN / Obligaciones a la Vista MN (N° de veces)",
    "caja y bancos en me / obligaciones a la vista me ( n° de veces)":
        "Caja y Bancos ME / Obligaciones a la Vista ME (N° de veces)",
    "caja y bancos en me / obligaciones a la vista me ( nº de veces)":
        "Caja y Bancos ME / Obligaciones a la Vista ME (N° de veces)",
}

# Normalización de nombres de Indicador (bold)
INDICADOR_CANON = {
    "eficiencia y gestión":  "EFICIENCIA Y GESTIÓN",
    "eficiencia y gestion":  "EFICIENCIA Y GESTIÓN",
    "calidad de activos":    "CALIDAD DE ACTIVOS",
    "solvencia":             "SOLVENCIA",
    "rentabilidad":          "RENTABILIDAD",
    "liquidez":              "LIQUIDEZ",
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

def periodo_desde_path(xls_path: Path) -> str:
    year  = xls_path.parent.name
    month = xls_path.stem.lower()
    mes_num = MESES.get(month)
    if mes_num and year.isdigit():
        return f"{year}{mes_num:02d}"
    return f"{year}{xls_path.stem}"


def normalizar_unicode(texto: str) -> str:
    """Corrige caracteres mal codificados comunes en estos archivos."""
    return (texto
        .replace("Ë", "Ó").replace("ë", "ó")   # GESTIÓN
        .replace("░", "°").replace("·", "ú")
        .replace("ß", "á").replace("Ý", "í")
        .replace("¾", "ó").replace("±", "ñ")
        .replace("Ú", "é").replace("ú", "ú")
    )


def limpiar_texto(texto: str) -> str:
    """Normaliza espacios, quita asteriscos y notas al pie."""
    s = normalizar_unicode(str(texto))
    s = re.sub(r"\s+", " ", s).strip()
    # Quitar asteriscos al final (**, ***, *)
    s = re.sub(r"\s*\*+\s*$", "", s).strip()
    # Quitar notas tipo "1/", "2/" al final
    s = re.sub(r"\s+\d+/\s*$", "", s).strip()
    # Quitar fechas entre paréntesis al final: "(al 31/12/2023)", "(al 28/02/2026)"
    s = re.sub(r"\s*\(al \d{2}/\d{2}/\d{4}\)\s*$", "", s).strip()
    return s


def canonizar_subindicador(texto: str) -> str:
    """Busca el nombre canónico del sub-indicador, o devuelve el texto limpio."""
    limpio = limpiar_texto(texto)
    # Normalizar para búsqueda: lower, quitar paréntesis de asteriscos
    clave = limpio.lower()
    # Quitar "(criterio sbs)*" → "(criterio sbs)"
    clave = re.sub(r"\(criterio sbs\)\*", "(criterio sbs)", clave)
    # Quitar asteriscos sueltos
    clave = re.sub(r"\*+", "", clave).strip()
    # Normalizar espacios múltiples
    clave = re.sub(r"\s+", " ", clave).strip()
    return SUBINDICADOR_CANON.get(clave, limpio)


def canonizar_indicador(texto: str) -> str:
    """Devuelve el nombre canónico del Indicador (bold)."""
    limpio = limpiar_texto(texto)
    clave = re.sub(r"\*+", "", limpio.lower()).strip()
    return INDICADOR_CANON.get(clave, limpio.upper())


def limpiar_banco(nombre: str) -> str:
    """Normaliza nombre de banco."""
    s = re.sub(r"\s+", " ", nombre).strip()
    s = re.sub(r"\*+\s*$", "", s).strip()  # quitar asteriscos finales
    s = s.rstrip(".").strip()
    # Quitar sufijos de sucursales
    for sufijo in SUFIJOS_BANCO:
        if s.lower().endswith(sufijo.lower()):
            s = s[: -len(sufijo)].strip()
    # Corregir 'B.Nombre' → 'B. Nombre'
    s = re.sub(r"^B\.([A-ZÁÉÍÓÚÑ])", r"B. \1", s)
    # Aplicar sinónimos
    clave = s.lower()
    if clave in BANCO_SINONIMOS:
        s = BANCO_SINONIMOS[clave]
    return s


def es_banco_excluido(nombre: str) -> bool:
    norm = re.sub(r"\s+", " ", nombre).strip().lower()
    return any(norm.startswith(p) for p in BANCOS_EXCLUIR_PREFIJOS)


def es_footer(valor) -> bool:
    if pd.isna(valor):
        return False
    s = normalizar_unicode(str(valor)).strip().lower()
    return any(s.startswith(kw) for kw in FOOTER_STARTS)


def abrir_archivo(xls_path: Path) -> tuple[pd.DataFrame, dict[int, bool], Path | None]:
    """
    Lee el archivo y devuelve (DataFrame, bold_map, tmp_path).
    bold_map: {fila_1based: bool} para col 0.
    """
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.close()
    tmp_path = Path(tmp.name)
    shutil.copy(xls_path, tmp_path)

    df = None
    bold_map = {}

    # Intentar openpyxl
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df = pd.read_excel(tmp_path, sheet_name=0, header=None, engine="openpyxl")

        # Leer estilos bold
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                wb = openpyxl.load_workbook(tmp_path, data_only=True)
            ws = wb.worksheets[0]
            for row in ws.iter_rows(min_col=1, max_col=1):
                cell = row[0]
                bold_map[cell.row] = bool(cell.font and cell.font.bold)
            wb.close()
        except Exception:
            pass  # sin bold_map, usaremos fallback

        return df, bold_map, tmp_path

    except Exception:
        pass

    # Fallback: xlrd (XLS binario)
    try:
        tmp_path.unlink()
    except Exception:
        pass

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df = pd.read_excel(xls_path, sheet_name=0, header=None, engine="xlrd")

    return df, bold_map, None


# ---------------------------------------------------------------------------
# Procesamiento de un archivo
# ---------------------------------------------------------------------------

def procesar_archivo(xls_path: Path) -> pd.DataFrame | None:
    periodo = periodo_desde_path(xls_path)

    try:
        df, bold_map, tmp_path = abrir_archivo(xls_path)
    except Exception as e:
        log.error("✗ No se pudo abrir '%s': %s", xls_path.name, e)
        return None
    finally:
        pass  # tmp_path se limpia dentro de abrir_archivo si falla

    try:
        # ------------------------------------------------------------------
        # 1. Leer cabecera de bancos (fila índice 5)
        # ------------------------------------------------------------------
        fila_bancos = df.iloc[5]
        bancos = {}  # {col_idx: nombre_limpio}
        for j in range(1, df.shape[1]):
            v = fila_bancos.iloc[j]
            if pd.notna(v) and str(v).strip():
                nombre_orig = str(v).strip()
                if not es_banco_excluido(nombre_orig):
                    bancos[j] = limpiar_banco(nombre_orig)

        if not bancos:
            log.warning("  Sin bancos en '%s'", xls_path.name)
            return None

        # ------------------------------------------------------------------
        # 2. Iterar filas de datos (desde índice 8)
        # ------------------------------------------------------------------
        registros = []
        indicador_actual    = None
        subindicador_actual = None

        for i in range(8, len(df)):
            v0 = df.iloc[i, 0]

            # Pie de página → parar
            if es_footer(v0):
                break

            # Fila vacía → separador
            if pd.isna(v0) or not str(v0).strip():
                continue

            texto_raw = str(v0).strip()

            # Determinar si es Indicador (bold) o Sub_Indicador
            fila_openpyxl = i + 1  # 0-based → 1-based
            es_bold = bold_map.get(fila_openpyxl, False)

            # Fallback si no hay bold_map: texto en MAYÚSCULAS = Indicador
            if not bold_map:
                texto_norm = normalizar_unicode(texto_raw)
                es_bold = texto_norm == texto_norm.upper() and any(c.isalpha() for c in texto_norm)

            if es_bold:
                indicador_actual    = canonizar_indicador(texto_raw)
                subindicador_actual = None
                continue  # la fila del indicador no tiene valores propios
            else:
                subindicador_actual = canonizar_subindicador(texto_raw)

            if indicador_actual is None:
                continue

            # Leer valores por banco
            for col_idx, banco in bancos.items():
                if col_idx >= df.shape[1]:
                    continue

                v = df.iloc[i, col_idx]

                if pd.isna(v) or str(v).strip() in ("-", ""):
                    valor = None
                else:
                    try:
                        valor = float(v)
                    except (ValueError, TypeError):
                        valor = None

                if valor is None:
                    continue  # omitir celdas vacías o "-"

                registros.append({
                    "Periodo":       periodo,
                    "Indicador":     indicador_actual,
                    "Sub_Indicador": subindicador_actual,
                    "Banco":         banco,
                    "Valor":         valor,
                })

    finally:
        if tmp_path and tmp_path.exists():
            try:
                tmp_path.unlink()
            except Exception:
                pass

    if not registros:
        return None

    return pd.DataFrame(registros)


# ---------------------------------------------------------------------------
# Función principal
# ---------------------------------------------------------------------------

def main():
    carpeta = BASE_DIR / CARPETA
    if not carpeta.exists():
        log.error("No se encontró '%s' dentro de '%s'.", CARPETA, BASE_DIR)
        return

    archivos = sorted(carpeta.rglob("*.xls"))
    if not archivos:
        log.warning("No se encontraron archivos .xls en '%s'", carpeta)
        return

    log.info("Procesando %d archivos de: %s", len(archivos), carpeta.name)
    log.info("─" * 60)

    frames = []
    for xls_path in archivos:
        log.info("  %s", xls_path.relative_to(BASE_DIR))
        df = procesar_archivo(xls_path)
        if df is not None:
            frames.append(df)
            log.info("    → %d registros", len(df))
        else:
            log.warning("    → Sin datos")

    if not frames:
        log.error("No se extrajo ningún dato.")
        return

    df_final = pd.concat(frames, ignore_index=True)
    df_final.sort_values(
        ["Periodo", "Indicador", "Sub_Indicador", "Banco"],
        inplace=True,
        na_position="first",
    )
    df_final.reset_index(drop=True, inplace=True)

    nombre_csv = "Indicadores_Financieros.csv"
    ruta_csv   = OUTPUT_DIR / nombre_csv
    df_final.to_csv(ruta_csv, index=False, encoding="utf-8-sig")

    log.info("═" * 60)
    log.info("✓ Exportado: %s", ruta_csv.name)
    log.info("  Filas: %d  |  Columnas: %s", len(df_final), df_final.columns.tolist())
    log.info("  Periodos: %d  |  Indicadores: %d  |  Bancos: %d",
             df_final["Periodo"].nunique(),
             df_final["Indicador"].nunique(),
             df_final["Banco"].nunique())
    log.info("  Indicadores únicos:")
    for ind in sorted(df_final["Indicador"].unique()):
        subs = df_final[df_final["Indicador"] == ind]["Sub_Indicador"].dropna().unique()
        log.info("    [%s]", ind)
        for s in sorted(subs):
            log.info("      - %s", s)
    log.info("  Bancos únicos:")
    for b in sorted(df_final["Banco"].unique()):
        log.info("    %s", b)


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    main()
