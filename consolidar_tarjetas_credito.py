"""
consolidar_tarjetas_credito.py
===============================
Consolida los archivos Excel de "Número de Tarjetas de Crédito por Tipo"
de la SBS en un único CSV con las columnas:

    Periodo | Banco | Nro_TC

Solo se extrae la columna "Créditos de Consumo" (col 1 en los datos).

Estructura del Excel (una sola hoja, XLSX):
  - Fila 1   : título
  - Fila 2   : fecha del periodo
  - Fila 4   : cabecera de tipos → col 0 = "Empresas",
               col 1 = "Créditos de Consumo", col 2..6 = otros tipos, col 7 = "Total"
  - Fila 6+  : datos → col 0 = Banco, col 1 = Nro_TC Consumo
  - Pie      : "TOTAL BANCA MÚLTIPLE" (excluir), "Fuente:", "Nota:", etc.
"""

import re
import shutil
import tempfile
import logging
import warnings
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
BASE_DIR   = Path(__file__).parent / "Data_SBS"
OUTPUT_DIR = Path(__file__).parent / "Output_SBS"
OUTPUT_DIR.mkdir(exist_ok=True)

CARPETA = "Número de Tarjetas de Crédito por Tipo"

# Columna objetivo en la cabecera (normalizada)
COL_OBJETIVO = "créditos de consumo"

# Bancos a excluir
BANCOS_EXCLUIR_PREFIJOS = ["total banca"]

# Pie de página
FOOTER_STARTS = ["fuente", "nota", "aprobado", "* mediante", "la resolución"]

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

def periodo_desde_path(xls_path: Path) -> str:
    year  = xls_path.parent.name
    month = xls_path.stem.lower()
    mes_num = MESES.get(month)
    if mes_num and year.isdigit():
        return f"{year}{mes_num:02d}"
    return f"{year}{xls_path.stem}"


def limpiar_banco(nombre: str) -> str:
    """Normaliza nombre de banco: quita asteriscos, puntos finales,
    sufijo de sucursales y corrige 'B.Nombre' → 'B. Nombre'."""
    s = re.sub(r"\s+", " ", nombre).strip()
    s = re.sub(r"\*+\s*$", "", s).strip()
    s = s.rstrip(".").strip()
    sufijo = " (con sucursales en el exterior)"
    if s.lower().endswith(sufijo):
        s = s[: -len(sufijo)].strip()
    s = re.sub(r"^B\.([A-ZÁÉÍÓÚÑ])", r"B. \1", s)
    return s


def es_banco_excluido(nombre: str) -> bool:
    norm = re.sub(r"\s+", " ", nombre).strip().lower()
    return any(norm.startswith(p) for p in BANCOS_EXCLUIR_PREFIJOS)


def es_footer(valor) -> bool:
    if pd.isna(valor):
        return False
    s = str(valor).strip().lower()
    return any(s.startswith(kw) for kw in FOOTER_STARTS)


def abrir_archivo(xls_path: Path) -> pd.DataFrame:
    """Lee el archivo XLS/XLSX. Intenta openpyxl, fallback a xlrd."""
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.close()
    tmp_path = Path(tmp.name)
    shutil.copy(xls_path, tmp_path)

    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df = pd.read_excel(tmp_path, sheet_name=0, header=None, engine="openpyxl")
        return df
    except Exception:
        pass
    finally:
        try:
            tmp_path.unlink()
        except Exception:
            pass

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df = pd.read_excel(xls_path, sheet_name=0, header=None, engine="xlrd")
    return df


# ---------------------------------------------------------------------------
# Procesamiento de un archivo
# ---------------------------------------------------------------------------

def procesar_archivo(xls_path: Path) -> pd.DataFrame | None:
    periodo = periodo_desde_path(xls_path)

    try:
        df = abrir_archivo(xls_path)
    except Exception as e:
        log.error("✗ No se pudo abrir '%s': %s", xls_path.name, e)
        return None

    # ------------------------------------------------------------------
    # 1. Localizar la fila de cabecera y la columna "Créditos de Consumo"
    #    La cabecera está en la fila donde col 0 contiene "Empresas"
    # ------------------------------------------------------------------
    idx_cabecera = None
    col_consumo  = None

    for i in range(min(8, len(df))):
        v0 = df.iloc[i, 0]
        if pd.notna(v0) and str(v0).strip().lower() == "empresas":
            idx_cabecera = i
            # Buscar la columna "Créditos de Consumo"
            for j in range(1, df.shape[1]):
                v = df.iloc[i, j]
                if pd.notna(v) and COL_OBJETIVO in str(v).strip().lower():
                    col_consumo = j
                    break
            break

    if idx_cabecera is None or col_consumo is None:
        log.warning("  No se encontró cabecera en '%s'", xls_path.name)
        return None

    # ------------------------------------------------------------------
    # 2. Leer filas de datos (desde idx_cabecera + 2, saltando fila vacía)
    # ------------------------------------------------------------------
    registros = []

    for i in range(idx_cabecera + 2, len(df)):
        v0 = df.iloc[i, 0]

        # Pie de página → parar
        if es_footer(v0):
            break

        # Fila vacía → parar (fin de datos)
        if pd.isna(v0) or not str(v0).strip():
            break

        nombre_orig = str(v0).strip()

        # Excluir totales
        if es_banco_excluido(nombre_orig):
            continue

        banco = limpiar_banco(nombre_orig)

        # Leer valor de Créditos de Consumo
        v_tc = df.iloc[i, col_consumo]
        if pd.isna(v_tc):
            nro_tc = 0
        else:
            try:
                nro_tc = int(float(v_tc))
            except (ValueError, TypeError):
                nro_tc = 0

        registros.append({
            "Periodo": periodo,
            "Banco":   banco,
            "Nro_TC":  nro_tc,
        })

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
            log.info("    → %d bancos", len(df))
        else:
            log.warning("    → Sin datos")

    if not frames:
        log.error("No se extrajo ningún dato.")
        return

    df_final = pd.concat(frames, ignore_index=True)
    df_final.sort_values(["Periodo", "Banco"], inplace=True)
    df_final.reset_index(drop=True, inplace=True)

    nombre_csv = "Numero_Tarjetas_Credito_Consumo.csv"
    ruta_csv   = OUTPUT_DIR / nombre_csv
    df_final.to_csv(ruta_csv, index=False, encoding="utf-8-sig")

    log.info("═" * 60)
    log.info("✓ Exportado: %s", ruta_csv.name)
    log.info("  Filas: %d  |  Columnas: %s", len(df_final), df_final.columns.tolist())
    log.info("  Periodos: %d  |  Bancos: %d", df_final["Periodo"].nunique(), df_final["Banco"].nunique())
    log.info("  Bancos únicos:")
    for b in sorted(df_final["Banco"].unique()):
        log.info("    %s", b)


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    main()
