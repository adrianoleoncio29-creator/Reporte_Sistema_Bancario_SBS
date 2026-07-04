"""
consolidar_tarjetas_debito.py
==============================
Consolida los archivos Excel de "Número de Tarjetas de Débito"
de la SBS en un único CSV con las columnas:

    Periodo | Banco | Nro_TD

Estructura del Excel (una sola hoja, 2 columnas):
  - Fila 1   : título
  - Fila 2   : fecha del periodo
  - Fila 4   : cabecera → col 0 = "Empresas", col 1 = "N° Tarjetas de débito"
  - Fila 7+  : datos → col 0 = Banco, col 1 = cantidad
  - Pie      : "TOTAL BANCA MÚLTIPLE" (excluir), "Fuente:", etc.
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

CARPETA = "Número de Tarjetas de Débito"

BANCOS_EXCLUIR_PREFIJOS = ["total banca"]
FOOTER_STARTS = ["fuente", "nota", "aprobado", "* mediante", "la resolución"]

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
# Utilidades (idénticas al script de Tarjetas de Crédito)
# ---------------------------------------------------------------------------

def periodo_desde_path(xls_path: Path) -> str:
    year  = xls_path.parent.name
    month = xls_path.stem.lower()
    mes_num = MESES.get(month)
    if mes_num and year.isdigit():
        return f"{year}{mes_num:02d}"
    return f"{year}{xls_path.stem}"


def limpiar_banco(nombre: str) -> str:
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

    # Localizar fila de cabecera: donde col 0 = "Empresas"
    idx_cabecera = None
    for i in range(min(8, len(df))):
        v0 = df.iloc[i, 0]
        if pd.notna(v0) and str(v0).strip().lower() == "empresas":
            idx_cabecera = i
            break

    if idx_cabecera is None:
        log.warning("  No se encontró cabecera en '%s'", xls_path.name)
        return None

    # Datos desde idx_cabecera + 2 (saltando fila vacía separadora, fila 5 en 0-based)
    # En la inspección vimos que los datos empiezan en fila 7 (idx=6), cabecera en idx=4
    # → idx_cabecera + 2 puede ser vacío, buscamos la primera fila con datos
    idx_datos = idx_cabecera + 1
    while idx_datos < len(df):
        v = df.iloc[idx_datos, 0]
        if pd.notna(v) and str(v).strip() and str(v).strip().lower() != "empresas":
            break
        idx_datos += 1

    registros = []
    for i in range(idx_datos, len(df)):
        v0 = df.iloc[i, 0]

        if es_footer(v0):
            break
        if pd.isna(v0) or not str(v0).strip():
            break

        nombre_orig = str(v0).strip()
        if es_banco_excluido(nombre_orig):
            continue

        banco = limpiar_banco(nombre_orig)

        v_td = df.iloc[i, 1] if df.shape[1] > 1 else None
        if pd.isna(v_td):
            nro_td = 0
        else:
            try:
                nro_td = int(float(v_td))
            except (ValueError, TypeError):
                nro_td = 0

        registros.append({
            "Periodo": periodo,
            "Banco":   banco,
            "Nro_TD":  nro_td,
        })

    return pd.DataFrame(registros) if registros else None


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

    ruta_csv = OUTPUT_DIR / "Numero_Tarjetas_Debito.csv"
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
