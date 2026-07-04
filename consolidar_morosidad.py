"""
consolidar_morosidad.py
========================
Consolida los archivos Excel de "Morosidad por tipo de crédito y modalidad"
de la SBS en un único CSV con las columnas:

    Periodo | Concepto | Sub_Concepto | Banco | Porcentaje

Estructura del Excel (una sola hoja, XLS binario):
  - Fila 1   : título
  - Fila 2   : fecha del periodo
  - Fila 3   : "(En porcentaje)"
  - Fila 5   : cabecera → col 0 = "Concepto", col 1+ = nombres de banco
  - Fila 6+  : datos en dos niveles:
                 Concepto    → fila que empieza con "Crédito" o "Total"
                 Sub_Concepto → filas siguientes hasta la próxima fila vacía
  - Pie      : "Nota:", "*", etc.

Bancos excluidos (sucursales en el exterior):
  - "B. de Crédito del Perú (con sucursales en el exterior)"
  - "Interbank (con sucursales en el exterior)"

Los valores ya están en porcentaje en el Excel.
'-' significa que el banco no opera en esa modalidad → Porcentaje = NaN.
"""

import re
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

CARPETA = "Morosidad por tipo de crédito y modalidad"

# Bancos a excluir completamente
BANCOS_EXCLUIR = {
    "total banca múltiple",
    "total banca multiple",
}

# Sufijos a quitar del nombre del banco para normalizar
SUFIJOS_QUITAR = [
    " (con sucursales en el exterior)",
]

# Prefijos que identifican una fila de Concepto (nivel 1)
CONCEPTO_PREFIJOS = (
    "crédito",
    "credito",
    "total créditos",
    "total creditos",
)

# Palabras clave de pie de página
FOOTER_STARTS = ["nota", "* ", "(*)", "la información", "aprobado"]

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


def limpiar_texto(texto: str) -> str:
    """Normaliza espacios, quita asteriscos y notas al pie (ej. '1/')."""
    s = re.sub(r"\s+", " ", str(texto)).strip()
    s = s.rstrip("*").strip()
    # Quitar notas tipo "1/", "2/" al final
    s = re.sub(r"\s+\d+/$", "", s).strip()
    return s


def limpiar_banco(nombre: str) -> str:
    """Normaliza nombre de banco: quita asteriscos, puntos finales,
    sufijos de sucursales y corrige 'B.Nombre' → 'B. Nombre'."""
    s = re.sub(r"\s+", " ", nombre).strip()
    s = s.rstrip("*").strip()
    s = s.rstrip(".").strip()
    # Quitar sufijos de sucursales en el exterior (case-insensitive)
    for sufijo in SUFIJOS_QUITAR:
        if s.lower().endswith(sufijo.lower()):
            s = s[: -len(sufijo)].strip()
    # Corregir 'B.Nombre' → 'B. Nombre'
    s = re.sub(r"^B\.([A-ZÁÉÍÓÚÑ])", r"B. \1", s)
    return s


def es_banco_excluido(nombre: str) -> bool:
    # Normalizar: minúsculas, quitar tildes para comparación robusta
    norm = re.sub(r"\s+", " ", nombre).strip().lower()
    norm_sin_tilde = (norm
        .replace("á","a").replace("é","e").replace("í","i")
        .replace("ó","o").replace("ú","u").replace("ñ","n"))
    return norm in BANCOS_EXCLUIR or norm_sin_tilde in BANCOS_EXCLUIR

def es_concepto(texto: str) -> bool:
    """True si la fila es un Concepto de nivel 1."""
    s = limpiar_texto(texto).lower()
    return any(s.startswith(p) for p in CONCEPTO_PREFIJOS)


def es_footer(valor) -> bool:
    if pd.isna(valor):
        return False
    s = str(valor).strip().lower()
    return any(s.startswith(kw) for kw in FOOTER_STARTS)


def abrir_archivo(xls_path: Path) -> pd.DataFrame:
    """Lee el archivo XLS. Intenta openpyxl primero, fallback a xlrd."""
    import shutil, tempfile
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

    # Fallback: XLS binario real
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
    # 1. Leer cabecera de bancos (fila 5, índice 5)
    # ------------------------------------------------------------------
    fila_bancos = df.iloc[5]
    bancos = {}  # {col_idx: nombre_banco_limpio}
    for j in range(1, df.shape[1]):
        v = fila_bancos.iloc[j]
        if pd.notna(v) and str(v).strip():
            nombre_original = str(v).strip()
            if not es_banco_excluido(nombre_original):
                bancos[j] = limpiar_banco(nombre_original)

    if not bancos:
        log.warning("  Sin bancos detectados en '%s'", xls_path.name)
        return None

    # ------------------------------------------------------------------
    # 2. Iterar filas de datos (desde fila 6, índice 6)
    # ------------------------------------------------------------------
    registros = []
    concepto_actual = None

    for i in range(6, len(df)):
        v0 = df.iloc[i, 0]

        # Pie de página → parar
        if es_footer(v0):
            break

        # Fila vacía → separador entre bloques, no hace nada
        if pd.isna(v0) or not str(v0).strip():
            continue

        texto = str(v0).strip()

        # Determinar si es Concepto (nivel 1) o Sub_Concepto (nivel 2)
        if es_concepto(texto):
            concepto_actual = limpiar_texto(texto)
            sub_concepto    = None   # el propio concepto no tiene sub
        else:
            sub_concepto = limpiar_texto(texto)

        if concepto_actual is None:
            continue  # aún no hemos encontrado el primer concepto

        # Leer valores por banco
        for col_idx, banco in bancos.items():
            if col_idx >= df.shape[1]:
                continue

            v = df.iloc[i, col_idx]

            if pd.isna(v) or str(v).strip() in ("-", ""):
                porcentaje = None
            else:
                try:
                    porcentaje = float(v)
                except (ValueError, TypeError):
                    porcentaje = None

            registros.append({
                "Periodo":      periodo,
                "Concepto":     concepto_actual,
                "Sub_Concepto": sub_concepto,   # None si es la fila del concepto mismo
                "Banco":        banco,
                "Porcentaje":   porcentaje,
            })

    if not registros:
        return None

    df_out = pd.DataFrame(registros)
    # Eliminar filas donde Porcentaje es NaN (banco sin actividad en esa modalidad)
    df_out = df_out.dropna(subset=["Porcentaje"])
    return df_out if not df_out.empty else None


# ---------------------------------------------------------------------------
# Función principal
# ---------------------------------------------------------------------------

def main():
    carpeta = BASE_DIR / CARPETA
    if not carpeta.exists():
        log.error(
            "No se encontró '%s' dentro de '%s'.",
            CARPETA, BASE_DIR,
        )
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
        ["Periodo", "Concepto", "Sub_Concepto", "Banco"],
        inplace=True,
        na_position="first",
    )
    df_final.reset_index(drop=True, inplace=True)

    nombre_csv = "Morosidad_por_Tipo_y_Modalidad.csv"
    ruta_csv   = OUTPUT_DIR / nombre_csv
    df_final.to_csv(ruta_csv, index=False, encoding="utf-8-sig")

    log.info("═" * 60)
    log.info("✓ Exportado: %s", ruta_csv.name)
    log.info("  Filas: %d  |  Columnas: %s", len(df_final), df_final.columns.tolist())
    log.info("  Periodos: %d  |  Conceptos: %d  |  Bancos: %d",
             df_final["Periodo"].nunique(),
             df_final["Concepto"].nunique(),
             df_final["Banco"].nunique())
    log.info("  Conceptos únicos:")
    for c in sorted(df_final["Concepto"].unique()):
        log.info("    - %s", c)
    log.info("  Sub_Conceptos únicos:")
    for s in sorted(df_final["Sub_Concepto"].dropna().unique()):
        log.info("    - %s", s)


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    main()
