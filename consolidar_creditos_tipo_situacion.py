"""
consolidar_creditos_tipo_situacion.py
======================================
Consolida los archivos Excel de "Créditos Directos según Tipo de Crédito
y Situación" de la SBS en un único CSV con las columnas:

    Periodo | Tipo | Sub_Tipo | Banco | Monto_miles

Solo se extraen los datos de la columna "Consumo":
  - Tipo "Revolventes"    → cols 21 (Vigentes), 22 (Refinanc. y Reestruct.), 23 (Atrasados)
  - Tipo "No Revolventes" → cols 24 (Vigentes), 25 (Refinanc. y Reestruct.), 26 (Atrasados)

Estructura del Excel (una sola hoja, XLS binario):
  - Fila 1   : título
  - Fila 2   : fecha del periodo
  - Fila 3   : "(En miles de soles)"
  - Fila 5   : tipos de crédito (col 0="Empresas", col 21="Consumo", ...)
  - Fila 6   : sub-tipos de Consumo (col 21="Revolventes", col 24="No Revolventes")
  - Fila 7   : situaciones (Vigentes / Refinanc. y Reestruct. / Atrasados)
  - Fila 9+  : datos por banco
  - Pie      : "TOTAL BANCA MÚLTIPLE" (excluir), "Fuente:", etc.

Nota: las columnas de Consumo son fijas (21-26) en todos los archivos inspeccionados.
Si en algún archivo cambian, el script las detecta dinámicamente buscando "Consumo"
en la fila de tipos.
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

CARPETA = "Créditos Directos según Tipo de Crédito y Situación"

BANCOS_EXCLUIR_PREFIJOS = ["total banca"]
FOOTER_STARTS = ["fuente", "nota", "(*)", "* la", "** mediante", "aprobado"]

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
    """Intenta openpyxl primero, fallback a xlrd para XLS binario."""
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
# Detección dinámica de columnas de Consumo
# ---------------------------------------------------------------------------

def detectar_columnas_consumo(df: pd.DataFrame) -> dict:
    """
    Busca la columna donde aparece "Consumo" en la fila de tipos (fila 5, idx=5),
    luego lee la fila de sub-tipos (fila 6, idx=6) para identificar
    "Revolventes" y "No Revolventes", y la fila de situaciones (fila 7, idx=7)
    para mapear Vigentes / Refinanc. / Atrasados.

    Devuelve un dict con estructura:
    {
        "Revolventes":    {"Vigentes": col, "Refinanc. y Reestruct.": col, "Atrasados": col},
        "No Revolventes": {"Vigentes": col, "Refinanc. y Reestruct.": col, "Atrasados": col},
    }
    """
    # Fila 5: encontrar col de inicio de "Consumo"
    col_consumo_inicio = None
    for j in range(df.shape[1]):
        v = df.iloc[5, j]
        if pd.notna(v) and "consumo" in str(v).strip().lower():
            col_consumo_inicio = j
            break

    if col_consumo_inicio is None:
        return {}

    # Fila 5: encontrar col de inicio del siguiente tipo (para saber el rango)
    col_siguiente_tipo = df.shape[1]
    for j in range(col_consumo_inicio + 1, df.shape[1]):
        v = df.iloc[5, j]
        if pd.notna(v) and str(v).strip() and str(v).strip().lower() != "consumo":
            col_siguiente_tipo = j
            break

    # Fila 6: dentro del rango de Consumo, buscar "Revolventes" y "No Revolventes"
    sub_tipos = {}  # {"Revolventes": col_inicio, "No Revolventes": col_inicio}
    for j in range(col_consumo_inicio, col_siguiente_tipo):
        v = df.iloc[6, j]
        if pd.notna(v) and str(v).strip():
            sub_tipos[str(v).strip()] = j

    if not sub_tipos:
        return {}

    # Fila 7: para cada sub-tipo, mapear las 3 situaciones
    resultado = {}
    sub_tipo_cols = sorted(sub_tipos.values())

    for idx_st, (nombre_st, col_st) in enumerate(sub_tipos.items()):
        # El sub-tipo ocupa desde col_st hasta el inicio del siguiente sub-tipo
        if idx_st + 1 < len(sub_tipo_cols):
            col_fin = sub_tipo_cols[idx_st + 1]
        else:
            col_fin = col_siguiente_tipo

        situaciones = {}
        for j in range(col_st, col_fin):
            v = df.iloc[7, j]
            if pd.notna(v) and str(v).strip():
                situaciones[str(v).strip()] = j

        resultado[nombre_st] = situaciones

    return resultado


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

    # Detectar columnas de Consumo dinámicamente
    mapa_consumo = detectar_columnas_consumo(df)
    if not mapa_consumo:
        log.warning("  No se encontró columna 'Consumo' en '%s'", xls_path.name)
        return None

    log.debug("  Mapa Consumo: %s", mapa_consumo)

    # Localizar inicio de datos: primera fila con banco (después de fila 7)
    idx_datos = 8
    while idx_datos < len(df):
        v = df.iloc[idx_datos, 0]
        if pd.notna(v) and str(v).strip() and not str(v).strip().lower().startswith("empresa"):
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

        # Iterar sobre Tipo (Revolventes / No Revolventes)
        for tipo, situaciones in mapa_consumo.items():
            # Iterar sobre Sub_Tipo (Vigentes / Refinanc. y Reestruct. / Atrasados)
            for sub_tipo, col_idx in situaciones.items():
                if col_idx >= df.shape[1]:
                    continue

                v = df.iloc[i, col_idx]
                if pd.isna(v) or str(v).strip() in ("-", ""):
                    monto = 0.0
                else:
                    try:
                        monto = float(v)
                    except (ValueError, TypeError):
                        monto = 0.0

                registros.append({
                    "Periodo":    periodo,
                    "Tipo":       tipo,
                    "Sub_Tipo":   sub_tipo,
                    "Banco":      banco,
                    "Monto_miles": monto,
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

    archivos = sorted(
        list(carpeta.rglob("*.xls")) + list(carpeta.rglob("*.xlsx")) + list(carpeta.rglob("*.XLSx"))
    )
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
    df_final.sort_values(["Periodo", "Tipo", "Sub_Tipo", "Banco"], inplace=True)
    df_final.reset_index(drop=True, inplace=True)

    ruta_csv = OUTPUT_DIR / "Creditos_Directos_Consumo_Tipo_Situacion.csv"
    df_final.to_csv(ruta_csv, index=False, encoding="utf-8-sig")

    log.info("═" * 60)
    log.info("✓ Exportado: %s", ruta_csv.name)
    log.info("  Filas: %d  |  Columnas: %s", len(df_final), df_final.columns.tolist())
    log.info("  Periodos: %d  |  Bancos: %d", df_final["Periodo"].nunique(), df_final["Banco"].nunique())
    log.info("  Tipos únicos: %s", sorted(df_final["Tipo"].unique()))
    log.info("  Sub_Tipos únicos: %s", sorted(df_final["Sub_Tipo"].unique()))
    log.info("  Bancos únicos:")
    for b in sorted(df_final["Banco"].unique()):
        log.info("    %s", b)


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    main()
