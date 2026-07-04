"""
actualizar_sbs.py
=================
Script maestro que:
  1. Copia automáticamente los archivos nuevos desde descargas_sbs/ → Data_SBS/
  2. Detecta qué periodos son nuevos (no están aún en el CSV de salida)
  3. Procesa SOLO los archivos nuevos y los agrega al CSV existente
  4. Genera un resumen de lo que se actualizó

Uso:
    python actualizar_sbs.py

No requiere argumentos. Detecta automáticamente los periodos faltantes.

Carpetas esperadas (junto a este script):
    descargas_sbs/   ← archivos descargados de la SBS (fuente)
    Data_SBS/        ← copia de trabajo (se gestiona automáticamente)
    Output_SBS/      ← CSVs de salida (se actualizan incrementalmente)
"""

import re
import shutil
import tempfile
import logging
import warnings
import sys
from pathlib import Path

import pandas as pd
import openpyxl

# ===========================================================================
# CONFIGURACIÓN GLOBAL
# ===========================================================================

BASE_SCRIPT = Path(__file__).parent
DESCARGAS   = BASE_SCRIPT / "descargas_sbs" / "Información de la Banca Múltiple"
DATA_SBS    = BASE_SCRIPT / "Data_SBS"
OUTPUT_SBS  = BASE_SCRIPT / "Output_SBS"
OUTPUT_SBS.mkdir(exist_ok=True)
DATA_SBS.mkdir(exist_ok=True)

# Extensiones válidas de archivos Excel
EXTENSIONES = {".xls", ".xlsx", ".xlsm", ".xlsx"}

# Meses en español → número de mes
MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "setiembre": 9, "septiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}

# Mapa: nombre de carpeta en descargas_sbs → nombre de carpeta en Data_SBS
# (por si hay diferencias de tildes u ortografía)
CARPETAS_FUENTE = {
    "Balance General y Estado de Ganancias y Pérdidas":
        "Balance General y Estado de Ganancias y Pérdidas",
    "Ranking de Créditos Directos por Tipo":
        "Ranking de Créditos Directos por Tipo",
    "Ranking de Créditos Directos por Modalidad de Operación":
        "Ranking de Créditos Directos por Modalidad de Operación",
    "Morosidad por tipo de crédito y modalidad":
        "Morosidad por tipo de crédito y modalidad",
    "Indicadores Financieros":
        "Indicadores Financieros",
    "Número de Tarjetas de Crédito por Tipo":
        "Número de Tarjetas de Crédito por Tipo",
    "Número de Tarjetas de Débito":
        "Número de Tarjetas de Débito",
    "Créditos Directos según Tipo de Crédito y Situación":
        "Créditos Directos según Tipo de Crédito y Situación",
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ===========================================================================
# PASO 1: SINCRONIZAR descargas_sbs → Data_SBS
# ===========================================================================

def sincronizar_archivos() -> dict[str, list[Path]]:
    """
    Copia todos los archivos Excel de descargas_sbs a Data_SBS
    que aún no existan en destino.
    Devuelve {carpeta_data: [archivos_nuevos_copiados]}.
    """
    nuevos: dict[str, list[Path]] = {}

    for nombre_fuente, nombre_destino in CARPETAS_FUENTE.items():
        carpeta_src = DESCARGAS / nombre_fuente
        carpeta_dst = DATA_SBS  / nombre_destino

        if not carpeta_src.exists():
            log.warning("  Carpeta fuente no encontrada: %s", carpeta_src.name)
            continue

        carpeta_dst.mkdir(exist_ok=True)
        nuevos[nombre_destino] = []

        # Buscar todos los Excel en la fuente (recursivo)
        for src_file in sorted(carpeta_src.rglob("*")):
            if src_file.suffix.lower() not in EXTENSIONES:
                continue

            # Ruta relativa dentro de la carpeta de categoría
            rel = src_file.relative_to(carpeta_src)

            # Destino siempre con extensión .xls para uniformidad
            dst_file = carpeta_dst / rel.parent / (rel.stem + ".xls")
            dst_file.parent.mkdir(parents=True, exist_ok=True)

            if not dst_file.exists():
                shutil.copy2(src_file, dst_file)
                nuevos[nombre_destino].append(dst_file)
                log.info("  + Copiado: %s", dst_file.relative_to(DATA_SBS))

    return nuevos


# ===========================================================================
# PASO 2: DETECTAR PERIODOS NUEVOS POR CSV
# ===========================================================================

def periodo_desde_path(xls_path: Path) -> str:
    """Construye el periodo YYYYMM desde la ruta del archivo."""
    year  = xls_path.parent.name
    month = xls_path.stem.lower()
    mes_num = MESES.get(month)
    if mes_num and year.isdigit():
        return f"{year}{mes_num:02d}"
    return f"{year}{xls_path.stem}"


def periodos_en_csv(ruta_csv: Path) -> set[str]:
    """Lee los periodos ya presentes en un CSV de salida."""
    if not ruta_csv.exists():
        return set()
    try:
        df = pd.read_csv(ruta_csv, encoding="utf-8-sig", usecols=["Periodo"])
        return set(df["Periodo"].astype(str).unique())
    except Exception:
        return set()


def archivos_nuevos(carpeta_data: Path, ruta_csv: Path) -> list[Path]:
    """
    Devuelve los archivos de carpeta_data cuyo periodo no está en ruta_csv.
    """
    existentes = periodos_en_csv(ruta_csv)
    resultado = []
    for f in sorted(
        list(carpeta_data.rglob("*.xls")) +
        list(carpeta_data.rglob("*.xlsx"))
    ):
        p = periodo_desde_path(f)
        if p not in existentes:
            resultado.append(f)
    return resultado


# ===========================================================================
# UTILIDADES COMUNES (compartidas por todos los procesadores)
# ===========================================================================

def abrir_excel(xls_path: Path) -> pd.DataFrame:
    """Lee un Excel con openpyxl (fallback xlrd). Devuelve DataFrame crudo."""
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
        try: tmp_path.unlink()
        except Exception: pass
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return pd.read_excel(xls_path, sheet_name=0, header=None, engine="xlrd")


def limpiar_banco_base(nombre: str, sufijos: list[str] | None = None) -> str:
    """Normalización base de nombre de banco."""
    s = re.sub(r"\s+", " ", nombre).strip()
    s = re.sub(r"\*+\s*$", "", s).strip()
    s = s.rstrip(".").strip()
    for sufijo in (sufijos or [" (con sucursales en el exterior)"]):
        if s.lower().endswith(sufijo.lower()):
            s = s[: -len(sufijo)].strip()
    s = re.sub(r"^B\.([A-ZÁÉÍÓÚÑ])", r"B. \1", s)
    return s


def es_excluido(nombre: str, prefijos: list[str]) -> bool:
    norm = re.sub(r"\s+", " ", nombre).strip().lower()
    return any(norm.startswith(p) for p in prefijos)


def es_footer_gen(valor, keywords: list[str]) -> bool:
    if pd.isna(valor): return False
    s = str(valor).strip().lower()
    return any(s.startswith(k) for k in keywords)


def guardar_incremental(df_nuevo: pd.DataFrame, ruta_csv: Path,
                        col_sort: list[str]) -> int:
    """
    Agrega df_nuevo al CSV existente (si existe), evitando duplicados de periodo.
    Devuelve el número de filas nuevas agregadas.
    """
    if ruta_csv.exists():
        df_existente = pd.read_csv(ruta_csv, encoding="utf-8-sig")
        # Eliminar periodos que ya vamos a reescribir (por si se re-procesa)
        periodos_nuevos = set(df_nuevo["Periodo"].astype(str).unique())
        df_existente = df_existente[
            ~df_existente["Periodo"].astype(str).isin(periodos_nuevos)
        ]
        df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
    else:
        df_final = df_nuevo.copy()

    df_final.sort_values(col_sort, inplace=True, na_position="first")
    df_final.reset_index(drop=True, inplace=True)
    df_final.to_csv(ruta_csv, index=False, encoding="utf-8-sig")
    return len(df_nuevo)


# ===========================================================================
# PROCESADORES ESPECÍFICOS POR CATEGORÍA
# ===========================================================================

# ---------------------------------------------------------------------------
# 1. Balance General y Estado de Ganancias y Pérdidas
# ---------------------------------------------------------------------------

def _procesar_balance_archivo(xls_path: Path) -> pd.DataFrame | None:
    """Procesador del Balance General (con filtro bold y tablas apiladas)."""
    from consolidar_sbs import procesar_archivo as _proc_bal
    categoria = "Balance General y Estado de Ganancias y Pérdidas"
    return _proc_bal(xls_path, categoria)


def actualizar_balance(archivos: list[Path]) -> int:
    ruta_csv = OUTPUT_SBS / "Balance_General_y_Estado_de_Ganancias_y_Perdidas.csv"
    frames = []
    for f in archivos:
        log.info("    %s", f.relative_to(DATA_SBS))
        df = _procesar_balance_archivo(f)
        if df is not None:
            frames.append(df)
            log.info("      → %d registros", len(df))
    if not frames:
        return 0
    df_nuevo = pd.concat(frames, ignore_index=True)
    return guardar_incremental(
        df_nuevo, ruta_csv,
        ["Periodo", "Hoja", "Tabla", "Banco", "Concepto_Cuenta"]
    )


# ---------------------------------------------------------------------------
# 2. Ranking de Créditos Directos por Tipo
# ---------------------------------------------------------------------------

def _procesar_ranking_tipo_archivo(xls_path: Path) -> pd.DataFrame | None:
    from consolidar_ranking_creditos import procesar_archivo as _proc
    return _proc(xls_path)


def actualizar_ranking_tipo(archivos: list[Path]) -> int:
    ruta_csv = OUTPUT_SBS / "Ranking_Creditos_Directos_por_Tipo.csv"
    frames = []
    for f in archivos:
        log.info("    %s", f.relative_to(DATA_SBS))
        df = _procesar_ranking_tipo_archivo(f)
        if df is not None:
            frames.append(df)
            log.info("      → %d registros", len(df))
    if not frames:
        return 0
    df_nuevo = pd.concat(frames, ignore_index=True)
    return guardar_incremental(
        df_nuevo, ruta_csv,
        ["Periodo", "Categoria", "Ranking"]
    )


# ---------------------------------------------------------------------------
# 3. Ranking de Créditos Directos por Modalidad
# ---------------------------------------------------------------------------

def _procesar_ranking_modalidad_archivo(xls_path: Path) -> pd.DataFrame | None:
    from consolidar_ranking_modalidad import procesar_archivo as _proc
    return _proc(xls_path)


def actualizar_ranking_modalidad(archivos: list[Path]) -> int:
    ruta_csv = OUTPUT_SBS / "Ranking_Creditos_Directos_por_Modalidad.csv"
    frames = []
    for f in archivos:
        log.info("    %s", f.relative_to(DATA_SBS))
        df = _procesar_ranking_modalidad_archivo(f)
        if df is not None:
            frames.append(df)
            log.info("      → %d registros", len(df))
    if not frames:
        return 0
    df_nuevo = pd.concat(frames, ignore_index=True)
    return guardar_incremental(
        df_nuevo, ruta_csv,
        ["Periodo", "Categoria", "Ranking"]
    )


# ---------------------------------------------------------------------------
# 4. Morosidad por tipo de crédito y modalidad
# ---------------------------------------------------------------------------

def _procesar_morosidad_archivo(xls_path: Path) -> pd.DataFrame | None:
    from consolidar_morosidad import procesar_archivo as _proc
    return _proc(xls_path)


def actualizar_morosidad(archivos: list[Path]) -> int:
    ruta_csv = OUTPUT_SBS / "Morosidad_por_Tipo_y_Modalidad.csv"
    frames = []
    for f in archivos:
        log.info("    %s", f.relative_to(DATA_SBS))
        df = _procesar_morosidad_archivo(f)
        if df is not None:
            frames.append(df)
            log.info("      → %d registros", len(df))
    if not frames:
        return 0
    df_nuevo = pd.concat(frames, ignore_index=True)
    return guardar_incremental(
        df_nuevo, ruta_csv,
        ["Periodo", "Concepto", "Sub_Concepto", "Banco"]
    )


# ---------------------------------------------------------------------------
# 5. Indicadores Financieros
# ---------------------------------------------------------------------------

def _procesar_indicadores_archivo(xls_path: Path) -> pd.DataFrame | None:
    from consolidar_indicadores import procesar_archivo as _proc
    return _proc(xls_path)


def actualizar_indicadores(archivos: list[Path]) -> int:
    ruta_csv = OUTPUT_SBS / "Indicadores_Financieros.csv"
    frames = []
    for f in archivos:
        log.info("    %s", f.relative_to(DATA_SBS))
        df = _procesar_indicadores_archivo(f)
        if df is not None:
            frames.append(df)
            log.info("      → %d registros", len(df))
    if not frames:
        return 0
    df_nuevo = pd.concat(frames, ignore_index=True)
    return guardar_incremental(
        df_nuevo, ruta_csv,
        ["Periodo", "Indicador", "Sub_Indicador", "Banco"]
    )


# ---------------------------------------------------------------------------
# 6. Número de Tarjetas de Crédito por Tipo
# ---------------------------------------------------------------------------

def _procesar_tc_archivo(xls_path: Path) -> pd.DataFrame | None:
    from consolidar_tarjetas_credito import procesar_archivo as _proc
    return _proc(xls_path)


def actualizar_tarjetas_credito(archivos: list[Path]) -> int:
    ruta_csv = OUTPUT_SBS / "Numero_Tarjetas_Credito_Consumo.csv"
    frames = []
    for f in archivos:
        log.info("    %s", f.relative_to(DATA_SBS))
        df = _procesar_tc_archivo(f)
        if df is not None:
            frames.append(df)
            log.info("      → %d registros", len(df))
    if not frames:
        return 0
    df_nuevo = pd.concat(frames, ignore_index=True)
    return guardar_incremental(df_nuevo, ruta_csv, ["Periodo", "Banco"])


# ---------------------------------------------------------------------------
# 7. Número de Tarjetas de Débito
# ---------------------------------------------------------------------------

def _procesar_td_archivo(xls_path: Path) -> pd.DataFrame | None:
    from consolidar_tarjetas_debito import procesar_archivo as _proc
    return _proc(xls_path)


def actualizar_tarjetas_debito(archivos: list[Path]) -> int:
    ruta_csv = OUTPUT_SBS / "Numero_Tarjetas_Debito.csv"
    frames = []
    for f in archivos:
        log.info("    %s", f.relative_to(DATA_SBS))
        df = _procesar_td_archivo(f)
        if df is not None:
            frames.append(df)
            log.info("      → %d registros", len(df))
    if not frames:
        return 0
    df_nuevo = pd.concat(frames, ignore_index=True)
    return guardar_incremental(df_nuevo, ruta_csv, ["Periodo", "Banco"])


# ---------------------------------------------------------------------------
# 8. Créditos Directos según Tipo de Crédito y Situación
# ---------------------------------------------------------------------------

def _procesar_creditos_tipo_archivo(xls_path: Path) -> pd.DataFrame | None:
    from consolidar_creditos_tipo_situacion import procesar_archivo as _proc
    return _proc(xls_path)


def actualizar_creditos_tipo(archivos: list[Path]) -> int:
    ruta_csv = OUTPUT_SBS / "Creditos_Directos_Consumo_Tipo_Situacion.csv"
    frames = []
    for f in archivos:
        log.info("    %s", f.relative_to(DATA_SBS))
        df = _procesar_creditos_tipo_archivo(f)
        if df is not None:
            frames.append(df)
            log.info("      → %d registros", len(df))
    if not frames:
        return 0
    df_nuevo = pd.concat(frames, ignore_index=True)
    return guardar_incremental(
        df_nuevo, ruta_csv,
        ["Periodo", "Tipo", "Sub_Tipo", "Banco"]
    )


# ===========================================================================
# MAPA: carpeta Data_SBS → (función actualizadora, CSV de salida)
# ===========================================================================

PROCESADORES = {
    "Balance General y Estado de Ganancias y Pérdidas": (
        actualizar_balance,
        "Balance_General_y_Estado_de_Ganancias_y_Perdidas.csv",
    ),
    "Ranking de Créditos Directos por Tipo": (
        actualizar_ranking_tipo,
        "Ranking_Creditos_Directos_por_Tipo.csv",
    ),
    "Ranking de Créditos Directos por Modalidad de Operación": (
        actualizar_ranking_modalidad,
        "Ranking_Creditos_Directos_por_Modalidad.csv",
    ),
    "Morosidad por tipo de crédito y modalidad": (
        actualizar_morosidad,
        "Morosidad_por_Tipo_y_Modalidad.csv",
    ),
    "Indicadores Financieros": (
        actualizar_indicadores,
        "Indicadores_Financieros.csv",
    ),
    "Número de Tarjetas de Crédito por Tipo": (
        actualizar_tarjetas_credito,
        "Numero_Tarjetas_Credito_Consumo.csv",
    ),
    "Número de Tarjetas de Débito": (
        actualizar_tarjetas_debito,
        "Numero_Tarjetas_Debito.csv",
    ),
    "Créditos Directos según Tipo de Crédito y Situación": (
        actualizar_creditos_tipo,
        "Creditos_Directos_Consumo_Tipo_Situacion.csv",
    ),
}


# ===========================================================================
# MAIN
# ===========================================================================

def main():
    log.info("=" * 65)
    log.info("  ACTUALIZADOR SBS - Banca Múltiple")
    log.info("=" * 65)

    # ------------------------------------------------------------------
    # PASO 1: Sincronizar archivos nuevos de descargas_sbs → Data_SBS
    # ------------------------------------------------------------------
    log.info("")
    log.info("PASO 1: Sincronizando archivos desde descargas_sbs...")
    log.info("─" * 65)

    if not DESCARGAS.exists():
        log.error("No se encontró la carpeta: %s", DESCARGAS)
        log.error("Asegúrate de que 'descargas_sbs' esté junto a este script.")
        sys.exit(1)

    nuevos_copiados = sincronizar_archivos()
    total_copiados = sum(len(v) for v in nuevos_copiados.values())
    if total_copiados == 0:
        log.info("  ✓ No hay archivos nuevos que copiar.")
    else:
        log.info("  ✓ %d archivo(s) nuevo(s) copiado(s) a Data_SBS.", total_copiados)

    # ------------------------------------------------------------------
    # PASO 2: Detectar y procesar periodos nuevos por categoría
    # ------------------------------------------------------------------
    log.info("")
    log.info("PASO 2: Detectando y procesando periodos nuevos...")
    log.info("─" * 65)

    resumen = []

    for nombre_carpeta, (fn_actualizar, nombre_csv) in PROCESADORES.items():
        carpeta_data = DATA_SBS / nombre_carpeta
        ruta_csv     = OUTPUT_SBS / nombre_csv

        if not carpeta_data.exists():
            log.warning("  [OMITIDA] Carpeta no encontrada: %s", nombre_carpeta)
            resumen.append((nombre_carpeta, nombre_csv, 0, "carpeta no encontrada"))
            continue

        # Detectar archivos con periodos nuevos
        nuevos = archivos_nuevos(carpeta_data, ruta_csv)

        if not nuevos:
            periodos_csv = periodos_en_csv(ruta_csv)
            log.info("  ✓ %-55s → ya actualizado (%d periodos)",
                     nombre_carpeta[:55], len(periodos_csv))
            resumen.append((nombre_carpeta, nombre_csv, 0, "ya actualizado"))
            continue

        periodos_nuevos_str = sorted({periodo_desde_path(f) for f in nuevos})
        log.info("")
        log.info("  ► %s", nombre_carpeta)
        log.info("    Periodos nuevos: %s", periodos_nuevos_str)

        try:
            filas_agregadas = fn_actualizar(nuevos)
            log.info("    ✓ %d filas agregadas al CSV.", filas_agregadas)
            resumen.append((nombre_carpeta, nombre_csv, filas_agregadas,
                            f"+{len(periodos_nuevos_str)} periodos"))
        except Exception as e:
            log.error("    ✗ Error: %s", e)
            resumen.append((nombre_carpeta, nombre_csv, 0, f"ERROR: {e}"))

    # ------------------------------------------------------------------
    # RESUMEN FINAL
    # ------------------------------------------------------------------
    log.info("")
    log.info("=" * 65)
    log.info("  RESUMEN")
    log.info("=" * 65)
    log.info("  %-50s  %s", "CSV", "Estado")
    log.info("  " + "-" * 62)
    for _, csv_name, filas, estado in resumen:
        marca = "✓" if filas > 0 or "actualizado" in estado else ("✗" if "ERROR" in estado else "─")
        log.info("  %s %-48s  %s", marca, csv_name[:48], estado)
    log.info("=" * 65)
    log.info("  Archivos CSV en: %s", OUTPUT_SBS)
    log.info("=" * 65)


if __name__ == "__main__":
    main()
