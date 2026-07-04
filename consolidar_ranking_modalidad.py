"""
consolidar_ranking_modalidad.py
================================
Consolida los archivos Excel de "Ranking de Créditos Directos por
Modalidad de Operación" de la SBS en un único CSV con las columnas:

    Periodo | Categoria | Ranking | Banco | Monto | Participacion

Estructura del Excel (una sola hoja, 6 columnas):
  - Fila 1   : título del reporte
  - Fila 2   : fecha del periodo
  - Fila 3   : unidad (miles de soles)
  - Fila 4   : nombre de categoría (ej. "Descuentos")  ← col 0, col 3 vacía
  - Fila 6   : cabeceras (Empresas / Monto / Participación / Acumulado)
  - Fila 7   : segunda línea de cabeceras (% / ACUMULADO)
  - Fila 8   : vacía
  - Fila 9+  : datos → col0=Ranking, col2=Banco, col3=Monto, col4=Participacion%
  - Pie      : "Fuente:", "No incluye..."

Categorías apiladas verticalmente (4 por archivo):
  Descuentos | Tarjetas de Crédito | Préstamos | Arrendamiento Financiero
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

CARPETA = "Ranking de Créditos Directos por Modalidad de Operación"

# Columnas dentro de cada bloque (un solo bloque por fila)
COL_RANKING      = 0
COL_BANCO        = 2
COL_MONTO        = 3
COL_PARTICIPACION = 4

# Palabras clave de pie de página
FOOTER_STARTS = [
    "fuente",
    "no incluye",
    "las definiciones",
    "aprobado",
    "(*)",
    "nota",
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

def periodo_desde_path(xls_path: Path) -> str:
    year  = xls_path.parent.name
    month = xls_path.stem.lower()
    mes_num = MESES.get(month)
    if mes_num and year.isdigit():
        return f"{year}{mes_num:02d}"
    return f"{year}{xls_path.stem}"


def limpiar_texto(texto: str) -> str:
    """Quita asteriscos, puntos finales sueltos y normaliza espacios."""
    s = re.sub(r"\s+", " ", texto).strip()
    s = s.rstrip("*").strip()
    s = s.rstrip(".").strip()
    return s


def limpiar_banco(nombre: str) -> str:
    """
    Normaliza nombre de banco:
    - Quita asteriscos y puntos finales
    - Corrige 'B.Nombre' → 'B. Nombre'
    """
    s = limpiar_texto(nombre)
    s = re.sub(r"^B\.([A-ZÁÉÍÓÚÑ])", r"B. \1", s)
    return s


def es_footer(valor) -> bool:
    if pd.isna(valor):
        return False
    s = str(valor).strip().lower()
    return any(s.startswith(kw) for kw in FOOTER_STARTS)


def es_numero_entero_positivo(valor) -> bool:
    if pd.isna(valor):
        return False
    try:
        n = float(valor)
        return n == int(n) and n > 0
    except (ValueError, TypeError):
        return False


def abrir_archivo(xls_path: Path) -> pd.DataFrame:
    """
    Lee el archivo XLS (que en realidad es XLSX) con openpyxl.
    Fallback a xlrd si el archivo es XLS binario real.
    """
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
        try:
            tmp_path.unlink()
        except Exception:
            pass
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df = pd.read_excel(xls_path, sheet_name=0, header=None, engine="xlrd")
        return df
    finally:
        try:
            tmp_path.unlink()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Detección de bloques verticales
# ---------------------------------------------------------------------------

def detectar_bloques(df: pd.DataFrame) -> list[tuple[int, int]]:
    """
    Detecta los pares (idx_titulo, idx_datos_inicio) de cada bloque.

    Un título de categoría es una fila donde:
      - col 0 tiene texto (no numérico, no fecha, no encabezado/pie)
      - col 3 está vacía (distingue de filas de datos que tienen Monto)

    Los datos empiezan en la primera fila con número entero en col 0,
    buscando desde idx_titulo + 2.
    """
    SKIP = ["ranking", "(en", "fuente", "no incl", "las def", "aprobado",
            "empresas", "monto", "%", "acum", "(*)", "nota"]

    bloques = []
    for i in range(len(df)):
        v0 = df.iloc[i, 0]
        v3 = df.iloc[i, 3] if df.shape[1] > 3 else None

        if pd.isna(v0) or not str(v0).strip():
            continue
        if not (pd.isna(v3) or not str(v3).strip()):
            continue  # col 3 tiene valor → fila de datos

        s = str(v0).strip()

        # Descartar números
        try:
            float(s)
            continue
        except (ValueError, TypeError):
            pass

        # Descartar fechas (datetime objects)
        if hasattr(v0, 'year'):
            continue

        # Descartar encabezados y pie de página
        if any(s.lower().startswith(p) for p in SKIP):
            continue

        # Buscar inicio de datos (primera fila con ranking numérico)
        idx_datos = None
        for k in range(i + 2, min(i + 6, len(df))):
            if es_numero_entero_positivo(df.iloc[k, 0]):
                idx_datos = k
                break

        if idx_datos is not None:
            bloques.append((i, idx_datos))

    return bloques


# ---------------------------------------------------------------------------
# Parseo de un bloque
# ---------------------------------------------------------------------------

def parsear_bloque(
    df: pd.DataFrame,
    idx_titulo: int,
    idx_datos: int,
    idx_fin: int,
    periodo: str,
) -> list[dict]:
    """Extrae los registros de un bloque vertical."""
    categoria = limpiar_texto(str(df.iloc[idx_titulo, 0]))
    registros = []

    for i in range(idx_datos, idx_fin):
        fila = df.iloc[i]

        if es_footer(fila.iloc[0]):
            break

        v_rank = fila.iloc[COL_RANKING]
        v_banco = fila.iloc[COL_BANCO]      if COL_BANCO        < len(fila) else None
        v_monto = fila.iloc[COL_MONTO]      if COL_MONTO        < len(fila) else None
        v_part  = fila.iloc[COL_PARTICIPACION] if COL_PARTICIPACION < len(fila) else None

        if not es_numero_entero_positivo(v_rank):
            continue
        if pd.isna(v_banco) or not str(v_banco).strip():
            continue

        banco = limpiar_banco(str(v_banco).strip())

        # Monto
        if pd.isna(v_monto) or str(v_monto).strip() in ("-", ""):
            monto = None
        else:
            try:
                monto = float(v_monto)
            except (ValueError, TypeError):
                monto = None

        # Participación
        if pd.isna(v_part) or str(v_part).strip() in ("-", ""):
            participacion = None
        else:
            try:
                participacion = float(v_part)
            except (ValueError, TypeError):
                participacion = None

        registros.append({
            "Periodo":       periodo,
            "Categoria":     categoria,
            "Ranking":       int(float(v_rank)),
            "Banco":         banco,
            "Monto":         monto,
            "Participacion": participacion,
        })

    return registros


# ---------------------------------------------------------------------------
# Procesamiento de un archivo
# ---------------------------------------------------------------------------

def procesar_archivo(xls_path: Path) -> pd.DataFrame | None:
    periodo = periodo_desde_path(xls_path)
    todos = []

    try:
        df = abrir_archivo(xls_path)
        bloques = detectar_bloques(df)

        if not bloques:
            log.warning("  Sin bloques detectados en '%s'", xls_path.name)
            return None

        for k, (idx_titulo, idx_datos) in enumerate(bloques):
            idx_fin = bloques[k + 1][0] if k + 1 < len(bloques) else len(df)
            registros = parsear_bloque(df, idx_titulo, idx_datos, idx_fin, periodo)
            todos.extend(registros)

    except Exception as e:
        log.error("✗ Error procesando '%s': %s", xls_path.name, e)
        return None

    return pd.DataFrame(todos) if todos else None


# ---------------------------------------------------------------------------
# Función principal
# ---------------------------------------------------------------------------

def main():
    carpeta = BASE_DIR / CARPETA
    if not carpeta.exists():
        log.error(
            "No se encontró '%s' dentro de '%s'.\n"
            "Asegúrate de que la carpeta exista con ese nombre exacto.",
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
    df_final.sort_values(["Periodo", "Categoria", "Ranking"], inplace=True)
    df_final.reset_index(drop=True, inplace=True)

    nombre_csv = "Ranking_Creditos_Directos_por_Modalidad.csv"
    ruta_csv   = OUTPUT_DIR / nombre_csv
    df_final.to_csv(ruta_csv, index=False, encoding="utf-8-sig")

    log.info("═" * 60)
    log.info("✓ Exportado: %s", ruta_csv.name)
    log.info("  Filas: %d  |  Columnas: %s", len(df_final), df_final.columns.tolist())
    log.info("  Periodos: %d  |  Categorías: %d  |  Bancos: %d",
             df_final["Periodo"].nunique(),
             df_final["Categoria"].nunique(),
             df_final["Banco"].nunique())
    log.info("  Categorías únicas:")
    for cat in sorted(df_final["Categoria"].unique()):
        log.info("    - %s", cat)


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    main()
