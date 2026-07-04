"""
consolidar_ranking_creditos.py
==============================
Consolida los archivos Excel de "Ranking de Créditos Directos por Tipo"
de la SBS en un único CSV con las columnas:

    Periodo | Categoria | Ranking | Banco | Participacion

Estructura del Excel (una sola hoja):
  - Fila 1   : título del reporte
  - Fila 2   : fecha del periodo
  - Fila 3   : unidad (miles de soles)
  - Fila 5   : nombres de categoría — 3 en paralelo (cols 0, 7, 16)
               + 3 bloques verticales (filas ~5, ~29, ~53) → 9 categorías
  - Fila 7-8 : cabeceras (Empresas / Monto / Participación / Acumulado)
  - Fila 10+ : datos (Ranking, vacía, Banco, Monto, Participación%, Acumulado%)
  - Pie      : "Fuente:", "Las definiciones...", "No incluye..."

Bloques en paralelo (columnas):
  Bloque A: cols 0-5   → Ranking=col0, Banco=col2, Participacion=col4
  Bloque B: cols 7-12  → Ranking=col7, Banco=col9, Participacion=col11
  Bloque C: cols 14-19 → Ranking=col14, Banco=col16, Participacion=col18

Bloques verticales: se detectan buscando filas donde col0 tiene texto de
categoría (texto no numérico, col3 vacía) y la fila siguiente+2 tiene
cabeceras de columna.
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

# Nombre de la carpeta fuente dentro de Data_SBS
CARPETA_RANKING = "Ranking de Créditos Directos por Tipo"

# Palabras clave de pie de página
FOOTER_STARTS = [
    "fuente",
    "las definiciones",
    "aprobado",
    "no incluye",
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

# Offsets de columna dentro de cada bloque paralelo
# (ranking_offset, banco_offset, monto_offset, participacion_offset)
BLOQUES_COLS = [
    (0, 2, 3, 4),    # Bloque A
    (7, 9, 10, 11),  # Bloque B
    (14, 16, 17, 18),# Bloque C
]

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
    """Construye el periodo YYYYMM desde la ruta del archivo."""
    year  = xls_path.parent.name
    month = xls_path.stem.lower()
    mes_num = MESES.get(month)
    if mes_num and year.isdigit():
        return f"{year}{mes_num:02d}"
    return f"{year}{xls_path.stem}"


def limpiar_banco(nombre: str) -> str:
    """
    Normaliza el nombre de un banco en los archivos de Ranking:
    - Quita asteriscos finales
    - Quita puntos finales sueltos (ej. 'B. Santander Perú.' → 'B. Santander Perú')
    - Normaliza espacios múltiples y saltos de línea
    - Corrige 'B.Falabella' → 'B. Falabella' (espacio faltante)
    """
    s = re.sub(r"\s+", " ", nombre).strip()
    s = s.rstrip("*").strip()
    s = s.rstrip(".").strip()
    # Corregir 'B.Nombre' → 'B. Nombre' (punto sin espacio)
    s = re.sub(r"^B\.([A-ZÁÉÍÓÚÑ])", r"B. \1", s)
    return s


def limpiar_categoria(texto: str) -> str:
    """
    Normaliza el nombre de categoría:
    - Quita asteriscos finales
    - Normaliza espacios y saltos de línea
    """
    s = re.sub(r"\s+", " ", texto).strip()
    s = s.rstrip("*").strip()
    return s


def es_footer(valor) -> bool:
    """True si la celda marca el inicio del pie de página."""
    if pd.isna(valor):
        return False
    s = str(valor).strip().lower()
    return any(s.startswith(kw) for kw in FOOTER_STARTS)


def es_numero_entero(valor) -> bool:
    """True si el valor es un entero positivo (número de ranking)."""
    if pd.isna(valor):
        return False
    try:
        n = float(valor)
        return n == int(n) and n > 0
    except (ValueError, TypeError):
        return False


def abrir_xls_como_xlsx(xls_path: Path) -> tuple[pd.DataFrame, Path]:
    """
    Copia el .xls a un temporal .xlsx y lo lee con openpyxl.
    Si falla (archivo XLS binario real), intenta con xlrd directamente.
    Devuelve (DataFrame crudo, ruta_temporal o None).
    """
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.close()
    tmp_path = Path(tmp.name)
    shutil.copy(xls_path, tmp_path)

    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df = pd.read_excel(tmp_path, sheet_name=0, header=None, engine="openpyxl")
        return df, tmp_path
    except Exception:
        # Fallback: intentar leer el original con xlrd (XLS binario real)
        try:
            tmp_path.unlink()
        except Exception:
            pass
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df = pd.read_excel(xls_path, sheet_name=0, header=None, engine="xlrd")
        return df, None


# ---------------------------------------------------------------------------
# Detección de bloques verticales
# ---------------------------------------------------------------------------

def detectar_bloques_verticales(df: pd.DataFrame) -> list[tuple[int, int]]:
    """
    Detecta los índices de fila donde comienza cada bloque vertical.
    Un bloque comienza cuando:
      - col 0 tiene texto de categoría (no numérico, no pie de página, no encabezado)
      - col 3 está vacía (distingue de filas de datos que tienen Monto en col 3)

    Devuelve lista de (idx_fila_categoria, idx_fila_datos_inicio).
    La fila de datos empieza 4 filas después del título
    (título → cabecera1 → cabecera2 → vacía → datos).
    """
    bloques = []
    SKIP_PREFIXES = [
        "ranking", "(en", "fuente", "las def", "aprobado", "no incl", "(*)", "nota"
    ]

    for i in range(len(df)):
        v0 = df.iloc[i, 0]
        v3 = df.iloc[i, 3] if df.shape[1] > 3 else None

        if pd.isna(v0) or not str(v0).strip():
            continue
        if not (pd.isna(v3) or not str(v3).strip()):
            continue  # col 3 tiene valor → fila de datos

        s = str(v0).strip()

        # Descartar si es número
        try:
            float(s)
            continue
        except (ValueError, TypeError):
            pass

        # Descartar encabezados y pie de página
        s_lower = s.lower()
        if any(s_lower.startswith(p) for p in SKIP_PREFIXES):
            continue

        # Descartar fechas
        if hasattr(v0, 'year'):  # datetime
            continue

        # Descartar si es muy corto o parece cabecera de columna
        if s in ("Empresas", "Monto", "%", "ACUMULADO"):
            continue

        # Calcular inicio de datos: buscar la primera fila con número en col 0
        # a partir de i+2 (saltando cabeceras)
        idx_datos = None
        for k in range(i + 2, min(i + 6, len(df))):
            if es_numero_entero(df.iloc[k, 0]):
                idx_datos = k
                break

        if idx_datos is not None:
            bloques.append((i, idx_datos))
            log.debug("  Bloque vertical: fila_cat=%d ('%s'), datos desde=%d", i, s, idx_datos)

    return bloques


# ---------------------------------------------------------------------------
# Parseo de un bloque (vertical × horizontal)
# ---------------------------------------------------------------------------

def parsear_bloque(
    df: pd.DataFrame,
    idx_cat: int,
    idx_datos: int,
    idx_datos_fin: int,
    periodo: str,
) -> list[dict]:
    """
    Procesa un bloque vertical completo (3 categorías en paralelo).
    Devuelve lista de dicts con los campos del CSV.
    """
    registros = []

    # Leer los 3 nombres de categoría de la fila idx_cat
    categorias_bloque = []
    for col_rank, col_banco, col_monto, col_part in BLOQUES_COLS:
        if col_rank < df.shape[1]:
            v = df.iloc[idx_cat, col_rank]
            cat = limpiar_categoria(str(v)) if pd.notna(v) and str(v).strip() else None
            categorias_bloque.append(cat)
        else:
            categorias_bloque.append(None)

    log.debug("    Categorías del bloque: %s", categorias_bloque)

    # Iterar filas de datos
    for i in range(idx_datos, idx_datos_fin):
        fila = df.iloc[i]

        # Verificar si es pie de página
        if es_footer(fila.iloc[0]):
            break

        # Procesar cada bloque paralelo
        for (col_rank, col_banco, col_monto, col_part), categoria in zip(BLOQUES_COLS, categorias_bloque):
            if categoria is None:
                continue
            if col_rank >= df.shape[1]:
                continue

            v_rank  = fila.iloc[col_rank]  if col_rank  < len(fila) else None
            v_banco = fila.iloc[col_banco] if col_banco < len(fila) else None
            v_monto = fila.iloc[col_monto] if col_monto < len(fila) else None
            v_part  = fila.iloc[col_part]  if col_part  < len(fila) else None

            # El ranking debe ser un número entero positivo
            if not es_numero_entero(v_rank):
                continue

            # El banco debe tener texto
            if pd.isna(v_banco) or not str(v_banco).strip():
                continue

            banco = limpiar_banco(str(v_banco).strip())

            # Monto: puede ser 0 o '-' para bancos sin créditos de ese tipo
            if pd.isna(v_monto) or str(v_monto).strip() in ("-", ""):
                monto = None
            else:
                try:
                    monto = float(v_monto)
                except (ValueError, TypeError):
                    monto = None

            # Participación: puede ser '-' (banco sin créditos de ese tipo)
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
# Procesamiento de un archivo XLS
# ---------------------------------------------------------------------------

def procesar_archivo(xls_path: Path) -> pd.DataFrame | None:
    """
    Procesa un archivo XLS de Ranking y devuelve un DataFrame con
    Periodo | Categoria | Ranking | Banco | Participacion.
    """
    periodo  = periodo_desde_path(xls_path)
    tmp_path = None
    todos_registros = []

    try:
        df, tmp_path = abrir_xls_como_xlsx(xls_path)

        # Detectar bloques verticales
        bloques = detectar_bloques_verticales(df)

        if not bloques:
            log.warning("  Sin bloques detectados en '%s'", xls_path.name)
            return None

        # Procesar cada bloque vertical
        for k, (idx_cat, idx_datos) in enumerate(bloques):
            # El bloque termina donde empieza el siguiente (o al final del df)
            if k + 1 < len(bloques):
                idx_fin = bloques[k + 1][0]
            else:
                idx_fin = len(df)

            registros = parsear_bloque(df, idx_cat, idx_datos, idx_fin, periodo)
            todos_registros.extend(registros)
            log.debug("  Bloque %d: %d registros", k, len(registros))

    except Exception as e:
        log.error("✗ Error procesando '%s': %s", xls_path, e)
        return None

    finally:
        if tmp_path and tmp_path.exists():
            try:
                tmp_path.unlink()
            except Exception:
                pass

    if not todos_registros:
        return None

    return pd.DataFrame(todos_registros)


# ---------------------------------------------------------------------------
# Función principal
# ---------------------------------------------------------------------------

def main():
    carpeta = BASE_DIR / CARPETA_RANKING
    if not carpeta.exists():
        # Intentar con nombre sin tilde (por si Data_SBS usa nombre simplificado)
        carpeta_alt = BASE_DIR / "Ranking de Creditos Directos por Tipo"
        if carpeta_alt.exists():
            carpeta = carpeta_alt
        else:
            log.error(
                "No se encontró la carpeta '%s' dentro de '%s'.\n"
                "Asegúrate de que la carpeta exista con ese nombre exacto.",
                CARPETA_RANKING, BASE_DIR,
            )
            return

    archivos = sorted(carpeta.rglob("*.xls"))
    if not archivos:
        log.warning("No se encontraron archivos .xls en '%s'", carpeta)
        return

    log.info("Procesando %d archivos de: %s", len(archivos), carpeta)
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

    # Ordenar
    df_final.sort_values(
        ["Periodo", "Categoria", "Ranking"],
        inplace=True,
    )
    df_final.reset_index(drop=True, inplace=True)

    # Exportar
    nombre_csv = "Ranking_Creditos_Directos_por_Tipo.csv"
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
