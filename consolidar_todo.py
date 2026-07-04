"""
consolidar_todo.py
==================
Script maestro que ejecuta TODOS los consolidadores de la SBS en secuencia
y genera el dashboard actualizado al final.

Orden de ejecución:
  1. consolidar_sbs.py             → Balance General y EGP
  2. consolidar_indicadores.py     → Indicadores Financieros
  3. consolidar_morosidad.py       → Morosidad por Tipo y Modalidad
  4. consolidar_ranking_creditos.py → Ranking Créditos Directos por Tipo
  5. consolidar_ranking_modalidad.py → Ranking Créditos por Modalidad
  6. consolidar_tarjetas_credito.py → Tarjetas de Crédito
  7. consolidar_tarjetas_debito.py  → Tarjetas de Débito
  8. consolidar_creditos_tipo_situacion.py → Créditos por Tipo y Situación
  9. generar_dashboard_sfp.py       → Dashboard HTML

Uso:
    python consolidar_todo.py

Flags opcionales:
    --solo-consolidar   Ejecuta solo los consolidadores, omite el dashboard
    --solo-dashboard    Ejecuta solo el dashboard (asume CSVs ya generados)
"""

import sys
import time
import logging
import importlib
from pathlib import Path

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Definición de pasos
# ---------------------------------------------------------------------------
# Cada entrada: (nombre_display, modulo_python, funcion_main)
PASOS_CONSOLIDACION = [
    ("Balance General y EGP",              "consolidar_sbs",                      "main"),
    ("Indicadores Financieros",            "consolidar_indicadores",              "main"),
    ("Morosidad por Tipo y Modalidad",     "consolidar_morosidad",                "main"),
    ("Ranking Créditos por Tipo",          "consolidar_ranking_creditos",         "main"),
    ("Ranking Créditos por Modalidad",     "consolidar_ranking_modalidad",        "main"),
    ("Tarjetas de Crédito",                "consolidar_tarjetas_credito",         "main"),
    ("Tarjetas de Débito",                 "consolidar_tarjetas_debito",          "main"),
    ("Créditos por Tipo y Situación",      "consolidar_creditos_tipo_situacion",  "main"),
]

PASO_DASHBOARD = ("Generar Dashboard HTML", "generar_dashboard_sfp", "main")

# ---------------------------------------------------------------------------
# Wrapper para módulos sin main()
# ---------------------------------------------------------------------------

def _ejecutar_dashboard():
    """
    generar_dashboard_sfp.py genera el HTML a nivel de módulo (no tiene main()).
    Lo ejecutamos con runpy para que se comporte igual a `py generar_dashboard_sfp.py`.
    """
    import runpy
    from pathlib import Path
    ruta = Path(__file__).parent / "generar_dashboard_sfp.py"
    runpy.run_path(str(ruta), run_name="__main__")

# ---------------------------------------------------------------------------
# Ejecutor
# ---------------------------------------------------------------------------

def ejecutar_paso(nombre: str, modulo: str, funcion: str) -> bool:
    """
    Importa el módulo y llama a su función main().
    Devuelve True si tuvo éxito, False si falló.
    """
    log.info("─" * 60)
    log.info("▶  %s", nombre)
    log.info("─" * 60)
    t0 = time.time()
    try:
        mod = importlib.import_module(modulo)
        # Recargar para que cambios en disco se reflejen si se corre más de una vez
        importlib.reload(mod)
        fn = getattr(mod, funcion)
        fn()
        elapsed = time.time() - t0
        log.info("✓  %s completado en %.1f s", nombre, elapsed)
        return True
    except Exception as e:
        elapsed = time.time() - t0
        log.error("✗  %s falló después de %.1f s: %s", nombre, elapsed, e)
        return False


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    args = sys.argv[1:]
    solo_consolidar = "--solo-consolidar" in args
    solo_dashboard  = "--solo-dashboard"  in args

    log.info("=" * 60)
    log.info("  CONSOLIDADOR MAESTRO SBS — Sistema Financiero Peruano")
    log.info("=" * 60)

    resultados = {}
    t_inicio = time.time()

    # ── Paso 1-8: Consolidadores ──────────────────────────────────────────
    if not solo_dashboard:
        log.info("")
        log.info("FASE 1: Consolidación de datos SBS")
        log.info("")
        for nombre, modulo, funcion in PASOS_CONSOLIDACION:
            ok = ejecutar_paso(nombre, modulo, funcion)
            resultados[nombre] = "✓ OK" if ok else "✗ ERROR"
            log.info("")

    # ── Paso 9: Dashboard ─────────────────────────────────────────────────
    if not solo_consolidar:
        log.info("FASE 2: Generación del Dashboard")
        log.info("")
        nombre = PASO_DASHBOARD[0]
        log.info("─" * 60)
        log.info("▶  %s", nombre)
        log.info("─" * 60)
        t0 = time.time()
        try:
            _ejecutar_dashboard()
            elapsed = time.time() - t0
            log.info("✓  %s completado en %.1f s", nombre, elapsed)
            resultados[nombre] = "✓ OK"
        except Exception as e:
            elapsed = time.time() - t0
            log.error("✗  %s falló después de %.1f s: %s", nombre, elapsed, e)
            resultados[nombre] = "✗ ERROR"

    # ── Resumen final ─────────────────────────────────────────────────────
    t_total = time.time() - t_inicio
    log.info("")
    log.info("=" * 60)
    log.info("  RESUMEN FINAL  (%.1f s total)", t_total)
    log.info("=" * 60)
    for paso, estado in resultados.items():
        log.info("  %s  %s", estado, paso)
    log.info("=" * 60)

    # Salir con código de error si algún paso falló
    errores = sum(1 for v in resultados.values() if "ERROR" in v)
    if errores:
        log.warning("  %d paso(s) con error.", errores)
        sys.exit(1)
    else:
        log.info("  Todo completado sin errores.")


if __name__ == "__main__":
    main()
