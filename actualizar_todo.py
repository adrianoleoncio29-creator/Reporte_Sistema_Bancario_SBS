"""
actualizar_todo.py
==================
Script maestro END-TO-END del Monitor del Sistema Financiero Peruano.

Ejecuta el pipeline completo en 4 fases:

  FASE 1 — Descarga (sbs_downloader.py)
    Abre Chromium, navega al portal de la SBS y descarga los archivos XLS
    nuevos de las subcarpetas configuradas → descargas_sbs/

  FASE 2 — Sincronización incremental (actualizar_sbs.py)
    Copia los archivos nuevos de descargas_sbs/ → Data_SBS/ y agrega
    solo los periodos nuevos a cada CSV de Output_SBS/ (no reprocesa lo
    que ya existe).

  FASE 3 — Consolidación completa (consolidar_todo.py)
    Re-genera todos los CSVs desde cero leyendo Data_SBS/, aplicando
    los parsers especializados de cada categoría.

  FASE 4 — Dashboard (generar_dashboard_sfp.py)
    Construye el JSON con todos los bancos y periodos, y escribe
    dashboard_sistema_financiero.html.

Uso:
    python actualizar_todo.py                   # pipeline completo
    python actualizar_todo.py --desde-fase 2    # saltar la descarga
    python actualizar_todo.py --desde-fase 3    # saltar descarga + sync
    python actualizar_todo.py --desde-fase 4    # solo dashboard
    python actualizar_todo.py --solo-descarga   # solo fase 1

Notas:
  - La Fase 1 (Playwright) abre un navegador visible.
    Si quieres ejecución silenciosa cambia headless=False → True
    en sbs_downloader.py antes de ejecutar.
  - El downloader detecta automáticamente qué archivos ya existen
    en descargas_sbs/ y solo descarga los que faltan. No hay que
    configurar ningún filtro de fecha manualmente.
  - Si alguna fase falla, el script detiene el pipeline y muestra
    el error. Las fases anteriores ya completadas no se revierten.
"""

import sys
import time
import asyncio
import logging
import runpy
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

BASE = Path(__file__).parent

# ---------------------------------------------------------------------------
# Helpers de ejecución
# ---------------------------------------------------------------------------

def _run_module(nombre: str, ruta_py: Path) -> None:
    """Ejecuta un script Python como __main__ usando runpy."""
    log.info("─" * 62)
    log.info("▶  %s", nombre)
    log.info("─" * 62)
    t0 = time.time()
    runpy.run_path(str(ruta_py), run_name="__main__")
    log.info("✓  %s completado en %.1f s", nombre, time.time() - t0)


def _run_async_module(nombre: str, ruta_py: Path) -> None:
    """
    Ejecuta un script async (Playwright) como __main__.
    runpy lo carga; si define main() como coroutine, la ejecutamos con asyncio.
    """
    log.info("─" * 62)
    log.info("▶  %s", nombre)
    log.info("─" * 62)
    t0 = time.time()

    import importlib.util, types

    spec = importlib.util.spec_from_file_location("__sbs_downloader__", ruta_py)
    mod  = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)

    main_fn = getattr(mod, "main", None)
    if main_fn is not None and asyncio.iscoroutinefunction(main_fn):
        asyncio.run(main_fn())
    else:
        # Fallback: ejecutar como script plano
        runpy.run_path(str(ruta_py), run_name="__main__")

    log.info("✓  %s completado en %.1f s", nombre, time.time() - t0)


# ---------------------------------------------------------------------------
# Definición del pipeline
# ---------------------------------------------------------------------------

FASES = [
    {
        "num":    1,
        "nombre": "Descarga SBS (Playwright)",
        "tipo":   "async",
        "script": BASE / "sbs_downloader.py",
    },
    {
        "num":    2,
        "nombre": "Sincronización incremental de CSVs",
        "tipo":   "sync",
        "script": BASE / "actualizar_sbs.py",
    },
    {
        "num":    3,
        "nombre": "Consolidación completa",
        "tipo":   "sync",
        "script": BASE / "consolidar_todo.py",
        "args":   ["--solo-consolidar"],   # consolidar_todo sin regenerar dashboard
    },
    {
        "num":    4,
        "nombre": "Generación del Dashboard HTML",
        "tipo":   "sync",
        "script": BASE / "generar_dashboard_sfp.py",
    },
]

# ---------------------------------------------------------------------------
# Parse de argumentos
# ---------------------------------------------------------------------------

def parsear_args() -> tuple[int, int]:
    """
    Devuelve (desde_fase, hasta_fase).
    Por defecto: desde=1, hasta=4 (pipeline completo).
    """
    args = sys.argv[1:]
    desde = 1
    hasta = 4

    if "--solo-descarga" in args:
        desde, hasta = 1, 1
    if "--desde-fase" in args:
        idx = args.index("--desde-fase")
        if idx + 1 < len(args):
            try:
                desde = int(args[idx + 1])
            except ValueError:
                pass

    return desde, hasta

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    desde, hasta = parsear_args()

    log.info("=" * 62)
    log.info("  PIPELINE COMPLETO — Monitor Sistema Financiero Peruano")
    log.info("=" * 62)

    fases_a_ejecutar = [f for f in FASES if desde <= f["num"] <= hasta]

    if not fases_a_ejecutar:
        log.error("No hay fases que ejecutar con --desde-fase %d", desde)
        sys.exit(1)

    log.info("  Fases a ejecutar: %s",
             ", ".join(f"[{f['num']}] {f['nombre']}" for f in fases_a_ejecutar))
    log.info("")

    resultados: dict[str, str] = {}
    t_total = time.time()

    for fase in fases_a_ejecutar:
        num    = fase["num"]
        nombre = fase["nombre"]
        script = fase["script"]
        tipo   = fase["tipo"]
        args_extra = fase.get("args", [])

        # Inyectar args extra en sys.argv para scripts que los leen
        argv_original = sys.argv[:]
        if args_extra:
            sys.argv = [str(script)] + args_extra

        try:
            if tipo == "async":
                _run_async_module(nombre, script)
            else:
                _run_module(nombre, script)
            resultados[f"Fase {num}: {nombre}"] = "✓ OK"
        except Exception as e:
            resultados[f"Fase {num}: {nombre}"] = f"✗ ERROR — {e}"
            log.error("")
            log.error("✗  La fase %d falló: %s", num, e)
            log.error("   Pipeline detenido. Revisa el error antes de continuar.")
            sys.argv = argv_original
            break
        finally:
            sys.argv = argv_original

        log.info("")

    # ── Resumen ──────────────────────────────────────────────────────────
    elapsed = time.time() - t_total
    log.info("=" * 62)
    log.info("  RESUMEN  (%.1f s total)", elapsed)
    log.info("=" * 62)
    for paso, estado in resultados.items():
        log.info("  %s  %s", estado, paso)

    fases_pendientes = [f for f in FASES
                        if f["num"] > hasta or
                           f["nombre"] not in " ".join(resultados.keys())]

    if any("ERROR" in v for v in resultados.values()):
        log.info("=" * 62)
        sys.exit(1)
    else:
        log.info("")
        log.info("  ✅  Pipeline completado sin errores.")
        output = BASE / "dashboard_sistema_financiero.html"
        if output.exists():
            log.info("  📊  Dashboard: %s  (%.0f KB)",
                     output.name, output.stat().st_size / 1024)
        log.info("=" * 62)


if __name__ == "__main__":
    main()
