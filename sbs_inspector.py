"""
sbs_inspector.py
================
Script de diagnóstico: abre el portal SBS, expande el nodo raíz y vuelca
TODOS los textos del árbol de menú para que puedas copiar los nombres exactos
que necesitas poner en SUBCARPETAS_OBJETIVO dentro de sbs_downloader.py.

Uso:
    python sbs_inspector.py
"""

import asyncio
from playwright.async_api import async_playwright, TimeoutError as PWTimeoutError

PORTAL_URL = "https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaBoletinEstadistico.aspx?p=1#"
ROOT_NODE_TEXT = "Información de la Banca Múltiple"


async def main():
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=False, slow_mo=400)
        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1280, "height": 900},
            locale="es-PE",
        )
        page = await context.new_page()

        print(f"🌐 Navegando a: {PORTAL_URL}")
        await page.goto(PORTAL_URL, wait_until="networkidle", timeout=60_000)
        await asyncio.sleep(3)

        # ── Vuelca el HTML completo del árbol de menú para diagnóstico ────
        print("\n── HTML del contenedor del menú (primeros 3000 chars) ──")
        try:
            # Intenta varios selectores comunes para el contenedor del árbol
            for sel in ["#TreeView1", "#tvMenu", ".treeview", "[id*='Tree']", "[id*='tree']"]:
                try:
                    await page.wait_for_selector(sel, timeout=3_000)
                    html = await page.locator(sel).first.inner_html()
                    print(f"Selector encontrado: {sel}")
                    print(html[:3000])
                    break
                except Exception:
                    continue
            else:
                # Si no encontró ningún selector específico, vuelca el body completo
                print("No se encontró selector de árbol. Volcando body completo (primeros 5000 chars):")
                body = await page.locator("body").inner_html()
                print(body[:5000])
        except Exception as e:
            print(f"Error al obtener HTML: {e}")

        # ── Intenta expandir el nodo raíz y listar todos los textos ───────
        print(f"\n── Buscando nodo raíz: '{ROOT_NODE_TEXT}' ──")
        try:
            # Busca el nodo por texto parcial (más tolerante)
            nodos = await page.get_by_text("Banca Múltiple").all()
            print(f"Nodos con 'Banca Múltiple': {len(nodos)}")
            for i, n in enumerate(nodos):
                try:
                    txt = await n.inner_text()
                    tag = await n.evaluate("el => el.tagName")
                    print(f"  [{i}] <{tag}> → repr: {repr(txt)}")
                except Exception:
                    pass

            # Hace clic en el primero que encuentre
            if nodos:
                await nodos[0].click()
                await page.wait_for_load_state("networkidle", timeout=20_000)
                await asyncio.sleep(2)
                print("  ✔ Clic realizado. Esperando expansión...")

        except Exception as e:
            print(f"Error expandiendo nodo raíz: {e}")

        # ── Lista TODOS los textos visibles en el árbol tras la expansión ──
        print("\n── Todos los textos del árbol de menú tras expansión ──")
        try:
            # Busca todos los <a> y <span> dentro del árbol
            todos = await page.locator("a, span").all()
            textos_visibles = set()
            for elem in todos:
                try:
                    if await elem.is_visible():
                        txt = (await elem.inner_text()).strip()
                        if txt and len(txt) > 3:
                            textos_visibles.add(repr(txt))
                except Exception:
                    continue

            print(f"Total de textos únicos visibles: {len(textos_visibles)}")
            for t in sorted(textos_visibles):
                print(f"  {t}")

        except Exception as e:
            print(f"Error listando textos: {e}")

        # ── Guarda screenshot para revisión visual ─────────────────────────
        await page.screenshot(path="diagnostico_menu.png", full_page=True)
        print("\n📸 Screenshot guardado: diagnostico_menu.png")

        input("\n⏸  Presiona ENTER para cerrar el navegador...")
        await context.close()
        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
