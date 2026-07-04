"""
sbs_inspector2.py
=================
Diagnóstico profundo del panel de resultados de una subcarpeta SBS.
Vuelca el HTML del iframe/panel de contenido para entender cómo están
estructurados los periodos y el botón de descarga Excel.

Uso:
    python sbs_inspector2.py
"""

import asyncio
from playwright.async_api import async_playwright, TimeoutError as PWTimeoutError

PORTAL_URL = (
    "https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaBoletinEstadistico.aspx?p=1#"
)
# URL directa de la subcarpeta (obtenida del inspector anterior)
SUBCARPETA_URL = (
    "https://www.sbs.gob.pe/app/stats_net/stats/"
    "EstadisticaSistemaFinancieroResultados.aspx?c=B-2311"
)


async def main():
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=False, slow_mo=300)
        context = await browser.new_context(
            accept_downloads=True,
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1280, "height": 900},
            locale="es-PE",
        )
        page = await context.new_page()

        # ── Navega directamente a la subcarpeta ───────────────────────────
        print(f"🌐 Navegando directamente a: {SUBCARPETA_URL}")
        await page.goto(SUBCARPETA_URL, wait_until="networkidle", timeout=60_000)
        await asyncio.sleep(3)

        # ── Vuelca el HTML completo de la página ──────────────────────────
        print("\n── HTML completo (primeros 8000 chars) ──")
        html = await page.content()
        print(html[:8000])

        # ── Lista todos los <select> ───────────────────────────────────────
        print("\n── SELECTs encontrados ──")
        selects = await page.locator("select").all()
        print(f"Total selects: {len(selects)}")
        for i, sel in enumerate(selects):
            sel_id = await sel.get_attribute("id") or "(sin id)"
            sel_name = await sel.get_attribute("name") or "(sin name)"
            opciones = await sel.locator("option").all()
            print(f"\n  Select [{i}] id='{sel_id}' name='{sel_name}' — {len(opciones)} opciones:")
            for op in opciones[:10]:  # Muestra las primeras 10
                txt = (await op.inner_text()).strip()
                val = await op.get_attribute("value")
                print(f"    value='{val}' → '{txt}'")
            if len(opciones) > 10:
                print(f"    ... y {len(opciones) - 10} más")

        # ── Lista todos los enlaces ────────────────────────────────────────
        print("\n── ENLACES encontrados (primeros 30) ──")
        enlaces = await page.locator("a").all()
        print(f"Total enlaces: {len(enlaces)}")
        for i, a in enumerate(enlaces[:30]):
            txt = (await a.inner_text()).strip()
            href = await a.get_attribute("href") or ""
            onclick = await a.get_attribute("onclick") or ""
            print(f"  [{i}] '{txt}' href='{href[:80]}' onclick='{onclick[:80]}'")

        # ── Lista todos los botones e inputs ──────────────────────────────
        print("\n── BOTONES/INPUTS encontrados ──")
        botones = await page.locator("input[type='button'], input[type='submit'], button").all()
        print(f"Total botones: {len(botones)}")
        for i, b in enumerate(botones):
            val = await b.get_attribute("value") or await b.inner_text()
            bid = await b.get_attribute("id") or "(sin id)"
            print(f"  [{i}] id='{bid}' value/text='{val}'")

        # ── Guarda screenshot ─────────────────────────────────────────────
        await page.screenshot(path="diagnostico_subcarpeta.png", full_page=True)
        print("\n📸 Screenshot: diagnostico_subcarpeta.png")

        input("\n⏸  Presiona ENTER para cerrar...")
        await context.close()
        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
