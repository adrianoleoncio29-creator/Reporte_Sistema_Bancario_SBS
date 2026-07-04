"""Mapea TODOS los nodos raíz del menú y sus subcarpetas con sus códigos."""
import asyncio
from playwright.async_api import async_playwright

PORTAL_URL = (
    "https://www.sbs.gob.pe/app/stats_net/stats/"
    "EstadisticaBoletinEstadistico.aspx?p=1#"
)

async def main():
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=False, slow_mo=150)
        page = await browser.new_page(locale="es-PE")
        await page.goto(PORTAL_URL, wait_until="networkidle", timeout=60000)
        await asyncio.sleep(3)

        # Obtiene todos los nodos raíz del menú (nivel 0)
        nodos_raiz = await page.evaluate("""() => {
            return Array.from(document.querySelectorAll('div.MEN01_nivel_0_OUT'))
                .map(d => ({
                    texto: d.querySelector('a')?.innerText?.trim() || d.innerText.trim().split('\\n')[0],
                    onclick: d.getAttribute('onclick') || ''
                }));
        }""")

        print(f"Nodos raíz encontrados: {len(nodos_raiz)}")
        for n in nodos_raiz:
            print(f"  {n}")

        # Expande TODOS los nodos raíz uno por uno y captura sus hijos
        print("\n── Expandiendo todos los nodos raíz ──")
        for i, nodo in enumerate(nodos_raiz):
            try:
                div = page.locator("div.MEN01_nivel_0_OUT").nth(i)
                await div.click(timeout=3000)
                await asyncio.sleep(1)
            except Exception as e:
                print(f"  Error expandiendo nodo {i}: {e}")

        await asyncio.sleep(1)

        # Ahora extrae TODOS los enlaces visibles del menú completo
        todos_links = await page.evaluate("""() => {
            return Array.from(document.querySelectorAll('#faqs a[href*="c="]'))
                .map(a => ({
                    texto: a.innerText.trim(),
                    href: a.getAttribute('href'),
                    visible: a.offsetParent !== null
                }));
        }""")

        print(f"\n── TODOS los enlaces del menú ({len(todos_links)} total) ──")
        for l in todos_links:
            codigo = l['href'].split('c=')[-1] if 'c=' in l['href'] else '?'
            vis = '✔' if l['visible'] else '✗'
            print(f"  [{vis}] {codigo:12} {l['texto']}")

        await browser.close()

asyncio.run(main())
