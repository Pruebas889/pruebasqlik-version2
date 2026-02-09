from playwright.sync_api import sync_playwright # pyright: ignore[reportMissingImports]
import time
import logging

LOG = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

URL = (
    "https://qlik.copservir.com/sense/app/d39c40fb-a304-4eaf-9a30-50b7279d33f1/"
    "sheet/4f191cdb-aa40-409d-86b2-497a427a8b6a/state/analysis"
)
USERNAME = "Qlikzona29"
PASSWORD = "pF2A3f2x*"


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=["--start-maximized"])
        context = browser.new_context(viewport={"width": 1920, "height": 1080})
        page = context.new_page()
        LOG.info('Opening %s', URL)
        page.goto(URL, wait_until="networkidle")
        # small wait for any dynamic focus behavior
        time.sleep(1)

        # Try typing into the active element and press Tab via Playwright keyboard
        try:
            LOG.info('Typing username via Playwright keyboard')
            page.keyboard.type(USERNAME, delay=80)  # delay in ms per character
            time.sleep(0.12)
            LOG.info('Pressing Tab via Playwright')
            page.keyboard.press('Tab')
            time.sleep(0.18)
            LOG.info('Typing password via Playwright keyboard')
            page.keyboard.type(PASSWORD, delay=80)
            time.sleep(0.5)
            LOG.info('Done (Playwright attempt)')
            # keep browser open briefly so you can observe
            time.sleep(2)
            browser.close()
            return
        except Exception as e:
            LOG.exception('Playwright keyboard attempt failed, will try fallback selectors')

        # Fallback: try to find password input and click/type
        try:
            pw = page.locator("input[type='password']")
            if pw.count() > 0:
                LOG.info('Found input[type=password] via Playwright locator; clicking and typing')
                pw.first.click()
                page.keyboard.type(PASSWORD, delay=80)
                time.sleep(0.5)
                browser.close()
                return
        except Exception:
            LOG.exception('Fallback locator attempt failed')

        LOG.info('All Playwright attempts finished; leaving browser open for inspection')
        try:
            time.sleep(10)
        finally:
            browser.close()


if __name__ == '__main__':
    main()
