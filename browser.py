import os
from playwright.sync_api import sync_playwright, Playwright, Browser, Page
from config import PDFS_FOLDER
from logger import get_logger

logger = get_logger()

DOWNLOADS_PATH = os.path.abspath(PDFS_FOLDER)


def create_browser() -> tuple[Playwright, Browser, Page]:
    """Lanza Chromium headless y retorna (playwright, browser, page)."""
    os.makedirs(DOWNLOADS_PATH, exist_ok=True)

    pw = sync_playwright().start()
    browser = pw.chromium.launch(headless=True)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()
    logger.info("Navegador Chromium iniciado (headless)")
    return pw, browser, page


def close_browser(pw: Playwright, browser: Browser) -> None:
    """Cierra contextos, browser y detiene Playwright sin dejar procesos huérfanos."""
    try:
        for context in browser.contexts:
            context.close()
        browser.close()
        pw.stop()
        logger.info("Navegador cerrado limpiamente")
    except Exception as e:
        logger.warning(f"Error al cerrar el navegador (ignorado): {e}")
