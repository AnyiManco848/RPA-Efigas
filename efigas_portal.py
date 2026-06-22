import re
from datetime import datetime
from playwright.sync_api import Page, Download
from playwright.sync_api import TimeoutError as PlaywrightTimeout
from logger import get_logger

logger = get_logger()

PORTAL_URL = "https://portal.efigas.com.co"

MONTHS_ES = {
    'Enero': 1, 'Febrero': 2, 'Marzo': 3, 'Abril': 4,
    'Mayo': 5, 'Junio': 6, 'Julio': 7, 'Agosto': 8,
    'Septiembre': 9, 'Octubre': 10, 'Noviembre': 11, 'Diciembre': 12,
}
_MONTH_PATTERN = re.compile(
    r'(' + '|'.join(MONTHS_ES.keys()) + r')\s+(\d{4})'
)


def _close_popup(page: Page) -> None:
    try:
        page.click('[aria-label="cerrar diálogo"]', timeout=5000)
        page.wait_for_timeout(500)
        logger.debug("Popup cerrado")
    except PlaywrightTimeout:
        pass


def login(page: Page, email: str, password: str) -> bool:
    """Navega al portal, hace login y retorna True si el dashboard cargó."""
    logger.info(f"Iniciando sesión: {email}")
    page.goto(PORTAL_URL, timeout=30000)
    page.wait_for_load_state('networkidle', timeout=15000)
    _close_popup(page)

    # Abrir modal de login (botón del nav)
    page.click('text=Iniciar sesión', timeout=10000)
    page.wait_for_selector('#email', timeout=10000)

    # Cloudflare Turnstile — suele resolverse solo en headless
    try:
        page.wait_for_selector('text=¡Operación exitosa!', timeout=30000)
        logger.debug("Cloudflare: Operación exitosa")
    except PlaywrightTimeout:
        logger.warning("Cloudflare: sin confirmación en 30s, continuando de todas formas")

    page.fill('#email', email)
    page.fill('#password', password)

    # Botón azul submit del modal
    page.click('button.MuiButton-contained', timeout=10000)

    # Confirmar que el dashboard cargó (aparece el selector "Casa" en el nav)
    try:
        page.wait_for_selector('button:has-text("Casa")', timeout=40000)
        page.wait_for_load_state('networkidle', timeout=10000)
        _close_popup(page)
        logger.info(f"Login exitoso: {email}")
        return True
    except PlaywrightTimeout:
        logger.error(f"Login fallido para {email}: dashboard no cargó en 40s")
        return False


def select_contract(page: Page, contract_number: str) -> bool:
    """Abre el menú Casa y selecciona el contrato. Queda en el dashboard del contrato."""
    logger.info(f"Buscando contrato: {contract_number}")

    page.click('button:has-text("Casa")', timeout=10000)
    page.wait_for_timeout(800)

    item = page.query_selector(f'li[name="{contract_number}"]')
    if not item:
        logger.warning(f"Contrato {contract_number} no existe en el portal")
        page.keyboard.press('Escape')
        page.wait_for_timeout(500)
        return False

    item.click()
    # Esperar que el dashboard muestre los datos de este contrato
    try:
        page.wait_for_function(
            f'() => document.body.innerText.includes("{contract_number}")',
            timeout=12000,
        )
    except PlaywrightTimeout:
        page.wait_for_load_state('networkidle', timeout=10000)
    logger.debug(f"Contrato {contract_number} seleccionado — dashboard cargado")
    return True


def validate_current_month(page: Page) -> bool:
    """Verifica que el mes de la factura visible en el dashboard sea el mes/año actual."""
    try:
        page.wait_for_load_state('networkidle', timeout=10000)
    except PlaywrightTimeout:
        pass

    months_js = '|'.join(MONTHS_ES.keys())
    js_check = f'() => /({months_js})\\s+\\d{{4}}/.test(document.body.innerText)'
    try:
        page.wait_for_function(js_check, timeout=10000)
    except PlaywrightTimeout:
        logger.warning("No se detectó fecha de factura en el dashboard (10s timeout)")
        return False

    body_text = page.evaluate('() => document.body.innerText')
    match = _MONTH_PATTERN.search(body_text)
    if not match:
        return False

    month = MONTHS_ES[match.group(1)]
    year = int(match.group(2))
    now = datetime.now()
    ok = (month == now.month and year == now.year)
    logger.debug(
        f"Factura en dashboard: {match.group(1)} {year} | Hoy: {now.month}/{now.year} | Válida: {ok}"
    )
    return ok


def click_ver_factura(page: Page) -> Download:
    """Hace click en Ver factura desde el dashboard del contrato y retorna el Download."""
    logger.debug("Clickeando Ver factura desde el dashboard")
    with page.expect_download(timeout=30000) as dl_info:
        page.click('button:has-text("Ver factura")', timeout=10000)
    download = dl_info.value
    logger.debug(f"Descarga iniciada: {download.suggested_filename}")
    return download
