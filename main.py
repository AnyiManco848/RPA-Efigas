import os
import sys
import signal
from collections import defaultdict
from playwright.sync_api import TimeoutError as PlaywrightTimeout
from playwright.sync_api import Error as PlaywrightError

from logger import get_logger
from config import get_credentials, EXCEL_PATH, PDFS_FOLDER
from excel_handler import get_pending_rows, write_comment
from browser import create_browser, close_browser
from efigas_portal import login, select_contract, validate_current_month, click_ver_factura
from downloader import save_pdf

logger = get_logger()

# Referencias globales para el handler de señales
_current_pw = None
_current_browser = None


def _cleanup(sig=None, frame=None):
    logger.info("Señal de interrupción recibida. Cerrando...")
    if _current_pw and _current_browser:
        close_browser(_current_pw, _current_browser)
    sys.exit(0)


signal.signal(signal.SIGINT, _cleanup)
signal.signal(signal.SIGTERM, _cleanup)


def _process_single_contract(page, row: dict) -> str:
    """Navega, valida y descarga un contrato. Retorna el comentario a escribir.
    Puede lanzar PlaywrightTimeout/PlaywrightError si el navegador falla."""
    contract = row['contrato']

    if not select_contract(page, contract):
        return 'Contrato no existe'

    if not validate_current_month(page):
        return 'Sin emisión'

    # Verificar que el botón exista antes de intentar el click
    # (contratos en estado "Pagado" no muestran "Ver factura")
    if not page.query_selector('button:has-text("Ver factura")'):
        logger.info(f"Contrato {contract}: sin botón 'Ver factura' (posiblemente pagado)")
        return 'Sin emisión'

    try:
        download = click_ver_factura(page)
        ok = save_pdf(download, contract, PDFS_FOLDER)
        return 'Descargada' if ok else 'Error descarga'
    except Exception as e:
        logger.error(f"Error al descargar PDF del contrato {contract}: {e}")
        return 'Error descarga'


def _process_user_group(email: str, rows: list) -> None:
    """Lanza el navegador, hace login y procesa todos los contratos del usuario."""
    global _current_pw, _current_browser

    # Obtener credenciales desde .env
    try:
        _, password = get_credentials(email)
    except ValueError as e:
        logger.error(str(e))
        for row in rows:
            write_comment(EXCEL_PATH, row['row_number'], 'Error descarga')
        return

    # Iniciar navegador y login
    pw, browser, page = create_browser()
    _current_pw, _current_browser = pw, browser

    if not login(page, email, password):
        logger.error(f"Login fallido para {email}. Saltando {len(rows)} contratos.")
        close_browser(pw, browser)
        _current_pw = _current_browser = None
        for row in rows:
            write_comment(EXCEL_PATH, row['row_number'], 'Error descarga')
        return

    i = 0
    while i < len(rows):
        row = rows[i]
        contract = row['contrato']
        logger.info(f"Contrato {contract} — fila {row['row_number']}")

        try:
            comment = _process_single_contract(page, row)

        except (PlaywrightTimeout, PlaywrightError) as e:
            # El navegador falló: reiniciar y volver a intentar este mismo contrato
            logger.error(f"Error de navegador en contrato {contract}: {e}. Reiniciando...")
            try:
                close_browser(pw, browser)
            except Exception:
                pass

            pw, browser, page = create_browser()
            _current_pw, _current_browser = pw, browser

            if not login(page, email, password):
                logger.error("Relogin fallido. Marcando restantes como error.")
                for remaining in rows[i:]:
                    write_comment(EXCEL_PATH, remaining['row_number'], 'Error descarga')
                    logger.info(f"Contrato {remaining['contrato']}: Error descarga")
                break

            # Reintentar el mismo contrato (sin incrementar i)
            try:
                comment = _process_single_contract(page, row)
            except Exception as retry_e:
                logger.error(f"Reintento fallido para {contract}: {retry_e}")
                comment = 'Error descarga'

        write_comment(EXCEL_PATH, row['row_number'], comment)
        logger.info(f"Contrato {contract}: {comment}")
        i += 1

    close_browser(pw, browser)
    _current_pw = _current_browser = None


def main() -> None:
    logger.info("=== Bot Efigas iniciado ===")

    # Validar que el Excel existe
    if not os.path.exists(EXCEL_PATH):
        logger.error(f"Excel no encontrado: {EXCEL_PATH}")
        sys.exit(1)

    # Asegurar que existe la carpeta de PDFs
    os.makedirs(PDFS_FOLDER, exist_ok=True)

    # Leer contratos pendientes
    rows = get_pending_rows(EXCEL_PATH)
    if not rows:
        logger.info("Sin contratos pendientes. Fin.")
        return

    # Agrupar por usuario (columna H)
    groups: dict = defaultdict(list)
    for row in rows:
        groups[row['usuario_email']].append(row)

    logger.info(f"Usuarios: {list(groups.keys())} — Total contratos: {len(rows)}")

    # Procesar cada usuario con su propio navegador
    for email, user_rows in groups.items():
        logger.info(f"--- Usuario: {email} ({len(user_rows)} contratos) ---")
        _process_user_group(email, user_rows)

    logger.info("=== Bot Efigas finalizado ===")


if __name__ == '__main__':
    main()
