import os
import urllib.request
from playwright.sync_api import Download
from logger import get_logger

logger = get_logger()


def save_pdf(download_or_url, contract_number: str, pdfs_folder: str) -> bool:
    """Guarda el PDF como {pdfs_folder}/{contract_number}.pdf.
    Acepta un objeto Download de Playwright o una URL directa.
    Retorna True si el archivo quedó guardado con tamaño > 0."""
    dest = os.path.join(pdfs_folder, f"{contract_number}.pdf")
    try:
        if isinstance(download_or_url, Download):
            download_or_url.save_as(dest)
        else:
            urllib.request.urlretrieve(download_or_url, dest)

        size = os.path.getsize(dest)
        if size == 0:
            logger.error(f"PDF vacío tras la descarga: {dest}")
            return False

        logger.info(f"PDF guardado: {dest} ({size} bytes)")
        return True

    except Exception as e:
        logger.error(f"Error al guardar PDF contrato {contract_number}: {e}")
        return False
