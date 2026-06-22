import os
import openpyxl
import win32com.client
from logger import get_logger

logger = get_logger()

# Índices de columna (1-based, según estructura del Excel)
COL_CONTRATO = 1       # A
COL_CLIENTE = 3        # C
COL_COMENTARIOS = 6    # F - notas manuales previas
COL_DESCARGA = 7       # G - filtro "Pendiente"
COL_USUARIO = 10       # J - email del usuario
COL_RESULTADO = 12     # L - el bot escribe aquí


def get_pending_rows(filepath: str) -> list[dict]:
    """Lee el Excel y retorna las filas donde col G == 'Pendiente'."""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb["Contratos"]

    pending = []
    for row in ws.iter_rows(min_row=2):
        descarga = row[COL_DESCARGA - 1].value
        if descarga is None:
            continue
        if str(descarga).strip().lower() != "pendiente":
            continue

        contrato = row[COL_CONTRATO - 1].value
        cliente = row[COL_CLIENTE - 1].value
        usuario_email = row[COL_USUARIO - 1].value
        comentario_previo = row[COL_COMENTARIOS - 1].value
        row_number = row[0].row

        pending.append({
            "row_number": row_number,
            "contrato": str(contrato).strip() if contrato is not None else "",
            "cliente": str(cliente).strip() if cliente is not None else "",
            "usuario_email": str(usuario_email).strip() if usuario_email is not None else "",
            "comentario_previo": str(comentario_previo).strip() if comentario_previo is not None else "",
        })

    wb.close()
    logger.info(f"Contratos pendientes encontrados: {len(pending)}")
    return pending


def write_comment(filepath: str, row_number: int, comment: str) -> None:
    """Escribe el comentario en col L usando Excel real (win32com) para preservar Power Query y fórmulas."""
    abs_path = os.path.abspath(filepath)
    xl = None
    wb = None
    try:
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.ScreenUpdating = False
        wb = xl.Workbooks.Open(abs_path)
        ws = wb.Sheets("Contratos")
        ws.Cells(row_number, COL_RESULTADO).Value = str(comment)
        wb.Save()
        logger.debug(f"Fila {row_number} → col L: '{comment}'")
    finally:
        if wb is not None:
            wb.Close(False)
        if xl is not None:
            xl.Quit()
