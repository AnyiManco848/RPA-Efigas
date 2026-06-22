"""Limpia la columna J del Excel para volver a correr main.py desde cero."""
import openpyxl

EXCEL_PATH = 'Contratos Efigas.xlsx'
wb = openpyxl.load_workbook(EXCEL_PATH, data_only=False)
ws = wb['Contratos']

cleared = 0
for row in ws.iter_rows(min_row=2):
    descarga = row[6].value  # col G
    if descarga and str(descarga).strip().lower() == 'pendiente':
        old_val = row[9].value  # col J
        row[9].value = None
        if old_val is not None:
            print(f"  Fila {row[0].row}: '{old_val}' -> limpiada")
            cleared += 1

wb.save(EXCEL_PATH)
wb.close()
print(f"Limpiadas {cleared} celdas en col J.")
