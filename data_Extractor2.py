"""
extraer_datos.py

Script para extraer datos de un archivo Excel.
Lee ruta del archivo y nombres de grupo desde un archivo .env.
"""

import os
from dotenv import load_dotenv
import xlwings as xw

# ─── Cargar configuración desde .env ─────────────────────────
load_dotenv()

EXCEL_FILE_PATH = os.getenv("EXCEL_OPE_PATH", "ruta_por_defecto.xlsx")
GROUP_A_NAME    = os.getenv("GROUP_A_NAME", "Group A")
GROUP_B_NAME    = os.getenv("GROUP_B_NAME", "Group B")

# ─── Función para limpiar valores ─────────────────────────────
def limpiar_valor(valor):
    if isinstance(valor, str):
        return valor.replace('\xa0', '').strip()
    return valor

# ─── Extraer datos ────────────────────────────────────────────
def extraer_datos(file_path=EXCEL_FILE_PATH):
    app = xw.App(visible=False)
    wb  = xw.Book(file_path)
    ws  = wb.sheets['Precios']

    grupo_a_data, grupo_b_data = [], []

    # Última fila con datos en la columna A
    last_row = ws.range(f'A{ws.cells.last_cell.row}').end('up').row

    for row in range(3, last_row + 1):
        empresa  = limpiar_valor(ws.range(f'I{row}').value) or 'Sin empresa'
        permiso  = limpiar_valor(ws.range(f'B{row}').value) or ''
        estacion = limpiar_valor(ws.range(f'C{row}').value) or ''
        grupo    = limpiar_valor(ws.range(f'H{row}').value) or ''

        regular = limpiar_valor(ws.range(f'E{row}').value)
        premium = limpiar_valor(ws.range(f'F{row}').value)
        diesel  = limpiar_valor(ws.range(f'G{row}').value)

        fila = {
            'fila':     row,
            'empresa':  empresa,
            'permiso':  permiso,
            'estacion': estacion,
            'regular':  regular,
            'premium':  premium
        }
        # Solo añadir diesel si es válido
        if diesel not in (None, '', '-'):
            fila['diesel'] = diesel

        # Distribuir según grupo
        if grupo == GROUP_A_NAME:
            grupo_a_data.append(fila)
        elif grupo == GROUP_B_NAME:
            grupo_b_data.append(fila)

    wb.close()
    app.quit()
    return grupo_a_data, grupo_b_data

# ─── Impresión formateada de resultados ────────────────────────
def imprimir_datos(label, datos):
    print(f"\nDatos de {label}:")
    for d in datos:
        linea = (
            f"Fila {d['fila']}: Empresa '{d['empresa']}', Permiso '{d['permiso']}', "
            f"Estación '{d['estacion']}', Regular={d['regular']}, Premium={d['premium']}"
        )
        if 'diesel' in d:
            linea += f", Diésel={d['diesel']}"
        print(linea)

# ─── Uso directo ───────────────────────────────────────────────
if __name__ == "__main__":
    arturo, carlos = extraer_datos()
    imprimir_datos(GROUP_A_NAME, arturo)
    imprimir_datos(GROUP_B_NAME, carlos)
