"""
Script para extraer cambios de precio coloreados de un archivo Excel.
Lee configuración (ruta, hoja y nombres de grupo) desde un archivo .env.
"""

import os
import xlwings as xw
from dotenv import load_dotenv

# ─── Cargar configuración desde .env ─────────────────────────
load_dotenv()

EXCEL_FILE_PATH = os.getenv("EXCEL_OPE_PATH", "ruta/por/defecto.xlsx")
SHEET_NAME      = os.getenv("PRECIOS_SHEET", "Precios")
GROUP_A_NAME    = os.getenv("GROUP_A_NAME", "Arturo Aceves")
GROUP_B_NAME    = os.getenv("GROUP_B_NAME", "Carlos Rodriguez")

# ─── Función para limpiar espacios y caracteres no imprimibles ─
def limpiar_valor(valor):
    if isinstance(valor, str):
        return valor.replace('\xa0', '').strip()
    return valor

# ─── Detectar cambios de precio según color de celda ─────────
def detectar_cambios_precios(ws, row):
    cambios = {}
    # Columnas E, F, G → Regular, Premium, Diésel
    for label, col in (("Regular", "E"), ("Premium", "F"), ("Diésel", "G")):
        color = ws.range(f"{col}{row}").color
        if color not in (None, (255,255,255)):
            cambios[label] = limpiar_valor(ws.range(f"{col}{row}").value)
    return cambios

# ─── Extraer datos de Excel y agrupar por usuario ───────────
def extraer_datos(file_path=EXCEL_FILE_PATH):
    app = xw.App(visible=False)
    wb  = xw.Book(file_path)
    ws  = wb.sheets[SHEET_NAME]

    grupo_a, grupo_b = [], []
    last_row = ws.range(f"A{ws.cells.last_cell.row}").end("up").row

    for row in range(3, last_row + 1):
        empresa   = limpiar_valor(ws.range(f"I{row}").value) or ""
        permiso   = limpiar_valor(ws.range(f"B{row}").value) or ""
        estacion  = limpiar_valor(ws.range(f"C{row}").value) or ""
        grupo     = limpiar_valor(ws.range(f"H{row}").value) or ""
        cambios   = detectar_cambios_precios(ws, row)

        fila = {
            "fila": row,
            "empresa": empresa,
            "permiso": permiso,
            "estacion": estacion,
            "cambios_precios": cambios
        }

        if grupo == GROUP_A_NAME:
            grupo_a.append(fila)
        elif grupo == GROUP_B_NAME:
            grupo_b.append(fila)

    wb.close()
    app.quit()
    return grupo_a, grupo_b

# ─── Ejecución directa ────────────────────────────────────────
if __name__ == "__main__":
    a_data, b_data = extraer_datos()
    print(f"\nCambios detectados para {GROUP_A_NAME}:")
    for d in a_data:
        if d["cambios_precios"]:
            print(f"  Fila {d['fila']}: {d['cambios_precios']}")

    print(f"\nCambios detectados para {GROUP_B_NAME}:")
    for d in b_data:
        if d["cambios_precios"]:
            print(f"  Fila {d['fila']}: {d['cambios_precios']}")
