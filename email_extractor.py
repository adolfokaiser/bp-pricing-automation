"""
extract_table_from_email.py

Script para extraer y procesar la tabla de precios desde Outlook a Excel.
Lee remitente, ruta de Excel y nombres de hoja desde un archivo .env.
"""

import os
import pyperclip
import win32com.client
import xlwings as xw
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from dotenv import load_dotenv

# ─── Cargar configuración desde .env ─────────────────────────
load_dotenv()

SENDER_NAME       = os.getenv("SENDER_NAME", "Remitente Nombre")
SENDER_EMAIL      = os.getenv("SENDER_EMAIL", "email@dominio.com")
EXCEL_FILE_PATH   = os.getenv("EXCEL_FILE_PATH", r"C:\ruta\por\defecto\archivo.xlsx")
PRECIOS_SHEET     = os.getenv("PRECIOS_SHEET", "Precios")
GPO_EMP_SHEET     = os.getenv("GPO_EMP_SHEET", "Gpo-Emp")

# ─── Función principal ────────────────────────────────────────
def extract_table_from_email():
    print("Iniciando extracción de tabla desde Outlook...")

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox   = outlook.GetDefaultFolder(6)
    messages = inbox.Items

    print(f"Buscando correo con asunto 'Precios' de {SENDER_NAME}...")
    latest_message = None
    latest_time = None
    for msg in messages:
        if 'Precios' in msg.Subject and (
           msg.SenderName == SENDER_NAME or
           msg.SenderEmailAddress == SENDER_EMAIL
        ):
            if latest_message is None or msg.ReceivedTime > latest_time:
                latest_message = msg
                latest_time = msg.ReceivedTime

    if not latest_message:
        print("❌ No se encontró correo válido.")
        return

    print("Correo encontrado. Procesando tabla...")
    soup = BeautifulSoup(latest_message.HTMLBody, 'html.parser')
    table = soup.find('table')
    if not table:
        print("❌ No se encontró tabla en el correo.")
        return

    html_clip = f"<html><body>{table}</body></html>"
    pyperclip.copy(html_clip)
    print("Tabla copiada al portapapeles.")

    # Abrir Excel y pegar
    app = xw.App(visible=False)
    app.api.DisplayAlerts = False
    wb = xw.Book(EXCEL_FILE_PATH)
    ws = wb.sheets[PRECIOS_SHEET]
    ws.activate()
    ws.range('A1').api.Paste()
    print(f"Tabla pegada en hoja '{PRECIOS_SHEET}'.")

    # Copiar columna adicional desde GPO_EMP_SHEET
    header = ws.range('2:2').value
    diesel_idx = next((i+1 for i,v in enumerate(header) if v == "Diesel"), None)
    if diesel_idx:
        src = wb.sheets[GPO_EMP_SHEET].range('B2:C48')
        ws.cells(2, diesel_idx+1).options(expand='table').value = src.value
    else:
        print("⚠️ Columna 'Diesel' no encontrada; se omite copia extra.")

    # Eliminar filas sin color en columnas C–H
    last_row = ws.range(f'A{ws.cells.last_cell.row}').end('up').row
    deleted = 0
    for r in range(last_row, 2, -1):
        fila_sin_color = True
        for col in ('C','D','E','F','G','H'):
            color = ws.range(f'{col}{r}').color
            if color not in (None, (255,255,255)):
                fila_sin_color = False
                break
        if fila_sin_color:
            ws.range(f'A{r}:K{r}').delete(shift='up')
            deleted += 1

    print(f"Filas sin formato eliminadas: {deleted}")

    # Restaurar filtros en fila 2
    if ws.api.AutoFilterMode:
        ws.api.AutoFilter.ShowAllData()
    last_col = ws.range(2, ws.cells(2, ws.cells.last_cell.column)
                       .end('left').column).column
    col_letter = chr(64 + last_col)
    ws.range(f"A2:{col_letter}2").api.AutoFilter(Field=1)

    # Guardar y cerrar
    wb.save()
    app.api.DisplayAlerts = True
    wb.close()
    app.quit()

    print("✅ Proceso completado exitosamente.")

if __name__ == "__main__":
    extract_table_from_email()
