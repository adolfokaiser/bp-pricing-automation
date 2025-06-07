"""
extract_table_from_email.py

Script para extraer la tabla de correos de Outlook.
"""

import win32com.client
from win32com.client import constants
import xlwings as xw
from bs4 import BeautifulSoup
import pyperclip
import os

# ——— CONFIGURACIÓN ———
SENDER_NAME       = "Nombre Remitente"         
SENDER_EMAIL      = "email@dominio.com"          
EXCEL_FILE_PATH   = r"C:\ruta\a\Actualizacion_precios_OPE.xlsx" 

def extract_table_from_email():
    """Busca el correo más reciente y pega su tabla en Excel."""
    print("Buscando el correo más reciente con asunto 'Precios'...")

    # Conectar a Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox   = outlook.GetDefaultFolder(6)  # Bandeja de entrada
    messages = inbox.Items

    latest_message      = None
    latest_received_time = None

    # Encontrar mensaje más reciente que cumpla condiciones
    for message in messages:
        if 'Precios' in message.Subject and (
           message.SenderName == SENDER_NAME or
           message.SenderEmailAddress == SENDER_EMAIL
        ):
            if latest_message is None or message.ReceivedTime > latest_received_time:
                latest_message = message
                latest_received_time = message.ReceivedTime

    if not latest_message:
        print("❌ No se encontró ningún correo válido.")
        return

    print("Correo encontrado. Extrayendo tabla...")

    # Parsear HTML y copiar la primera tabla al portapapeles
    soup  = BeautifulSoup(latest_message.HTMLBody, 'html.parser')
    table = soup.find('table')
    if not table:
        print("❌ No se encontró tabla en el cuerpo del correo.")
        return

    html_table = f"<html><body>{table}</body></html>"
    pyperclip.copy(html_table)

    # Abrir Excel en segundo plano y pegar
    with xw.App(visible=False) as app:
        app.api.DisplayAlerts   = False
        app.api.ScreenUpdating  = False
        app.api.Calculation     = constants.xlCalculationManual

        wb    = app.books.open(EXCEL_FILE_PATH)
        sheet = wb.sheets['Precios']

        # Pegar la tabla en A1
        sheet.api.Paste(sheet.api.Range("A1"))

        # ——— Ejemplo: copiar datos de otra hoja si hace falta ———
        # vals = wb.sheets['Gpo-Emp'].range('B2:C48').value
        # diesel_col = next((i+1 for i,v in enumerate(sheet.range('2:2').value) if v=="Diesel"), None)
        # if diesel_col:
        #     sheet.range((2, diesel_col)).options(expand='table').value = vals

        # Eliminar filas ejemplo
        for row in (21, 18, 15):
            sheet.api.Rows(row).Delete()

        # Borrar A42:G42 si la estación es 'Malecon'
        station = sheet.range('A42').value or ""
        if isinstance(station, str) and station.strip().lower() == 'malecon':
            sheet.api.Range("A42:G42").Delete(Shift=constants.xlShiftUp)

        # Restaurar filtros en la primera fila de datos (fila 2)
        if sheet.api.AutoFilterMode:
            sheet.api.AutoFilter.ShowAllData()
        last_col = sheet.cells(2, sheet.cells.last_cell.column).end('left').column
        col_letter = chr(64 + last_col)
        sheet.range(f"A2:{col_letter}2").api.AutoFilter(Field=1)

        # Guardar y cerrar
        wb.save()
        app.api.Calculation    = constants.xlCalculationAutomatic
        app.api.ScreenUpdating = True

    print("¡Proceso completado!")

if __name__ == "__main__":
    extract_table_from_email()
