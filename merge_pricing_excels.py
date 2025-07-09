import os
import glob
import pandas as pd
import time
import sys
from send2trash import send2trash
import win32com.client as win32
from datetime import datetime

# Configura las rutas de trabajo en base al argumento (all o bp)
if sys.argv[1] == "all":
    workFileDir = r"C:\Users\2514xo\OneDrive - BP\Desktop\Task\workfileNational"
    totalExpectedFile = 15
elif sys.argv[1] == "bp":
    workFileDir = r"C:\Users\2514xo\OneDrive - BP\Desktop\Task\workfileBP"
    totalExpectedFile = 6

column_names = ["No de Permiso", "Nombre de la gasolinera", 'Dirección', 'Producto', 'Subproducto', 'Precio Registrado']
main_df = pd.DataFrame()

totalFile = 0
while totalFile < totalExpectedFile:
    os.chdir(workFileDir)
    files = glob.glob('*.xlsx')
    totalFile = len(files)
    if totalFile == totalExpectedFile:
        break
    time.sleep(15)

for file in files:
    df = pd.read_excel(workFileDir + "/" + file)
    pathToDelete = workFileDir + "\\" + file
    send2trash(pathToDelete)
    main_df = pd.concat([main_df, df])
print(main_df)
main_df.columns = column_names
print("Combinación completada")

# Configura la ruta de salida para el archivo combinado
if sys.argv[1] == "all":
    main_df.to_excel(r"C:\Users\2514xo\OneDrive - BP\Desktop\Task\combinedNational.xlsx", index=False)
    attachment = r"C:\Users\2514xo\OneDrive - BP\Desktop\Task\combinedNational.xlsx"
elif sys.argv[1] == "bp":
    main_df.to_excel(r"C:\Users\2514xo\OneDrive - BP\Desktop\Task\combinedBP.xlsx", index=False)
    attachment = r"C:\Users\2514xo\OneDrive - BP\Desktop\Task\combinedBP.xlsx"

# Configuración del correo: ahora solo se envía a adolfo.gomez@bp.com
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.Subject = "Today's Pricing data - " + datetime.now().strftime('%#d %b %Y')
mail.To = "adolfo.gomez@bp.com"
# Si ya no se requieren CC, se omite esta línea o se puede comentar
# mail.CC = ""

mail.Attachments.Add(attachment)
mail.HTMLBody = """"""
mail.Send()
