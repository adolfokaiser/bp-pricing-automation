"""
Script de automatización.
Lee credenciales y rutas sensibles desde un archivo .env.
"""

import os
import time
import base64
import sys
import re
from datetime import datetime, timedelta

import pandas as pd
import urllib3
from dotenv import load_dotenv

# Carga variables de entorno desde .env
load_dotenv()

# Evita warnings SSL
os.environ['WDM_SSL_VERIFY'] = '0'
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ───────────────────  UTILIDADES DE CONSOLA  ───────────────────
_BAR_FULL   = "■"
_BAR_EMPTY  = "·"
_BAR_LEN    = 30
_SEP        = "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
script_start_time = time.time()
station_timer     = None

def _fmt_dur(sec: float) -> str:
    """Formatea segundos a 'X h Y min Z s', 'Y min Z s' o 'Z s'."""
    m, s = divmod(int(sec), 60)
    h, m = divmod(m, 60)
    if h:
        return f"{h} h {m} min {s} s"
    if m:
        return f"{m} min {s} s"
    return f"{s} s"

def banner_inicio(fecha, hora, carpeta, n_usuario1, n_usuario2):
    print(f"🚀 Iniciando proceso automático | Fecha: {fecha} | Hora: {hora} 🚀")
    print(_SEP)
    print("📁 Carpeta de Acuses creada:")
    print(f"└─ 📂 \"{carpeta}\"")
    print("📧 Extracción de datos desde correo:")
    print("✔️ Tabla extraída correctamente y almacenada en Excel.")
    print(f"└─ 👤 Usuario A → {n_usuario1} estaciones asignadas")
    print(f"└─ 👤 Usuario B → {n_usuario2} estaciones asignadas")
    print(_SEP)

def barra_progreso(done, total):
    pct = done / total if total else 1
    llenos = int(pct * _BAR_LEN)
    vacios = _BAR_LEN - llenos
    return f"🟢 Progreso: [{_BAR_FULL*llenos}{_BAR_EMPTY*vacios}] {pct*100:>3.0f}% ({done}/{total})"

def login_usuario(nombre):
    print(f"👤 Iniciando sesión → {nombre} ✔️")
    print(_SEP)

def msg_omitida(estacion):
    print(f"✅ {estacion} → Acuse existente, saltado.")

def inicio_estacion(idx, total, estacion, completadas):
    global station_timer
    station_timer = time.time()
    print(barra_progreso(completadas, total))
    print(f"🚩 Procesando estación #{idx}/{total}: {estacion}")

def producto_ok(prod):
    print(f"├─ 🛢️ Producto: {prod} ✔️")

def pdf_ok(ruta):
    dur = _fmt_dur(time.time() - station_timer)
    print("🔏 Firmando documento...")
    print("📄 PDF generado y guardado correctamente:")
    print(f"└─ 📌 \"{ruta}\"")
    print(f"⏱️ Tiempo en estación: {dur}")
    print(_SEP)

def cierre_global():
    print("\n" + _SEP)
    print(f"🎉 Proceso terminado | Tiempo total: {_fmt_dur(time.time() - script_start_time)} 🎉")
    print(_SEP)

def pretty_log(*_args, **_kwargs):
    """Stub para compatibilidad—ignora llamadas heredadas."""
    pass

# ─────────────────────  SELENIUM IMPORTS  ──────────────────────
from email_extractor2 import extract_table_from_email
from data_Extractor2   import extraer_datos

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

TIMEOUT_CLICK = 15
TIMEOUT_INPUT = 5

# ─────────────────── FUNCIONES SELENIUM AUXILIARES ───────────────────
def click_element(driver, locator, timeout=TIMEOUT_CLICK, sleep_time=0.5):
    try:
        el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable(locator))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(sleep_time)
        try:
            el.click()
        except ElementClickInterceptedException:
            driver.execute_script("arguments[0].click();", el)
        return el
    except TimeoutException as e:
        print(f"Error: timeout esperando {locator}: {e}")
        raise

def web_scraping(username, password, existing_driver=None):
    """Inicia sesión en OPE y regresa el driver listo."""
    options = Options()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--start-maximized')
    options.add_argument('--log-level=3')
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    driver = existing_driver or webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

    driver.get('https://ope.cne.gob.mx/Seguridad/InicioSesion')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'txtUsuarioNombre')))

    # Ingreso de credenciales
    usr = driver.find_element(By.ID, 'txtUsuarioNombre')
    pwd = driver.find_element(By.ID, 'txtContrasena')
    usr.clear(); usr.send_keys(username)
    pwd.clear(); pwd.send_keys(password)
    pwd.send_keys(Keys.TAB)

    click_element(driver, (By.ID, 'btnIniciarSesion'))
    WebDriverWait(driver, 10).until_not(EC.url_contains("InicioSesion"))

    return driver

def regresar_a_asistente_por_link(driver):
    """Vuelve a la pantalla de captura tras generar un PDF."""
    driver.switch_to.window(driver.window_handles[0])
    driver.get('https://ope.cne.gob.mx/Wizard/Index')
    # Asume que ya cargan widgets de captura
    click_element(driver, (By.CSS_SELECTOR, 'input[value="2"]'))
    click_element(driver, (By.ID, 'btnSiguiente'))

def regresar_a_inicio(driver):
    """Cierra sesión y regresa al login."""
    driver.switch_to.window(driver.window_handles[0])
    try:
        logout = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, 'Inicio')))
        logout.click()
    except Exception:
        pass
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'txtUsuarioNombre')))

def seleccionar_empresa(driver, empresa_nombre):
    click_element(driver, (By.ID, 's2id_DdlEmpresas'))
    drop = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "select2-drop")))
    inp  = drop.find_element(By.CSS_SELECTOR, "input.select2-input")
    inp.send_keys(empresa_nombre)
    inp.send_keys(Keys.ENTER)

def seleccionar_permiso(driver, cre_value):
    click_element(driver, (By.ID, 'DdlPermisos'))
    opciones = WebDriverWait(driver, 5).until(lambda d: d.find_elements(By.TAG_NAME, 'option'))
    for op in opciones:
        if cre_value.lower() in op.text.lower():
            op.click(); break
    click_element(driver, (By.ID, 'btnSiguiente'))

def seleccionar_fila_producto(driver, producto):
    xpath = f"//tr[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{producto.lower()}')]"
    click_element(driver, (By.XPATH, xpath))

def seleccionar_fecha_y_hora(driver, date_str):
    click_element(driver, (By.ID, 'txtFechaAplicacion'))
    day = datetime.strptime(date_str, '%d/%m/%Y').day
    click_element(driver, (By.XPATH, f"//a[text()='{day}']"))
    js = "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));"
    elem = driver.find_element(By.ID, 'txtHoraAplicacion')
    driver.execute_script(js, elem, "17:00")

def ingresar_precio(driver, precio, producto):
    inp = click_element(driver, (By.ID, 'txtNuevoPrecio'))
    inp.clear(); inp.send_keys(str(precio))
    chk = driver.find_element(By.ID, 'ckbEsPrecioCorrecto')
    if not chk.is_selected(): chk.click()
    click_element(driver, (By.ID, 'btnGuardarCapturaPrecio'))
    seleccionar_fila_producto(driver, producto)
    click_element(driver, (By.ID, 'btnGuardarCapturaPrecio'))

def firmar_y_enviar(driver):
    click_element(driver, (By.ID, 'cmdEnviar'))

def firmar_documento(driver, cer_path, key_path, pwd):
    click_element(driver, (By.ID, "certificado")).send_keys(cer_path)
    click_element(driver, (By.ID, "clave")).send_keys(key_path)
    pwd_in = click_element(driver, (By.ID, "contrasena")); pwd_in.send_keys(pwd)
    click_element(driver, (By.CSS_SELECTOR, "#btnConcluirFirma"))
    WebDriverWait(driver, 10).until(lambda d: True)

def procesar_fila(driver, fila):
    seleccionar_empresa(driver, fila['empresa'])
    seleccionar_permiso(driver, fila['permiso'])
    cambios = {
        'Regular': fila.get('regular'),
        'Premium': fila.get('premium'),
        'Diésel':  fila.get('diesel'),
    }
    for prod, price in cambios.items():
        if price is None: continue
        seleccionar_fila_producto(driver, prod)
        fecha = datetime.today().strftime('%d/%m/%Y')
        seleccionar_fecha_y_hora(driver, fecha)
        ingresar_precio(driver, price, prod)

# ----------------- MAIN SCRIPT -----------------
def main():
    # Preparar carpeta de acuses
    today       = datetime.today()
    today_fmt   = today.strftime('%d%m%Y')
    folder_name = f"{today_fmt}_Ratificacion"
    base_path   = os.getenv("ACUSES_BASE_PATH")
    new_folder  = os.path.join(base_path, folder_name)
    os.makedirs(new_folder, exist_ok=True)

    # Extraer datos de correo + Excel
    extract_table_from_email()
    excel_path = os.getenv("EXCEL_OPE_PATH")
    data_A, data_B = extraer_datos(excel_path)

    banner_inicio(
        fecha=today.strftime('%d/%m/%Y'),
        hora=today.strftime('%H:%M'),
        carpeta=new_folder,
        n_usuario1=len(data_A),
        n_usuario2=len(data_B)
    )

    omit = os.getenv("OMITIR_ESTACIONES", "").split(",")

    # Bloque Usuario A
    if data_A:
        login_usuario("Usuario A")
        drvA = web_scraping(
            os.getenv("ARTURO_EMAIL"),
            os.getenv("ARTURO_PASS")
        )
        cerA = os.getenv("ARTURO_CER_PATH")
        keyA = os.getenv("ARTURO_KEY_PATH")
        pwdA = os.getenv("ARTURO_KEY_PWD")

        totalA = sum(1 for f in data_A if f['estacion'] not in omit)
        doneA  = 0
        for idx, fila in enumerate(data_A, start=1):
            est = fila['estacion']
            if est in omit or os.path.exists(os.path.join(new_folder, f"{est}.pdf")):
                doneA += 1; continue
            inicio_estacion(idx, totalA, est, doneA)
            try:
                procesar_fila(drvA, fila)
                firmar_y_enviar(drvA)
                firmar_documento(drvA, cerA, keyA, pwdA)
                # Generar PDF vía CDP
                WebDriverWait(drvA, 10).until(lambda d: len(d.window_handles) > 1)
                drvA.switch_to.window(drvA.window_handles[-1])
                pdf_data = drvA.execute_cdp_cmd("Page.printToPDF", {"printBackground": True})
                pdf_path = os.path.join(new_folder, f"{est}.pdf")
                with open(pdf_path, "wb") as f:
                    f.write(base64.b64decode(pdf_data['data']))
                drvA.close()
                drvA.switch_to.window(drvA.window_handles[0])
                pdf_ok(pdf_path)
            except Exception as e:
                print(f"❌ ERROR en {est}: {e}")
            finally:
                doneA += 1
                if doneA < totalA:
                    regresar_a_asistente_por_link(drvA)
        regresar_a_inicio(drvA)

    # Bloque Usuario B
    if data_B:
        login_usuario("Usuario B")
        drvB = web_scraping(
            os.getenv("CARLOS_EMAIL"),
            os.getenv("CARLOS_PASS")
        )
        cerB = os.getenv("CARLOS_CER_PATH")
        keyB = os.getenv("CARLOS_KEY_PATH")
        pwdB = os.getenv("CARLOS_KEY_PWD")

        totalB = sum(1 for f in data_B if f['estacion'] not in omit)
        doneB  = 0
        for idx, fila in enumerate(data_B, start=1):
            est = fila['estacion']
            if est in omit or os.path.exists(os.path.join(new_folder, f"{est}.pdf")):
                doneB += 1; continue
            inicio_estacion(idx, totalB, est, doneB)
            try:
                procesar_fila(drvB, fila)
                firmar_y_enviar(drvB)
                firmar_documento(drvB, cerB, keyB, pwdB)
                WebDriverWait(drvB, 10).until(lambda d: len(d.window_handles) > 1)
                drvB.switch_to.window(drvB.window_handles[-1])
                pdf_data = drvB.execute_cdp_cmd("Page.printToPDF", {"printBackground": True})
                pdf_path = os.path.join(new_folder, f"{est}.pdf")
                with open(pdf_path, "wb") as f:
                    f.write(base64.b64decode(pdf_data['data']))
                drvB.close()
                drvB.switch_to.window(drvB.window_handles[0])
                pdf_ok(pdf_path)
            except Exception as e:
                print(f"❌ ERROR en {est}: {e}")
            finally:
                doneB += 1
                if doneB < totalB:
                    regresar_a_asistente_por_link(drvB)
        drvB.quit()

    cierre_global()

if __name__ == "__main__":
    main()
