import os
import time
import base64
import sys
import re
from datetime import datetime, timedelta

import pandas as pd
import urllib3
from dotenv import load_dotenv

# ─── Cargar variables de entorno desde .env ─────────────────────
load_dotenv()  # Asegúrate de tener un .env con las claves genéricas:
               # USER_A_EMAIL, USER_A_PASSWORD, USER_A_CERT_PATH, USER_A_KEY_PATH, USER_A_KEY_PWD
               # USER_B_EMAIL, USER_B_PASSWORD, USER_B_CERT_PATH, USER_B_KEY_PATH, USER_B_KEY_PWD
               # ACUSES_BASE_PATH, EXCEL_OPE_PATH, OMITIR_ESTACIONES,...

# Evitar warnings SSL
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
    m, s = divmod(int(sec), 60)
    h, m = divmod(m, 60)
    if h: return f"{h} h {m} min {s} s"
    if m: return f"{m} min {s} s"
    return f"{s} s"

def banner_inicio(fecha, hora, carpeta, n_a, n_b):
    print(f"🚀 Ratificación semanal | Fecha: {fecha} | Hora: {hora} 🚀")
    print(_SEP)
    print("📁 Carpeta de acuses creada:")
    print(f"└─ 📂 \"{carpeta}\"")
    print("📧 Extracción de datos y tabla lista.")
    print(f"├─ 👤 Usuario A → {n_a} estaciones asignadas")
    print(f"└─ 👤 Usuario B → {n_b} estaciones asignadas")
    print(_SEP)

def barra_progreso(done, total):
    pct = 1.0 if total <= 0 else done/total
    llenos = int(pct * _BAR_LEN)
    vacios = _BAR_LEN - llenos
    disp = min(done, total)
    return f"🟢 Progreso: [{'■'*llenos}{'·'*vacios}] {pct*100:>3.0f}% ({disp}/{total})"

def login_usuario(nombre):
    print(_SEP)
    print(f"👤 Iniciando sesión → {nombre} ✔️")
    print(_SEP)

def msg_omitida(estacion):
    print(f"✅ {estacion} → Acuse existente, saltado.")

def inicio_estacion(idx, total, estacion, done):
    global station_timer
    station_timer = time.time()
    print(barra_progreso(done, total))
    print(f"🚩 Procesando estación #{idx}/{total}: {estacion}")

def producto_ok(prod):
    print(f"   🔃 Producto: {prod} ✔️")

def pdf_ok(ruta):
    global station_timer
    dur = time.time() - station_timer if station_timer else 0
    print("🔏 Firmando documento...")
    print("📄 PDF generado y guardado en:")
    print(f"└─ 📌 \"{ruta}\"")
    print(f"⏱️ Tiempo en estación: {_fmt_dur(dur)}")
    print(f"✅ Estación procesada exitosamente.")
    print(_SEP)

def cierre_global():
    print("\n" + _SEP)
    print(f"🎉 Proceso terminado | Tiempo total: {_fmt_dur(time.time() - script_start_time)} 🎉")
    print(_SEP)

# ─────────────────────  SELENIUM IMPORTS  ──────────────────────
from email_extractor import extract_table_from_email
from data_extractor  import extraer_datos

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

# ────────── FUNCIONES SELENIUM AUXILIARES ──────────────────────
def click_element(driver, locator, timeout=TIMEOUT_CLICK, sleep_time=0.5):
    el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable(locator))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(sleep_time)
    try:
        el.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", el)
    return el

def web_scraping(username, password, existing_driver=None):
    opts = Options()
    opts.add_argument('--ignore-certificate-errors')
    opts.add_argument('--start-maximized')
    opts.add_argument('--log-level=3')
    drv = existing_driver or webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=opts
    )
    drv.get(os.getenv("OPE_LOGIN_URL", "https://ope.cne.gob.mx/Seguridad/InicioSesion"))
    WebDriverWait(drv, 10).until(EC.presence_of_element_located((By.ID, 'txtUsuarioNombre')))
    drv.find_element(By.ID, 'txtUsuarioNombre').send_keys(username)
    pwd_inp = drv.find_element(By.ID, 'txtContrasena')
    pwd_inp.send_keys(password)
    pwd_inp.send_keys(Keys.TAB)
    click_element(drv, (By.ID, 'btnIniciarSesion'))
    WebDriverWait(drv, 10).until_not(EC.url_contains("InicioSesion"))
    return drv

def regresar_a_asistente(driver):
    driver.switch_to.window(driver.window_handles[0])
    click_element(driver, (By.LINK_TEXT, 'Asistente'))
    click_element(driver, (By.CSS_SELECTOR, 'input[value="2"]'))
    click_element(driver, (By.ID, 'cbxTerminos'))
    click_element(driver, (By.ID, 'btnSiguiente'))

def regresar_a_inicio(driver):
    try:
        driver.switch_to.window(driver.window_handles[0])
        driver.execute_script("document.getElementById('logoutForm').submit();")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'txtUsuarioNombre')))
        print("✅ Logout exitoso.")
        print(_SEP)
    except Exception as e:
        print(f"❌ Error logout: {e}")

def seleccionar_empresa(driver, empresa):
    click_element(driver, (By.ID, 's2id_DdlEmpresas'))
    drop = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "select2-drop")))
    inp = drop.find_element(By.CSS_SELECTOR, "input.select2-input")
    inp.send_keys(empresa); inp.send_keys(Keys.ENTER)

def seleccionar_permiso(driver, permiso):
    click_element(driver, (By.ID, 'DdlPermisos'))
    opts = WebDriverWait(driver, 5).until(lambda d: d.find_elements(By.TAG_NAME, 'option'))
    for op in opts:
        if permiso.lower() in op.text.lower():
            op.click(); break
    click_element(driver, (By.ID, 'btnSiguiente'))

def seleccionar_fila(driver, producto):
    xpath = f"//tr[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'{producto.lower()}')]"
    click_element(driver, (By.XPATH, xpath))

def seleccionar_fecha(driver, date_str):
    click_element(driver, (By.ID, 'txtFechaAplicacion'))
    day = datetime.strptime(date_str, '%d/%m/%Y').day
    click_element(driver, (By.XPATH, f"//a[text()='{day}']"))
    hora_in = driver.find_element(By.ID, 'txtHoraAplicacion')
    driver.execute_script("arguments[0].value='17:00';arguments[0].dispatchEvent(new Event('change'));", hora_in)

def ingresar_precio(driver, precio, producto):
    inp = click_element(driver, (By.ID, 'txtNuevoPrecio'))
    inp.clear(); inp.send_keys(str(precio))
    chk = driver.find_element(By.ID, 'ckbEsPrecioCorrecto')
    if not chk.is_selected(): chk.click()
    click_element(driver, (By.ID, 'btnGuardarCapturaPrecio'))
    seleccionar_fila(driver, producto)
    click_element(driver, (By.ID, 'btnGuardarCapturaPrecio'))

def firmar_y_enviar(driver):
    click_element(driver, (By.ID, 'cmdEnviar'))

def firmar_documento(driver, cer, key, pwd):
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "dlgFirma")))
    driver.find_element(By.ID, "certificado").send_keys(cer)
    driver.find_element(By.ID, "clave").send_keys(key)
    click_element(driver, (By.ID, "contrasena")).send_keys(pwd)
    click_element(driver, (By.CSS_SELECTOR, "#btnConcluirFirma"))

def procesar_fila(driver, fila):
    seleccionar_empresa(driver, fila['empresa'])
    seleccionar_permiso(driver, fila['permiso'])
    date_str = (datetime.today() + timedelta(days=1)).strftime('%d/%m/%Y')
    for prod, price in fila['cambios_precios'].items():
        seleccionar_fila(driver, prod)
        seleccionar_fecha(driver, date_str)
        ingresar_precio(driver, price, prod)

# ─────────────────────── MAIN ───────────────────────
def main():
    # Preparar carpeta de acuses
    today      = datetime.today()
    vnext      = today + timedelta(days=(4 - today.weekday()) % 7)
    vfmt       = vnext.strftime('%d%m%Y')
    base_path  = os.getenv("ACUSES_BASE_PATH")
    new_folder = os.path.join(base_path, f"{vfmt}_Ratificacion")
    os.makedirs(new_folder, exist_ok=True)

    # Extraer datos
    extract_table_from_email()
    excel_path     = os.getenv("EXCEL_OPE_PATH")
    data_A, data_B = extraer_datos(excel_path)

    banner_inicio(
        fecha=vnext.strftime('%d/%m/%Y'),
        hora=vnext.strftime('%H:%M'),
        carpeta=new_folder,
        n_a=len(data_A), n_b=len(data_B)
    )

    omit_list = os.getenv("OMITIR_ESTACIONES", "").split(",")

    # Bloque Usuario A
    if data_A:
        login_usuario("Usuario A")
        drv = web_scraping(os.getenv("USER_A_EMAIL"), os.getenv("USER_A_PASSWORD"))
        cer = os.getenv("USER_A_CERT_PATH"); key = os.getenv("USER_A_KEY_PATH"); pwd = os.getenv("USER_A_KEY_PWD")
        total = len(data_A); done = 0

        for idx, fila in enumerate(data_A, start=1):
            est = fila['estacion']
            pdf_path = os.path.join(new_folder, f"{est}.pdf")
            if est in omit_list or os.path.exists(pdf_path):
                msg_omitida(est); done += 1; print(_SEP); continue

            inicio_estacion(idx, total, est, done)
            try:
                procesar_fila(drv, fila)
                firmar_y_enviar(drv)
                firmar_documento(drv, cer, key, pwd)
                WebDriverWait(drv, 10).until(lambda d: len(d.window_handles) > 1)
                drv.switch_to.window(drv.window_handles[-1])
                pdf = drv.execute_cdp_cmd("Page.printToPDF", {"printBackground": True})
                with open(pdf_path, "wb") as f:
                    f.write(base64.b64decode(pdf['data']))
                pdf_ok(pdf_path)
                drv.close(); drv.switch_to.window(drv.window_handles[0])
            except Exception as e:
                print(f"❌ Error en {est}: {e}"); print(_SEP)
            finally:
                done += 1
                if idx < total:
                    regresar_a_asistente(drv)

        regresar_a_inicio(drv)

    # Bloque Usuario B
    if data_B:
        login_usuario("Usuario B")
        drv = web_scraping(
            os.getenv("USER_B_EMAIL"),
            os.getenv("USER_B_PASSWORD"),
            existing_driver=drv if 'drv' in locals() else None
        )
        cer = os.getenv("USER_B_CERT_PATH"); key = os.getenv("USER_B_KEY_PATH"); pwd = os.getenv("USER_B_KEY_PWD")
        total = len(data_B); done = 0

        for idx, fila in enumerate(data_B, start=1):
            est = fila['estacion']
            pdf_path = os.path.join(new_folder, f"{est}.pdf")
            if est in omit_list or os.path.exists(pdf_path):
                msg_omitida(est); done += 1; print(_SEP); continue

            inicio_estacion(idx, total, est, done)
            try:
                procesar_fila(drv, fila)
                firmar_y_enviar(drv)
                firmar_documento(drv, cer, key, pwd)
                WebDriverWait(drv, 10).until(lambda d: len(d.window_handles) > 1)
                drv.switch_to.window(drv.window_handles[-1])
                pdf = drv.execute_cdp_cmd("Page.printToPDF", {"printBackground": True})
                with open(pdf_path, "wb") as f:
                    f.write(base64.b64decode(pdf['data']))
                pdf_ok(pdf_path)
                drv.close(); drv.switch_to.window(drv.window_handles[0])
            except Exception as e:
                print(f"❌ Error en {est} (B): {e}"); print(_SEP)
            finally:
                done += 1
                if idx < total:
                    regresar_a_asistente(drv)

        regresar_a_inicio(drv)
        drv.quit()

    cierre_global()

if __name__ == "__main__":
    main()
