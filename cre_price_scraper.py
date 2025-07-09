import time  # For time delay, timer
import requests  # For interacting with API
import urllib3
from urllib3.exceptions import InsecureRequestWarning
import sys
import pandas as pd  # For processing data & interact with excel

# For web scraping with browser
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# For running Chrome browser
options = Options()
options.add_argument("--incognito")  # Run in incognito
options.add_argument(r"profile-directory=Default")  # Use default profile
options.add_argument("--headless=new")  # Run in background (sin abrir ventana)
options.add_argument("--ignore-certificate-error")
options.add_argument("--ignore-ssl-errors")
options.add_experimental_option('excludeSwitches', ['enable-logging'])

# Configura el path del ChromeDriver en tu laptop
webdriverPath = Service(executable_path=r"C:\SeleniumDrivers\chromedriver.exe")
driver = webdriver.Chrome(service=webdriverPath, options=options)
try:
    driver.get(r'https://www.cre.gob.mx/ConsultaPrecios/GasolinasyDiesel/GasolinasyDiesel.html')
except:
    executable_path = ChromeDriverManager().install()
    webdriverPath = Service(executable_path=executable_path)
    driver = webdriver.Chrome(service=webdriverPath, options=options)
    driver.get(r'https://www.cre.gob.mx/ConsultaPrecios/GasolinasyDiesel/GasolinasyDiesel.html')

# Especifica la lista de municipios a scrapear
if sys.argv[3] == "bp":
    bpSite_df = pd.read_excel(r"C:\Users\2514xo\OneDrive - BP\Desktop\Task\Lista oficial de municipios.xlsx")
elif sys.argv[3] == "all":
    bpSite_df = pd.read_excel(r"C:\Users\2514xo\OneDrive - BP\Desktop\Task\nationalLevelList.xlsx")

# Configuración para 2Captcha
req2captcha = ""
req2captcha_token = ""
apikey = ""   # Configura tu API key
sitekey = ""  # Configura tu sitekey
method = "userrecaptcha"
webpage = "https://www.cre.gob.mx/ConsultaPrecios/GasolinasyDiesel/GasolinasyDiesel.html"
urllib3.disable_warnings(InsecureRequestWarning)

idList = []  # Lista global para almacenar IDs de captcha
remainingCount = int  # Contador de municipios restantes

def refillCaptchaToken():
    global idList
    global remainingCount
    post_url = req2captcha + "?key=" + apikey + "&method=" + method + "&googlekey=" + sitekey + "&pageurl=" + webpage
    reserveLimit = 5
    if remainingCount < reserveLimit:
        reserveLimit = remainingCount
        print("Reduciendo el número de captcha reservados a " + str(remainingCount) + ".")
    while len(idList) < reserveLimit:
        xStatCode = 0
        while not (xStatCode == 200):
            try:
                time.sleep(1)
                x = requests.get(post_url, verify=False)
                xStatCode = x.status_code
            except:
                pass
        idList.append(x.text.replace("OK|", ""))
        print("Reservado " + str(len(idList)) + "/" + str(reserveLimit) + " código(s) captcha.")
        time.sleep(2)

def claimCatpchaToken():
    global idList
    claimFail = True
    while claimFail:
        for claimId in idList:
            token_url = req2captcha_token + "?key=" + apikey + "&action=get&id=" + claimId
            try:
                y = requests.get(token_url, verify=False)
            except:
                print("Error al reclamar el captcha. Reintentando en 5 segundos")
                time.sleep(5)
                break
            if len(y.text) > 60:  # Reclamo exitoso
                idList.remove(claimId)
                captchaToken = y.text.replace("OK|", "")
                claimFail = False
                print(claimId + " reclamada exitosamente.")
                break
            elif y.text == "CAPCHA_NOT_READY":
                print(claimId + " aún no está lista.")
                time.sleep(2)
            else:
                idList.remove(claimId)
                refillCaptchaToken()
                print("Error: " + y.text + ". Re-solicitando captcha.")
    return captchaToken

def webScraping(fed, municipios):
    global idList
    driver.refresh()
    refillCaptchaToken()
    failure = True
    while failure:
        try:
            select1 = Select(driver.find_element(By.ID, "entidadFederativa"))
            select1.select_by_visible_text(fed)
            time.sleep(2)
            select2 = Select(driver.find_element(By.ID, "municipio"))
            select2.select_by_visible_text(municipios)
            print("Municipio seleccionado. Continuando.")
        except:
            print("No se pudo seleccionar el municipio. Reintentando...")
            pass
        try:
            captchaToken = claimCatpchaToken()
            driver.execute_script('var element=document.getElementById("g-recaptcha-response"); element.style.display="";')
            driver.execute_script("document.getElementById('g-recaptcha-response').innerHTML = " + "'" + captchaToken + "'")
            driver.execute_script(""" $("input[class='btn btn-primary']").click() """)
            pending = WebDriverWait(driver, 10)
            print("Verificando envío de consulta...")
            pending.until(EC.visibility_of_element_located((By.ID, 'precios_next')))
            print("Consulta enviada.")
            failure = False
        except:
            print("El código pudo haber expirado. Refrescando y reintentando...")
            driver.refresh()
            refillCaptchaToken()
            pass

    select3 = Select(driver.find_element(By.NAME, "precios_length"))
    select3.select_by_visible_text("100")
    time.sleep(3)
    df2 = pd.DataFrame()
    failure = True
    failCount = 0
    while failure:
        try:
            lastPage = driver.execute_script("""
                var count = document.getElementsByClassName('paginate_button ')[(document.getElementsByClassName('paginate_button ').length) - 2 ].textContent;
                return count
            """)
            pageCount = int(lastPage)
            failure = False
        except:
            failCount += 1
            if failCount == 3:
                pageCount = 0
                break
            print("No se encontró datos para " + fed + ", " + municipios)
            time.sleep(3)

    page = 1
    df = pd.DataFrame()
    while page <= pageCount:
        time.sleep(2)
        jsTable = driver.execute_script("""
            var table = document.querySelector("table");
            var ths = table.querySelectorAll("tr[role='row']");
            var price = [...ths].map(th => {
                return th.innerText;
            });
            return price;
            """)
        df = pd.DataFrame(jsTable)
        for row in df:
            df = df[row].str.split("\t", expand=True)
        df = df.iloc[1:]
        df2 = pd.concat([df2, df])
        driver.execute_script(""" $("a[id='precios_next']").click() """)
        print("Página " + str(page) + "/" + str(pageCount) + " raspada.")
        page += 1
    return df2

def loopSites(fromIdx, toIdx):
    global remainingCount
    df = pd.DataFrame()
    fromIdxCopy = fromIdx
    while fromIdx <= toIdx:
        fed = bpSite_df["Estados"][fromIdx-2]
        municipios = bpSite_df["Municipios"][fromIdx-2]
        print("Solicitando datos para " + fed + ", " + municipios)
        remainingCount = (toIdx - fromIdx + 1)
        df2 = webScraping(fed, municipios)
        df = pd.concat([df, df2])
        fromIdx += 1
    df = df.reset_index(drop=True)
    # Define la carpeta de trabajo con la nueva ruta
    if sys.argv[3] == "all":
        workFileDir = r"C:"
    elif sys.argv[3] == "bp":
        workFileDir = r"C:"
    df.to_excel(workFileDir + str(fromIdxCopy) + '-' + str(toIdx) + '.xlsx', index=False)
    
startTime = time.time()
argInput = sys.argv  # Ejemplo de uso: python main.py 1 10 all
loopSites(int(argInput[1]), int(argInput[2]))
driver.close()
executionTime = (time.time() - startTime)
print('¡Finalizado! Tiempo de ejecución = ' + str(executionTime))
