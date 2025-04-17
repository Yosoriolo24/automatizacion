from datetime import datetime, timedelta
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import time

import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials
import os

# CONFIGURA TUS CREDENCIALES Y URL DE OKTA
OKTA_USERNAME = "jh@hot.com"
OKTA_PASSWORD = "jh"
OKTA_URL = "https://openeducation.okta.com"
reporte_casos_url = "https://openeducation.my.salesforce.com/00O3r000006mGPy"

carpeta_descargas = "C:\\Users\\yonat\\OneDrive\\Documentos\\programas\\reporte"

scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("C:\\Users\\yonat\\OneDrive\\Documentos\\programas\\auto.json", scope)
client = gspread.authorize(creds)

# Configurar opciones de Chrome
options = Options()
options.add_experimental_option("prefs", {
    "download.default_directory": carpeta_descargas,  # sin comillas dobles
    "download.prompt_for_download": False,
    "directory_upgrade": True
})

# ABRIR CHROME
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.get(OKTA_URL)

# ESPERA QUE CARGUE LA PÁGINA DE OKTA
try:
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, "input28"))
    )
    print("✅ Página de Okta cargada correctamente.")
except TimeoutException:
    print("❌ Tiempo de espera agotado. No se cargó la página de Okta.")
    driver.quit()
    exit()

# COMPLETAR LOGIN EN OKTA
try:
    driver.find_element(By.ID, "input28").send_keys(OKTA_USERNAME)
    driver.find_element(By.ID, "input36").send_keys(OKTA_PASSWORD)
    driver.find_element(By.XPATH, '//input[@value="Iniciar sesión"]').click()
    print("🔐 Intentando iniciar sesión...")
except NoSuchElementException:
    print("❌ No se encontró el formulario de login.")
    driver.quit()
    exit()

# ESPERA A QUE CARGUE EL DASHBOARD Y CLIC EN SALESFORCE
try:
    salesforce_btn = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, '//a[contains(@href, "salesforce")]'))
    )
    salesforce_btn.click()
    print("🚀 Acceso a Salesforce exitoso.")
    print("✅ Ya estás dentro de Salesforce.")
    # Ir al link del reporte directamente
    
    #reporte de casos   
    print("🔗 Ingreso a informes de Casos...")
    time.sleep(3)
    driver.execute_script("window.open('https://openeducation.my.salesforce.com/00O3r000006mGPy', '_blank');")
    driver.switch_to.window(driver.window_handles[1])

    ayer = datetime.now() - timedelta(days=1)
    fecha_ayer = ayer.strftime('%d/%m/%Y')  # Ajusta el formato si Salesforce espera otro
    print("📅 Fecha del día anterior:", fecha_ayer)
    time.sleep(10)
    date_picker = driver.find_element(By.XPATH, '//*[@id="edate"]')
    date_picker.click()
    date_picker.clear()

    time.sleep(1)
    date_picker.send_keys(fecha_ayer)
    
    driver.find_element(By.XPATH, '//input[@value="Ejecutar informe"]').click()
    time.sleep(2)
    driver.find_element(By.XPATH, '//input[@value="Exportar detalles"]').click()
    time.sleep(3)
    dropdown = Select(driver.find_element(By.ID, "xf"))
    dropdown.select_by_value("localecsv")
    time.sleep(1)
    driver.find_element(By.XPATH, '//input[@value="Exportar"]').click()

    #Copiar reporte de casos
    sheetsCases = client.open_by_key("1FDok-e6IZdtTfoPbCou2St7fmiuw4eOKhzkd4unsHlc")
    sheets = sheetsCases.worksheet("Casos")
    time.sleep(10)

    def archivo_mas_reciente(carpeta, extension=".csv"):
        archivos = [f for f in os.listdir(carpeta) if f.endswith(extension)]
        if not archivos:
            raise FileNotFoundError("No se encontró ningún archivo CSV en la carpeta.")
        
        archivos = [os.path.join(carpeta, f) for f in archivos]
        archivo_reciente = max(archivos, key=os.path.getctime)
        return archivo_reciente
    ruta_archivo = archivo_mas_reciente("C:\\Users\\yonat\\OneDrive\\Documentos\\programas\\reporte")
    print("📄 Archivo más reciente:", ruta_archivo)
    df = pd.read_csv(ruta_archivo, sep=None, engine='python')
    df = df[:-5]

    # Limpiar la hoja anterior (opcional)
    sheets.clear()

    # Escribir los nuevos datos
    set_with_dataframe(sheets, df)

    print("✅ Hoja 'Casos' actualizada en Google Sheets.")
    time.sleep(5)
    
#-----------------------------------------------------------------------------------------------------------------------------#
    #reporte de Ohs   
    print("🔗 Ingreso a informes de Ohs...")
    time.sleep(3)
    driver.execute_script("window.open('https://openeducation.my.salesforce.com/00O0Z0000070nf9','_parent');")
    driver.switch_to.window(driver.window_handles[1])
    driver.find_element(By.XPATH, '//input[@value="Exportar detalles"]').click()
    time.sleep(3)
    dropdown = Select(driver.find_element(By.ID, "xf"))
    dropdown.select_by_value("localecsv")
    time.sleep(1)
    driver.find_element(By.XPATH, '//input[@value="Exportar"]').click()
        
    #Copiar reporte de Ohs
    sheetsCases = client.open_by_key("1FDok-e6IZdtTfoPbCou2St7fmiuw4eOKhzkd4unsHlc")
    sheets = sheetsCases.worksheet("Ohs")
    
    time.sleep(10)
    def archivo_mas_reciente(carpeta, extension=".csv"):
        archivos = [f for f in os.listdir(carpeta) if f.endswith(extension)]
        if not archivos:
            raise FileNotFoundError("No se encontró ningún archivo CSV en la carpeta.")
        
        archivos = [os.path.join(carpeta, f) for f in archivos]
        archivo_reciente = max(archivos, key=os.path.getctime)
        return archivo_reciente
    ruta_archivo = archivo_mas_reciente("C:\\Users\\yonat\\OneDrive\\Documentos\\programas\\reporte")
    print("📄 Archivo más reciente:", ruta_archivo)
    df = pd.read_csv(ruta_archivo, sep=None, engine='python')
    df = df[:-5]

    # Limpiar la hoja anterior (opcional)
    sheets.clear()

    # Escribir los nuevos datos
    set_with_dataframe(sheets, df)

    print("✅ Hoja 'Ohs' actualizada en Google Sheets.")

#---------------------------------------------------------------------------------------------------------------------------------#
    #reporte de OH_QuienProcesa
    print("🔗 Ingreso a informes de OH_QuienProcesa...")
    time.sleep(3)
    driver.execute_script("window.open('https://openeducation.my.salesforce.com/00O0Z0000070lP4', '_parent');")
    driver.switch_to.window(driver.window_handles[1])
    time.sleep(2)
    driver.find_element(By.XPATH, '//input[@value="Exportar detalles"]').click()
    time.sleep(3)
    dropdown = Select(driver.find_element(By.ID, "xf"))
    dropdown.select_by_value("localecsv")
    time.sleep(1)
    driver.find_element(By.XPATH, '//input[@value="Exportar"]').click()
        
    #Copiar reporte de OH_QuienProcesa
    sheetsCases = client.open_by_key("1FDok-e6IZdtTfoPbCou2St7fmiuw4eOKhzkd4unsHlc")
    sheets = sheetsCases.worksheet("OH_QuienProcesa")
    
    time.sleep(10)
    def archivo_mas_reciente(carpeta, extension=".csv"):
        archivos = [f for f in os.listdir(carpeta) if f.endswith(extension)]
        if not archivos:
            raise FileNotFoundError("No se encontró ningún archivo CSV en la carpeta.")
        
        archivos = [os.path.join(carpeta, f) for f in archivos]
        archivo_reciente = max(archivos, key=os.path.getctime)
        return archivo_reciente
    ruta_archivo = archivo_mas_reciente("C:\\Users\\yonat\\OneDrive\\Documentos\\programas\\reporte")
    print("📄 Archivo más reciente:", ruta_archivo)
    df = pd.read_csv(ruta_archivo, sep=None, engine='python')
    df = df[:-5]

    # Limpiar la hoja anterior (opcional)
    sheets.clear()

    # Escribir los nuevos datos
    set_with_dataframe(sheets, df)

    print("✅ Hoja 'OH_QuienProcesa' actualizada en Google Sheets.")
    
#---------------------------------------------------------------------------------------------------------------------------------#

    #reporte de Acc&Subs&OH   
    print("🔗 Ingreso a informes de Acc&Subs&OH...")
    time.sleep(3)
    driver.execute_script("window.open('https://openeducation.my.salesforce.com/00O3r000006vjiK','_parent');")
    driver.switch_to.window(driver.window_handles[1])
    driver.find_element(By.XPATH, '//input[@value="Exportar detalles"]').click()
    time.sleep(3)
    dropdown = Select(driver.find_element(By.ID, "xf"))
    dropdown.select_by_value("localecsv")
    time.sleep(1)
    driver.find_element(By.XPATH, '//input[@value="Exportar"]').click()
        
    #Copiar reporte de Acc&Subs&OH
    sheetsCases = client.open_by_key("1FDok-e6IZdtTfoPbCou2St7fmiuw4eOKhzkd4unsHlc")
    sheets = sheetsCases.worksheet("Acc&Subs&OH")
    
    time.sleep(10)
    def archivo_mas_reciente(carpeta, extension=".csv"):
        archivos = [f for f in os.listdir(carpeta) if f.endswith(extension)]
        if not archivos:
            raise FileNotFoundError("No se encontró ningún archivo CSV en la carpeta.")
        
        archivos = [os.path.join(carpeta, f) for f in archivos]
        archivo_reciente = max(archivos, key=os.path.getctime)
        return archivo_reciente
    ruta_archivo = archivo_mas_reciente("C:\\Users\\yonat\\OneDrive\\Documentos\\programas\\reporte")
    print("📄 Archivo más reciente:", ruta_archivo)
    df = pd.read_csv(ruta_archivo, sep=None, engine='python')
    df = df[:-5]

    # Limpiar la hoja anterior (opcional)
    sheets.clear()

    # Escribir los nuevos datos
    set_with_dataframe(sheets, df)

    print("✅ Hoja 'Acc&Subs&OH' actualizada en Google Sheets.")

#----------------------------------------------------------------------------------------------------------------------------------#

    #reporte de transfer
    spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1WaNSjdgPVJh0MaAqUxOxoO5Lu1WK9DQhhvsJQvQSJrg/edit?gid=0#gid=0")  # o client.open_by_url("URL_DEL_SHEET")

        # Selecciona la hoja (worksheet)
    worksheet = spreadsheet.worksheet("hoja 1")

        # Obtiene todos los datos
    data = worksheet.get_all_values()

    df = pd.DataFrame(data[1:], columns=data[0])  # Usa la primera fila como encabezado
    # Guarda como Excel
    df.to_excel("reporte/transfer.xlsx", index=False)
    
    #Copiar reporte de Transfer
    sheetsCases = client.open_by_key("1FDok-e6IZdtTfoPbCou2St7fmiuw4eOKhzkd4unsHlc")
    sheets = sheetsCases.worksheet("Transfer")
    
    time.sleep(10)
    def archivo_mas_reciente(carpeta, extension=".xlsx"):
        archivos = [f for f in os.listdir(carpeta) if f.endswith(extension)]
        if not archivos:
            raise FileNotFoundError("No se encontró ningún archivo CSV en la carpeta.")
        
        archivos = [os.path.join(carpeta, f) for f in archivos]
        archivo_reciente = max(archivos, key=os.path.getctime)
        return archivo_reciente
    ruta_archivo = archivo_mas_reciente("C:\\Users\\yonat\\OneDrive\\Documentos\\programas\\reporte")
    print("📄 Archivo más reciente:", ruta_archivo)
    df = pd.read_excel(ruta_archivo)
    df = df[:-5]

    # Limpiar la hoja anterior (opcional)
    sheets.clear()

    # Escribir los nuevos datos
    set_with_dataframe(sheets, df)

    print("✅ Hoja 'Transfer' actualizada en Google Sheets.")

except TimeoutException:
    print("❌ No se encontró el botón de Salesforce.")
    driver.quit()
    exit()



# Puedes dejar el navegador abierto o cerrarlo después de pruebas:
#driver.quit()
