from playwright.sync_api import sync_playwright
from datetime import datetime, timedelta
import os
import pandas as pd
import xlwings as xw
import requests
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import time
import pygetwindow as gw

# --------- CONFIGURACIÓN ---------
download_dir = os.path.abspath(r"G:\\Unidades compartidas\\Workforce\\r\\descargas\\Recovery Brazil")
os.makedirs(download_dir, exist_ok=True)

yesterday = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")

report_links = [
    "https://openeducation.my.salesforce.com/00OKV000008PVfQ",  # Casos Casos
    "1NvE6QAS2nl4M2tGtJ-6h5fNHPFrF_RW0-QnYGVFXpMg"              # Transfer (Google Sheet ID)
]

sheet_names = ["Casos", "Transfer"]

SCOPES = ['https://www.googleapis.com/auth/drive']
SERVICE_ACCOUNT_FILE = 'G:\\Unidades compartidas\\Workforce\\r\\clave_api\\auto.json'

creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=creds)

excel_path = 'G:\\Unidades compartidas\\Workforce\\r\\Recovery_Monthly_Brasil.xlsx'

ruta_base = "G:\\Unidades compartidas\\Workforce\\r\\registros"
fecha_hoy = datetime.now().strftime("%Y-%m-%d")

# --------- AUTOMATIZACIÓN ---------

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)  # Cambiado a headless=True
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()
    
    # Espera a que la ventana aparezca
    time.sleep(2)  # ajusta según lo que tarde en abrir
    for w in gw.getWindowsWithTitle('Chromium'):  # o usa 'Edge' si estás con Microsoft Edge
        w.minimize()

    # Login en Okta
    page.goto("https://openeducation.okta.com")
    page.fill('input[name="identifier"]', "yonatan.osorio@openenglish.com")
    page.fill('input[name="credentials.passcode"]', "Matias2023#")
    page.click('input[value="Iniciar sesión"]')
    page.wait_for_timeout(5000)

    # Abrimos Excel una sola vez
    app = xw.App(visible=False)
    wb = app.books.open(excel_path)
    print("Ruta del archivo Excel:", excel_path)

    for link, sheet_name in zip(report_links, sheet_names):
        print(f"📥 Procesando reporte: {sheet_name}")

        if sheet_name == "Transfer":
            try:
                export_url = f"https://docs.google.com/spreadsheets/d/{link}/export?format=xlsx"
                headers = {"Authorization": f"Bearer {creds.token}"}
                response = requests.get(export_url, headers=headers)
                response.raise_for_status()
                file_path = os.path.join(download_dir, "Transfer.xlsx")
                with open(file_path, "wb") as f:
                    f.write(response.content)
                print(f"✅ Archivo Transfer descargado: {file_path}")

                df = pd.read_excel(file_path)

                # Convertir columnas de fecha
                for col in df.columns:
                    if "fecha" in col.lower() or "date" in col.lower():
                        try:
                            df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')
                        except Exception as e:
                            print(f"⚠️ No se pudo convertir la columna '{col}' a fecha: {e}")

                ws = wb.sheets[sheet_name]
                ws.range('A2').expand('table').clear_contents()  # Limpiar solo el rango
                ws.range('A2').value = df.values

                print(f"✅ Datos pegados correctamente en la hoja {sheet_name}")
            except Exception as e:
                print(f"❌ Error descargando Transfer: {e}")
            continue

        # Salesforce Reports
        page.goto(link)
        page.wait_for_load_state("networkidle")

        if sheet_name == "Casos":
            try:
                page.click("input[id='edate']")
                page.fill("input[id='edate']", yesterday)
                page.click('input[value="Ejecutar informe"]')
                page.wait_for_timeout(3000)
            except Exception as e:
                print(f"⚠️ Error al aplicar filtro de fecha para {sheet_name}: {e}")

        try:
            with page.expect_download() as download_info:
                page.click('input[value="Exportar detalles"]')
                page.select_option("#xf", "localecsv")
                page.click('input[value="Exportar"]')
            download = download_info.value
            download_path = os.path.join(download_dir, f"{sheet_name}.csv")
            download.save_as(download_path)
            print(f"✅ Archivo descargado: {download_path}")
        except Exception as e:
            print(f"❌ Error descargando {sheet_name}: {e}")
            continue

        try:
            df = pd.read_csv(download_path, sep=';', on_bad_lines='skip')
            df = df.iloc[:-5]  # Omitir últimas 5 filas

            # Convertir columnas de fecha
            for col in df.columns:
                if "fecha" in col.lower() or "date" in col.lower():
                    try:
                        df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')
                    except Exception as e:
                        print(f"⚠️ No se pudo convertir la columna '{col}' a fecha: {e}")

            ws = wb.sheets[sheet_name]

            # Sobrescribir solo el rango necesario sin tocar fórmulas
            last_row = len(df) + 1
            ws.range(f'A2:A{last_row}').value = df.values

            print(f"✅ Datos pegados correctamente en la hoja {sheet_name}")
        except Exception as e:
            print(f"❌ Error procesando Excel para {sheet_name}: {e}")

    # --------- ACTUALIZAR TODO EL ARCHIVO ---------

    try:
        print("🔄 Refrescando consultas, conexiones y pivots...")
        wb.api.RefreshAll()
        time.sleep(100)

        wb.app.calculate()

        wb.save()
        wb.close()
        app.quit()
        
        nombre_archivo = f"Monthly Recovery Brazil - log_{fecha_hoy}.txt"
        ruta_completa = os.path.join(ruta_base, nombre_archivo)

        with open(ruta_completa, "a", encoding="utf-8") as f:
            f.write(f"✔ Monthly Recovery Brazil actualizado actualizado el {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

        print(f"✅ Log generado: {nombre_archivo}")
        print("✅ Archivo Excel actualizado correctamente.")
    except Exception as e:
        print(f"❌ Error al refrescar el archivo Excel: {e}")

    print("🏁 Proceso finalizado.")
