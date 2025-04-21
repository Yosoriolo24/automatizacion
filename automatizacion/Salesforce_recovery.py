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
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from PIL import ImageGrab

# --------- CONFIGURACIÓN ---------
download_dir = os.path.abspath(r"G:\\Unidades compartidas\\Workforce\\r\\descargas\\Recovery")
os.makedirs(download_dir, exist_ok=True)

yesterday = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")

report_links = [
    "https://openeducation.my.salesforce.com/00O3r000006mGPy",  # Casos
    "https://openeducation.my.salesforce.com/00O0Z0000070nf9",  # Ohs
    "https://openeducation.my.salesforce.com/00O0Z0000070lP4",  # OH_QuienProcesa
    "https://openeducation.my.salesforce.com/00O3r000006vjiK",  # Acc&Subs&OH
    "1ghpEpBu_-6UyCwja356CQ4lXiKk9XsWdUdGGd2ayPek"              # Transfer (Google Sheet ID)
]

sheet_names = ["Casos", "Ohs", "OH_QuienProcesa", "Acc&Subs&OH", "Transfer"]

SCOPES = ['https://www.googleapis.com/auth/drive']
SERVICE_ACCOUNT_FILE = 'G:\\Unidades compartidas\\Workforce\\r\\clave_api\\auto.json'

creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=creds)

excel_path = 'G:\\Unidades compartidas\\Workforce\\r\\1321321 Recovery_Monthly.xlsx'

ruta_base = "G:\\Unidades compartidas\\Workforce\\r\\registros"
fecha_hoy = datetime.now().strftime("%Y-%m-%d")

sheets_picture = ['Front', 'Front Fallidas', 'Front BNPL']
ranges_picture = ['A1:Y152', 'A1:Y97', 'A1:Y68']
temp_folder = 'G:\\Unidades compartidas\\Workforce\\r\\descargas\\Recovery\\Capturas\\'

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
            
    # --------- CAPTURAR IMÁGENES DE LAS HOJAS ---------
    temp_folder = 'G:\\Unidades compartidas\\Workforce\\r\\descargas\\Recovery\\Capturas\\'
    os.makedirs(temp_folder, exist_ok=True)

    for sheet_name, cell_range in zip(sheets_picture, ranges_picture):
        ws = wb.sheets[sheet_name]
        pic = ws.range(cell_range).api.CopyPicture(Format=1)  # Copiar como imagen
        image = ws.api.Paste()  # Pegar imagen en la hoja
        image = ws.pictures[-1]
        image_path = os.path.join(temp_folder, f"{sheet_name}.png")
        image.api.Copy()
        image.api.Delete()

        # Guardar imagen desde el portapapeles
        img = ImageGrab.grabclipboard()
        if img:
            img.save(image_path)
            print(f"✅ Imagen guardada: {image_path}")
        else:
            print(f"⚠️ No se pudo capturar imagen para {sheet_name}")

    # --------- ACTUALIZAR TODO EL ARCHIVO ---------

    try:
        print("🔄 Refrescando consultas, conexiones y pivots...")
        wb.api.RefreshAll()
        time.sleep(60)

        wb.app.calculate()

        wb.save()
        wb.close()
        app.quit()
        
        nombre_archivo = f"Monthly Recovery - log_{fecha_hoy}.txt"
        ruta_completa = os.path.join(ruta_base, nombre_archivo)

        with open(ruta_completa, "a", encoding="utf-8") as f:
            f.write(f"✔ Monthly Recovery actualizado actualizado el {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

        print(f"✅ Log generado: {nombre_archivo}")
        print("✅ Archivo Excel actualizado correctamente.")
    except Exception as e:
        print(f"❌ Error al refrescar el archivo Excel: {e}")
        
        
    try:
        # Detalles del servidor SMTP (Gmail)
        servidor = smtplib.SMTP('smtp.gmail.com', 587)
        servidor.starttls()

            # Información del correo
        remitente = "yonatan.osorio@openenglish.com"
        destinatario = "yonatan.osorio@openenglish.com"
        contraseña_aplicacion = "kabz icgm xjyc wofo"  # Contraseña de aplicación generada

            # Conectarse y autenticar
        servidor.login(remitente, contraseña_aplicacion)

            # Crear el mensaje
        mensaje = MIMEMultipart()
        mensaje['From'] = remitente
        mensaje['To'] = destinatario
        mensaje['Subject'] = "Informe de Recovery"

            # Cuerpo del mensaje
        cuerpo = "Aquí van los reportes y las imágenes solicitadas."
        mensaje.attach(MIMEText(cuerpo, 'plain'))

        # Adjuntar imágenes de las hojas
        for sheet_name in sheets_picture:
            image_path = os.path.join(temp_folder, f"{sheet_name}.png")
            part = MIMEBase('application', 'octet-stream')
            with open(image_path, 'rb') as f:
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={sheet_name}.png')
            mensaje.attach(part)

        # Enviar el correo
        servidor.sendmail(remitente, destinatario, mensaje.as_string())
        servidor.quit()

        print("✅ Correo enviado correctamente con las imágenes adjuntas.")

    except Exception as e:
        print(f"❌ Error enviando el correo: {e}")

    print("🏁 Proceso finalizado.")
