from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import os
import datetime
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC    
import random
import shutil
import fitz  # PyMuPDF
import re
from pathlib import Path
import pdfplumber

import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter


usuario = 'mi40921'
password = 'mact2636'


DOCUMENTS_UPLOAD = 'Documentos'
DOCUMENTS_WIZARD = 'Documentos Wizard'
# Verificamos que el excel Resultados.xlsx no exista
if os.path.exists('Resultados_acuses.xlsx'):
    os.remove('Resultados_acuses.xlsx')

def dividir_pdf_usando_excel(pdf_path, excel_path, carpeta_salida='Documentos'):

    df = pd.read_excel(excel_path, sheet_name="SUGO")    
    lista_nombres = df["Folio"].dropna().astype(str).tolist()

    lector = PdfReader(pdf_path)
    total_paginas = len(lector.pages) -1

    if len(lista_nombres) != total_paginas:
        raise ValueError(f"La lista de nombres ({len(lista_nombres)}) no coincide con el número de páginas ({total_paginas}) del PDF.")

    os.makedirs(carpeta_salida, exist_ok=True)

    for i in range(total_paginas):
        escritor = PdfWriter()
        escritor.add_page(lector.pages[i+1])

        nombre_archivo = f"{lista_nombres[i].strip().replace('/', '_')}.pdf"
        ruta_salida = os.path.join(carpeta_salida, nombre_archivo)

        with open(ruta_salida, "wb") as salida:
            escritor.write(salida)
        #print(f"La hoja {i+2} corresponde al folio Sugo {lista_nombres[i]}")

    print(f"Se dividió el PDF en {total_paginas} archivos individuales en la carpeta '{carpeta_salida}'.")

    nombre_del_pdf = None
for archivo in os.listdir('.'):  
    if archivo.lower().endswith('.pdf'):
        nombre_del_pdf = archivo
        break  

if nombre_del_pdf:
    dividir_pdf_usando_excel(nombre_del_pdf, 'Oficios.xlsx')
else:
    print("No se encontró ningún archivo .pdf en el directorio.")

def estandarizar_fechas(fecha):

    from datetime import datetime
    # Mapeo de meses en español a sus equivalentes numéricos
    meses_es = {
        "ene": "01", "feb": "02", "mar": "03", "abr": "04", "may": "05", "jun": "06",
        "jul": "07", "ago": "08", "sep": "09", "oct": "10", "nov": "11", "dic": "12"
    }

    fecha = str(fecha).strip().lower()

    if pd.isnull(fecha) or fecha == 'nat' or fecha == '':
        return ''

    # Reemplazar mes en texto por número
    for mes, num in meses_es.items():
        if mes in fecha:
            fecha = fecha.replace(mes, num)
            break  # Solo reemplazamos el primer mes encontrado

    formatos = [
        "%Y-%m-%d",  # yyyy-mm-dd
        "%d-%m-%Y",  # dd-mm-yyyy
        "%d/%m/%Y",  # dd/mm/yyyy
        "%Y/%m/%d",  # yyyy/mm/dd
        "%d-%m-%y",  # dd-mm-yy
        "%d-%m-%Y",  # dd-mm-yyyy
        "%d %m %Y",
        "%m/%d/%Y",
        "%d/%m/%Y %H:%M:%S",  # dd/mm/yyyy hh:mm:ss
        "%d/%m/%y %H:%M:%S",  # dd/mm/yy hh:mm:ss
        "%d-%m-%Y %H:%M:%S",  # dd-mm-yyyy hh:mm:ss
        "%Y-%m-%d %H:%M:%S",  # yyyy-mm-dd hh:mm:ss
        "%d-%m-%y %H:%M:%S",  # dd-mm-yy hh:mm:ss
    ]

    for formato in formatos:
        try:
            fecha_obj = datetime.strptime(fecha, formato)
            return fecha_obj.strftime("%Y-%m-%d")
        except ValueError:
            continue

    return ""

print('\nEjecutando Bot para SUGO...')
# ==== 3 - Procedemos a cargarlos en SUGO

datos_df = pd.read_excel('Oficios.xlsx')

datos_df['Fecha_cierre'] = datos_df['Fecha_cierre'].apply(estandarizar_fechas)
datos_df['Fecha_cierre'] = pd.to_datetime(datos_df['Fecha_cierre'])

datos = datos_df.to_dict(orient='records')

driver = webdriver.Chrome()
wait = WebDriverWait(driver, 15)
driver.get("https://acprod.intranet.com.mx/mbom_mx_ws/mbom_mx_web/PortalLogon")

# Login
driver.find_element(By.CLASS_NAME, "name").send_keys(usuario)
driver.find_element(By.CLASS_NAME, "pass").send_keys(password)
driver.find_element(By.CLASS_NAME, "btnEntrar").click()

# Esperar a que se abra la nueva ventana
time.sleep(2)
wait.until(EC.number_of_windows_to_be(2))
driver.switch_to.window(driver.window_handles[1])
driver.get("https://acprod.intranet.com.mx:443/boixp_mx_web/boixp_mx_web/servlet/ServletOperacionWeb?OPERACION=VGOMX012&LOCALE=es_ES&DATOS_ENTRADA.FLUJO_LANZAR=GOMXFL10090")
time.sleep(2)

def cargar_acuses(dato, resultados, driver, contador, intento=1, max_intentos=3):
    if intento > max_intentos:
        print(f"{contador} - {dato['Folio']}: Falló después de {max_intentos} intentos\n")
        resultados.append([dato['Folio'], "ERROR"])
        return
    
    try:
        driver.get("https://acprod.intranet.com.mx:443/boixp_mx_web/boixp_mx_web/servlet/ServletOperacionWeb?OPERACION=VGOMX012&LOCALE=es_ES&DATOS_ENTRADA.FLUJO_LANZAR=GOMXFL10090")
        print(f"{contador} - Procesando folio: {dato['Folio']} (Intento {intento})......")
        wait.until(EC.presence_of_element_located((By.ID, "rSugo")))
        checkbox = driver.find_element(By.ID, "rSugo")
        if not checkbox.is_selected():
            checkbox.click()

        wait.until(EC.presence_of_element_located((By.ID, "fSugo")))
        input_folio = driver.find_element(By.ID, "fSugo")
        input_folio.clear()
        input_folio.send_keys(dato['Folio'])
        driver.find_element(By.ID, "busqueda").click()

        wait.until(EC.presence_of_element_located((By.ID, "radSelec0")))
        checkbox = driver.find_element(By.ID, "radSelec0")
        if not checkbox.is_selected():
            checkbox.click()

        wait.until(EC.presence_of_element_located((By.ID, "btnIMAXAcu")))
        btnAdjuntar = driver.find_element(By.ID, "btnIMAXAcu")
        btnAdjuntar.click()

        time.sleep(1)

        wait.until(lambda d: len(d.window_handles) >= 2)

        driver.switch_to.window(driver.window_handles[2])

        # Esperar a que cargue la página
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "viewerFrame")))
        #driver.get("https://imaxdigita.intranet.com.mx:7030/fimxsugo/imgmng/VisualizadorDeImagen.jsp?select=SUGO_G2454470C2D1&image.maxidx=0&image.index=-1&version=1&auditor=false&actual=true&claveDocto=null&externo=true")
        time.sleep(1)

        menu_superior = wait.until(EC.presence_of_element_located((By.ID, "DivSup")))
        tabla_superior = menu_superior.find_element(By.TAG_NAME, 'table')
        tbody_superior = tabla_superior.find_element(By.TAG_NAME, 'tbody')
        fila_sup = tbody_superior.find_element(By.TAG_NAME, 'tr')
        columnas_sup = fila_sup.find_elements(By.TAG_NAME, 'td')

        link_upload = columnas_sup[0].find_element(By.TAG_NAME, "a")
        link_upload.click()


        # Busca un input cuyo tipo sea 'file'
        time.sleep(1)
        wait.until(lambda d: len(d.window_handles) >= 3)
        driver.switch_to.window(driver.window_handles[3])
        time.sleep(0.6)
        file_input = wait.until(EC.presence_of_element_located((By.ID, "file0")))

        ruta = os.path.abspath(os.path.join(DOCUMENTS_UPLOAD, f"{dato['Folio'].strip()}.pdf"))
        file_input.send_keys(ruta)

        btn = driver.find_element(By.XPATH, "//input[@type='submit']")
        btn.click()

        # Esperar a que se complete la carga del archivo
        # si se cerro una ventana, significa que se subio el archivo
        wait.until(lambda d: len(d.window_handles) < 4)
        driver.switch_to.window(driver.window_handles[2])  # Volver a la ventana del visualizador
        driver.close()
        driver.switch_to.window(driver.window_handles[1])

        driver.get("https://acprod.intranet.com.mx/boixp_mx_web/boixp_mx_web/servlet/ServletOperacionWeb?OPERACION=VGOMX064&LOCALE=es_ES&DATOS_ENTRADA.FLUJO_LANZAR=GOMXFL15240")


        wait.until(EC.presence_of_element_located((By.ID, "porFolio")))
        checkbox = driver.find_element(By.ID, "porFolio")
        if not checkbox.is_selected():
            checkbox.click()

        wait.until(EC.presence_of_element_located((By.ID, "noFolio")))
        input_folio = driver.find_element(By.ID, "noFolio")
        input_folio.clear()
        input_folio.send_keys(dato['Folio'])
        driver.find_element(By.ID, "btnBuscaHab").click()
        
        wait.until(EC.presence_of_element_located((By.ID, "resultados")))
        tabla = driver.find_element(By.ID, "resultados")
        tbody = tabla.find_element(By.TAG_NAME, "tbody")
        filas = tbody.find_elements(By.TAG_NAME, "tr")
        columnas = filas[2].find_elements(By.TAG_NAME, "td")
        check_folio = columnas[0].find_element(By.TAG_NAME, "input")
        if not check_folio.is_selected():
            check_folio.click()
        
        fecha = dato['Fecha_cierre'].strftime('%d/%m/%Y')
        driver.execute_script("document.getElementById('fechaFin2').value = ''")
        driver.execute_script(f"document.getElementById('fechaFin2').value = '{fecha}'")

        btn_cierre = wait.until(EC.element_to_be_clickable((By.ID, "btnCierreHab")))
        btn_cierre.click()

        wait.until(EC.presence_of_element_located((By.ID, "resultados")))
        
        time.sleep(1)
        driver.get("https://acprod.intranet.com.mx:443/boixp_mx_web/boixp_mx_web/servlet/ServletOperacionWeb?OPERACION=VGOMX012&LOCALE=es_ES&DATOS_ENTRADA.FLUJO_LANZAR=GOMXFL10090")
        
        resultados.append([dato['Folio'], "Correcto"])
        print(f"{contador} - {dato['Folio']}: Cierre de acuse exitoso\n")


    except Exception as e:
        print(f"{contador} - {dato['Folio']}: Error en intento {intento} - {str(e)}")
        print(f'{contador} - {dato['Folio']}: Intentando de nuevo...')

        # Borramos todas las ventanas excepto la 0 y 1, despues volvemos a la 1
        all_ventanas = driver.window_handles
        for ventana in all_ventanas:
            if ventana != driver.window_handles[0] and ventana != driver.window_handles[1]:
                driver.switch_to.window(ventana)
                driver.close()

        driver.switch_to.window(driver.window_handles[1])
        
        # Reintentar
        time.sleep(1)  # Espera antes de reintentar
        cargar_acuses(dato, resultados, driver, contador, intento + 1, max_intentos)

contador = 0
resultados = []

time_start = time.time()

for dato in datos:
    contador += 1
    cargar_acuses(dato, resultados, driver, contador)

# Finalizar
driver.quit()

# Tiempo de ejecución
end_time = time.time()
print(f"Tiempo total de ejecución: {(end_time - time_start)/60:.2f} minutos")

# Guardar resultados en un DataFrame
resultados_df = pd.DataFrame(resultados, columns=['Folio', 'Estado'])
# Guardar resultados en un archivo Excel
resultados_df.to_excel('Resultados_acuses.xlsx', index=False)