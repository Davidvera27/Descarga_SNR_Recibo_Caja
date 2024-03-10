import openpyxl
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import os
import time

def consultar_anexos():
    # Configurar el servicio del controlador de Chrome
    service = Service('PYTHON\\WebScraping\\chromedriver-win64\\chromedriver.exe')

    # Configurar las opciones del navegador
    options = Options()
    options.add_argument('--start-maximized')  # Para iniciar maximizado

    # Inicializar el navegador
    driver = WebDriver(service=service, options=options)

    try:
        # Navegar a la página de inicio de sesión
        driver.get("https://radicacion.supernotariado.gov.co/app/inicio.dma")

        # Diligenciar el formulario de inicio de sesión
        usuario = "JULIAN.ZAPATA"
        contraseña = "Notaria15"

        usuario_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "formLogin:usrlogin")))
        usuario_field.send_keys(usuario)
        
        contraseña_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "formLogin:j_idt8")))
        contraseña_field.send_keys(contraseña)
        
        # Hacer clic en el botón "Autenticarse"
        iniciar_sesion_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "formLogin:j_idt11")))
        iniciar_sesion_btn.click()

        # Esperar un momento para que la página se cargue completamente
        driver.implicitly_wait(5)

        # Cargar el archivo Excel
        archivo_excel = "C:/Users/DAVID/Desktop/DAVID/N-15/DAVID/LIBROS XLSM/HISTORICO.xlsm"
        wb = openpyxl.load_workbook(archivo_excel, data_only=True)
        sheet = wb["PRINCIPAL"]

        # Iterar sobre las filas
        for row in sheet.iter_rows(min_row=2, max_col=11, values_only=False):
            if row[10].value == "DESCARGADA":
                nir = row[3].value  # Valor de la columna "D"
                if nir is not None:  # Verificar si el valor no es None
                    # Navegar a la página de búsqueda
                    driver.get("https://radicacion.supernotariado.gov.co/app/external/documentary-manager.dma")

                    # Esperar un momento para que la página de búsqueda se cargue completamente
                    driver.implicitly_wait(5)

                    # Enviar el NIR al formulario de búsqueda
                    nir_field = driver.find_element(By.ID, "formFilterDocManager:j_idt48")
                    nir_field.clear()
                    nir_field.send_keys(nir)

                    # Hacer clic en el botón de búsqueda
                    buscar_btn = driver.find_element(By.ID, "formFilterDocManager:j_idt79")
                    buscar_btn.click()

                    # Esperar a que se cargen los resultados de la búsqueda
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "formDocsManager")))

                    # Verificar si el elemento específico está presente en la página
                    try:
                        recibo_caja_link = driver.find_element(By.ID, "formDocsManager:j_idt84:0:j_idt158")
                        recibo_caja_link.click()  # Hacer clic en el enlace del recibo de caja

                        # Esperar a que se cargue el recibo de caja
                        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "formDocsManager:j_idt197")))

                        # Hacer clic en el botón de descarga
                        descargar_btn = driver.find_element(By.ID, "formDocsManager:j_idt197")
                        descargar_btn.click()

                        # Esperar hasta que el archivo se descargue completamente
                        tiempo_inicial = time.time()
                        while True:
                            if time.time() - tiempo_inicial > 60:
                                print("Tiempo de espera para la descarga excedido.")
                                break
                            archivos_en_descargas = [f for f in os.listdir("E:/Downloads") if f.endswith('.pdf')]
                            if archivos_en_descargas:
                                break
                            time.sleep(1)

                        # Obtener el nombre del último archivo descargado
                        if archivos_en_descargas:
                            nombre_archivo = archivos_en_descargas[-1]
                            nuevo_nombre = f"BR{row[1].value}.pdf"
                            ruta_archivo_original = os.path.join("E:/Downloads", nombre_archivo)
                            ruta_archivo_nuevo = os.path.join("E:/Downloads", nuevo_nombre)
                            
                            # Esperar hasta que el archivo esté completamente descargado
                            tiempo_inicial = time.time()
                            while time.time() - tiempo_inicial < 60:
                                if os.path.exists(ruta_archivo_original):
                                    # Renombrar el archivo
                                    os.rename(ruta_archivo_original, ruta_archivo_nuevo)
                                    print(f"Archivo PDF descargado y guardado correctamente como {nuevo_nombre}")
                                    break
                                time.sleep(1)
                            else:
                                print("El archivo no se ha descargado correctamente.")
                    except NoSuchElementException:
                        print(f"Escritura {row[1].value} = 'Recibo de caja Pendiente'")
                else:
                    print(f"La escritura {row[1].value} no contiene NIR")
    finally:
        # Cerrar el navegador al finalizar
        driver.quit()

# Llamar a la función para iniciar la consulta
consultar_anexos()