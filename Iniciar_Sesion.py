import openpyxl
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PyQt5.QtWidgets import QMessageBox

def iniciar_sesion(usuario, contraseña):
    try:
        driver = webdriver.Chrome()
        driver.get("https://radicacion.supernotariado.gov.co/app/inicio.dma")

        usr_input = driver.find_element(By.ID, "formLogin:usrlogin")
        usr_input.send_keys(usuario)

        pwd_input = driver.find_element(By.ID, "formLogin:j_idt8")
        pwd_input.send_keys(contraseña)

        driver.find_element(By.ID, "formLogin:j_idt11").click()

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "infoBienvenida")))
        driver.find_element(By.ID, "formInfoBienvenida:id-acepta-bienvenida").click()

        # Redirigir al Visor Documental
        driver.get("https://radicacion.supernotariado.gov.co/app/external/documentary-manager.dma")

        # Esperar a que se cargue completamente la página
        time.sleep(5)

        # Mostrar mensaje de inicio de sesión exitoso
        QMessageBox.information(None, "Inicio de Sesión Exitoso", "¡Inicio de sesión exitoso!")
        
        return driver  # Devolver el controlador de Selenium

    except Exception as e:
        QMessageBox.critical(None, "Error", "Error durante el inicio de sesión: El sitio web no está disponible. Por favor, inténtelo de nuevo más tarde.")
        print("Error durante el inicio de sesión:", e)
        
def show_labels_and_textboxes(driver):
    try:
        # Mostrar las etiquetas y cuadros de texto para la consulta
        driver.execute_script("document.getElementById('lbl_nir').style.display = 'block';")
        driver.execute_script("document.getElementById('txt_nir').style.display = 'block';")
        driver.execute_script("document.getElementById('btn_buscar_nir').style.display = 'block';")
        driver.execute_script("document.getElementById('lbl_escritura').style.display = 'block';")
        driver.execute_script("document.getElementById('txt_escritura').style.display = 'block';")
        driver.execute_script("document.getElementById('btn_buscar_escritura').style.display = 'block';")
        driver.execute_script("document.getElementById('lbl_certificado').style.display = 'block';")
        driver.execute_script("document.getElementById('txt_certificado').style.display = 'block';")
        driver.execute_script("document.getElementById('btn_buscar_certificado').style.display = 'block';")

        # Simular espera mientras se procesa la autenticación
        time.sleep(5)

        # Imprimir el título de la página
        print(driver.title)

    except Exception as e:
        print("Error durante el inicio de sesión:", e)

def cerrar_sesion(driver):
    try:
        if driver is not None:
            driver.quit()
        # Mostrar mensaje de cierre de sesión exitoso
        QMessageBox.information(None, "Cierre de Sesión Exitoso", "¡Cierre de sesión exitoso!")
    except Exception as e:
        print("Error al cerrar sesión:", e)
        
def buscar_escritura(escritura, txt_nir):
    try:
        wb = openpyxl.load_workbook(r"C:\Users\DAVID\Desktop\DAVID\N-15\DAVID\LIBROS XLSM\HISTORICO.xlsm", data_only=True)
        sheet = wb["PRINCIPAL"]
        column_b_values = [str(cell.value).strip() for cell in sheet['B'] if cell.value]
        found = False
        for row_index, value in enumerate(column_b_values):
            if value == escritura:
                nir = sheet.cell(row=row_index + 1, column=4).value
                txt_nir.setText(str(nir))  # Set the NIR text
                found = True
                break

        if found:
            print("NIR encontrado:", nir)
            return nir  # Return the found NIR
        else:
            print("Número de escritura no encontrado.")
            return None  # Return None if the escritura is not found

    except Exception as e:
        print("Error al leer el archivo Excel:", e)
        return None

    finally:
        wb.close()

def realizar_busqueda_sitio_web(driver, nir):
    try:
        # Navigate to the webpage where the search will be performed
        driver.get("URL_DEL_SITIO_WEB")

        # Wait for the page to load and the NIR input field to be present
        wait = WebDriverWait(driver, 10)
        input_nir = wait.until(EC.presence_of_element_located((By.ID, "formFilterDocManager:j_idt48")))

        # Input the NIR into the text field
        input_nir.clear()
        input_nir.send_keys(nir)

        # Find and click the search button
        btn_buscar = driver.find_element(By.ID, "formFilterDocManager:j_idt79")
        btn_buscar.click()

        # Wait for the search results to load (you can add more code here as needed)

    except Exception as e:
        print("Error al interactuar con el sitio web:", e)