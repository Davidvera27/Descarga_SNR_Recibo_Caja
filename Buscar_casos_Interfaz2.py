import openpyxl
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

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
