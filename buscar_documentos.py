from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def buscar_documentos(driver, nir):
    # Navegar a la página de búsqueda
    driver.get("https://radicacion.supernotariado.gov.co/app/external/documentary-manager.dma")

    # Esperar un momento para que la página de búsqueda se cargue completamente
    driver.implicitly_wait(3)

    # Enviar el NIR al formulario de búsqueda
    nir_field = driver.find_element(By.ID, "formFilterDocManager:j_idt48")
    nir_field.clear()
    nir_field.send_keys(nir)

    # Hacer clic en el botón de búsqueda
    buscar_btn = driver.find_element(By.ID, "formFilterDocManager:j_idt79")
    buscar_btn.click()

    # Esperar a que se cargen los resultados de la búsqueda
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "formDocsManager")))
