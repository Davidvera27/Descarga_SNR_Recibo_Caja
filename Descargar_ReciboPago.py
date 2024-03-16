import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Función para verificar las condiciones de procesamiento de cada fila
def verificar_condiciones(row):
    columna_b = row[1].value
    columna_d = row[3].value
    columna_k = row[10].value
    return columna_b and not columna_k and columna_d != "No encontrado" and columna_d is not None

# Función para realizar la consulta en el sitio web y procesar los resultados
def consultar_nir(driver, nir, escritura_numero):
    try:
        # Esperar a que se cargue el campo de búsqueda NIR
        input_nir = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "filtroForm:j_idt41")))
        input_nir.clear()
        input_nir.send_keys(nir)

        # Hacer clic en el botón de búsqueda
        driver.find_element(By.ID, "filtroForm:j_idt43").click()

        # Esperar a que se carguen los resultados de la búsqueda
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.ui-g-12.ui-md-6.ui-gl-6")))
        
        # Esperar un breve momento para que se estabilice la página
        time.sleep(1)

        # Buscar elementos dentro del contenedor principal para determinar el estado del recibo
        container = driver.find_element(By.CSS_SELECTOR, "div.ui-g-12.ui-md-6.ui-gl-6")

        # Buscar elementos específicos dentro del contenedor para determinar el estado del recibo
        rechazado_texto = container.find_elements(By.XPATH, "//label[contains(text(), 'Documento en estado Ingreso Rechazado')]")
        pendiente_pago = container.find_elements(By.XPATH, "//label[contains(text(), 'Se deben pagar todos los impuestos de registro para generar el recibo de pago')]")
        pago_pendiente = container.find_elements(By.XPATH, "//label[contains(text(), 'Documento en estado Pago Pendiente')]")
        pagar_en_linea = container.find_elements(By.XPATH, "//div[contains(text(), 'PAGAR EN LÍNEA')]")
        visualizar_generar = container.find_elements(By.XPATH, "//span[@class='ui-button-text ui-c' and contains(text(), 'Visualizar y generar')]")
        pago_realizado = container.find_elements(By.XPATH, "//div[@style='font-size: 11px; font-weight: 600; color: #4bb04f' and contains(text(), 'PAGO REALIZADO')]")

        # Lógica para imprimir la respuesta de acuerdo a la presencia de los elementos
        if rechazado_texto:
            print(f"Escriitura {escritura_numero}: Proceso Rechazado. Valide correcciones y envíe nuevamente.")
        elif pendiente_pago and pago_pendiente:
            print(f"Escriitura {escritura_numero}: Aún no se ha cargado boleta de rentas para el caso.")
        elif pagar_en_linea:
            print(f"Escriitura {escritura_numero}: Recibo de pago descargado y sin cancelar")
        elif visualizar_generar:
            print(f"Escriitura {escritura_numero}: Recibo de pago listo para descargar")
        elif pago_realizado:
            print(f"Escriitura {escritura_numero}: Recibo de pago descargado y cancelado")
        else:
            print(f"Escriitura {escritura_numero}: No se puede determinar el estado del recibo")
    except TimeoutException:
        print(f"Escriitura {escritura_numero}: Tiempo de espera agotado")
    except NoSuchElementException as e:
        print(f"Escriitura {escritura_numero}: Elemento no encontrado: {str(e)}")

# Función para cerrar el navegador al presionar ENTER
def cerrar_navegador_con_enter(driver):
    input("Presiona ENTER para cerrar el navegador...")
    driver.quit()

# Configuración de Selenium
options = Options()
options.add_argument("--start-maximized")
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# Abrir la página web y hacer clic en "Pagos en línea" solo una vez
driver.get("https://radicacion.supernotariado.gov.co/app/inicio.dma")
driver.find_element(By.ID, "formLinks:paymentLink").click()
WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "filtroForm:j_idt41")))

# Cargar el archivo Excel y seleccionar la hoja principal
wb = load_workbook("C:/Users/DAVID/Desktop/DAVID/N-15/DAVID/LIBROS XLSM/HISTORICO.xlsm", data_only=True)
ws = wb["PRINCIPAL"]

# Iterar sobre las filas y procesar aquellas que cumplan con las condiciones
for row in ws.iter_rows(min_row=2, max_col=11):
    if verificar_condiciones(row):
        nir = row[3].value
        escritura_numero = row[1].value
        consultar_nir(driver, nir, escritura_numero)

# Cierra el navegador al final después de procesar todas las consultas
cerrar_navegador_con_enter(driver)
