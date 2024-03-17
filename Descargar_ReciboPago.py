import logging
import re
import time
import webbrowser
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem, QPushButton, QSizePolicy, QHeaderView


# Configuración del logger
logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)

# Definir variables globales
SEARCH_BOX_ID = "filtroForm:j_idt41"
SEARCH_BUTTON_ID = "filtroForm:j_idt43"
MAIN_CONTAINER_CSS = "div.ui-g-12.ui-md-6.ui-gl-6"

# Estructura de datos para almacenar resultados de consultas
results_data = []

# Función para verificar las condiciones de procesamiento de cada fila
def verificar_condiciones(row):
    columna_b = row[1].value
    columna_d = row[3].value
    columna_k = row[10].value
    return columna_b and not columna_k and columna_d != "No encontrado" and columna_d is not None

# Función para realizar la consulta en el sitio web y procesar los resultados
def consultar_nir(driver, nir, escritura_numero, timeout=20):
    try:
        # Esperar a que se cargue el campo de búsqueda NIR
        input_nir = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.ID, SEARCH_BOX_ID)))
        input_nir.clear()
        input_nir.send_keys(nir)

        # Hacer clic en el botón de búsqueda
        driver.find_element(By.ID, SEARCH_BUTTON_ID).click()

        # Esperar a que se carguen los resultados de la búsqueda
        WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, MAIN_CONTAINER_CSS)))
        
        # Esperar un breve momento para que se estabilice la página
        time.sleep(1)

        # Buscar elementos dentro del contenedor principal para determinar el estado del recibo
        container = driver.find_element(By.CSS_SELECTOR, MAIN_CONTAINER_CSS)

        # Buscar elementos específicos dentro del contenedor para determinar el estado del recibo
        rechazado_texto = container.find_elements(By.XPATH, "//label[contains(text(), 'Documento en estado Ingreso Rechazado')]")
        pendiente_pago = container.find_elements(By.XPATH, "//label[contains(text(), 'Se deben pagar todos los impuestos de registro para generar el recibo de pago')]")
        pago_pendiente = container.find_elements(By.XPATH, "//label[contains(text(), 'Documento en estado Pago Pendiente')]")
        pagar_en_linea = container.find_elements(By.XPATH, "//div[contains(text(), 'PAGAR EN LÍNEA')]")
        visualizar_generar = container.find_elements(By.XPATH, "//span[@class='ui-button-text ui-c' and contains(text(), 'Visualizar y generar')]")
        pago_realizado = container.find_elements(By.XPATH, "//div[@style='font-size: 11px; font-weight: 600; color: #4bb04f' and contains(text(), 'PAGO REALIZADO')]")

        # Lógica para almacenar los resultados en la estructura de datos y mostrarlos en la ventana emergente
        if rechazado_texto:
            result = f"Escriitura {escritura_numero}: Proceso Rechazado. Valide correcciones y envíe nuevamente."
        elif pendiente_pago and pago_pendiente:
            result = f"Escriitura {escritura_numero}: Aún no se ha cargado boleta de rentas para el caso."
        elif pagar_en_linea:
            result = f"Escriitura {escritura_numero}: Recibo de pago descargado y sin cancelar"
        elif visualizar_generar:
            result = f"Escriitura {escritura_numero}: Recibo de pago listo para descargar"
        elif pago_realizado:
            result = f"Escriitura {escritura_numero}: Recibo de pago descargado y cancelado"
        else:
            result = f"Escriitura {escritura_numero}: No se puede determinar el estado del recibo"
        
        results_data.append(result)
    except TimeoutException:
        results_data.append(f"Escriitura {escritura_numero}: Tiempo de espera agotado")
    except NoSuchElementException as e:
        results_data.append(f"Escriitura {escritura_numero}: Elemento no encontrado: {str(e)}")

# Configuración de Selenium
options = Options()
options.add_argument("--headless")  # Ejecutar en modo headless
options.add_argument("--start-maximized")
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# Abrir la página web y hacer clic en "Pagos en línea" solo una vez
driver.get("https://radicacion.supernotariado.gov.co/app/inicio.dma")
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "formLinks:paymentLink"))).click()
WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, SEARCH_BOX_ID)))

# Cargar el archivo Excel y seleccionar la hoja principal
wb = load_workbook("C:/Users/DAVID/Desktop/DAVID/N-15/DAVID/LIBROS XLSM/HISTORICO.xlsm", data_only=True)
ws = wb["PRINCIPAL"]

# Iterar sobre las filas y procesar aquellas que cumplan con las condiciones
for row in ws.iter_rows(min_row=2, max_col=11):
    if verificar_condiciones(row):
        nir = row[3].value
        escritura_numero = row[1].value
        consultar_nir(driver, nir, escritura_numero)

# Crear ventana emergente con PyQt
class ResultsWindow(QWidget):
    def __init__(self, results):
        super().__init__()
        self.results = results
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Resultados de Consulta')
        self.setGeometry(100, 100, 800, 400)

        layout = QVBoxLayout()

        tableWidget = QTableWidget()
        tableWidget.setRowCount(len(self.results))
        tableWidget.setColumnCount(2)  # Se agrega una columna adicional para el botón ejecutable
        tableWidget.setHorizontalHeaderLabels(["Resultado", "Acción"])

        for i, result in enumerate(self.results):
            item = QTableWidgetItem(result)
            tableWidget.setItem(i, 0, item)
            
            # Creamos el botón en la columna de acción
            button = QPushButton("Abrir Consulta")
            button.clicked.connect(lambda _, row=i: self.open_query(row))
            tableWidget.setCellWidget(i, 1, button)

        # Ajustar automáticamente la altura de la ventana según el contenido
        tableWidget.setSizeAdjustPolicy(QTableWidget.AdjustToContents)

        # Ajustar automáticamente el tamaño de la columna al contenido
        tableWidget.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)

        layout.addWidget(tableWidget)

        self.setLayout(layout)
        self.show()
    
    def open_query(self, row):
        # En esta función se abrirá la consulta correspondiente al caso en la fila seleccionada
        # Obtenemos la información necesaria del resultado
        resultado = self.results[row]
        escritura_numero = resultado.split(":")[0].split(" ")[1]
        nir = resultado.split(":")[0].split(" ")[1]  # Suponiendo que el nir y la escritura_numero son lo mismo
        
        # Realizamos la consulta en el sitio web
        consultar_nir(driver, nir, escritura_numero)
        
        # Abrimos el caso en el navegador web
        consulta_url = f"https://radicacion.supernotariado.gov.co/app/inicio.dma?nir={nir}"  # URL de la consulta con el NIR
        webbrowser.open(consulta_url)

if __name__ == '__main__':
    app = QApplication([])
    results_window = ResultsWindow(results_data)
    app.exec_()
