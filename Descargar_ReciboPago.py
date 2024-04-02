import sys
import os
from selenium.webdriver.common.keys import Keys
from PyQt5.QtCore import QTimer, QUrl
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, QPushButton
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Resultados de la Consulta")
        self.setGeometry(100, 100, 800, 600)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()
        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(3)
        self.table_widget.setHorizontalHeaderLabels(["ESCRITURA", "Estado del Recibo", "ABRIR CASO"])
        self.layout.addWidget(self.table_widget)
        self.central_widget.setLayout(self.layout)

        # Configuración de Selenium
        self.options = Options()
        self.options.add_argument("--start-maximized")
        self.service = Service(ChromeDriverManager().install())
        self.driver = None  # Inicializar el driver como None

        # Cargar estilos CSS personalizados
        self.load_custom_css()

        # Iniciar el temporizador para la consulta en segundo plano
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.consultar_casos)
        self.timer.start(1000)  # Inicia la consulta después de 1 segundo

    def load_custom_css(self):
        # Cargar el archivo CSS personalizado desde otro módulo
        css_file_path = os.path.join(os.path.dirname(__file__), "styles_Interfaz_ReciboPago.css")
        with open(css_file_path, "r") as f:
            self.setStyleSheet(f.read())

    def iniciar_sesion(self):
        # Abrir una sesión del navegador si no está abierta
        if self.driver is None:
            self.driver = webdriver.Chrome(service=self.service, options=self.options)
            self.driver.get("https://radicacion.supernotariado.gov.co/app/inicio.dma")
            # Seleccionar la opción "Pagos en línea"
            self.driver.find_element(By.ID, "formLinks:paymentLink").click()
            # Esperar a que aparezca el campo de búsqueda NIR
            WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//input[@id='filtroForm:j_idt41']")))

    def consultar_casos(self):
        try:
            # Iniciar sesión si no está iniciada
            self.iniciar_sesion()
        
            # Ruta del archivo Excel
            excel_file_path = os.path.join(os.getenv("USERPROFILE"), "Desktop", "DAVID", "N-15", "DAVID", "LIBROS XLSM", "HISTORICO.xlsm")
            # Cargar el archivo Excel y seleccionar la hoja principal
            wb = load_workbook(excel_file_path, data_only=True)
            ws = wb["PRINCIPAL"]

            # Consultar casos
            for row in ws.iter_rows(min_row=2, max_col=11):
                if verificar_condiciones(row):
                    nir = row[3].value
                    escritura_numero = row[1].value
                    estado_recibo = consultar_nir(self.driver, nir, escritura_numero)
                    self.mostrar_resultado(escritura_numero, estado_recibo, nir)

            # Detener el temporizador una vez que todas las filas se han procesado
            self.timer.stop()

        except Exception as e:
            print(f"Error al consultar casos: {str(e)}")


    def mostrar_resultado(self, escritura_numero, estado_recibo, nir):
        row_position = self.table_widget.rowCount()
        self.table_widget.insertRow(row_position)
        self.table_widget.setItem(row_position, 0, QTableWidgetItem(str(escritura_numero)))
        self.table_widget.setItem(row_position, 1, QTableWidgetItem(estado_recibo))
        
        abrir_button = QPushButton("ABRIR", self)
        abrir_button.clicked.connect(lambda _, nir_value=nir: self.abrir_caso(nir_value))
        self.table_widget.setCellWidget(row_position, 2, abrir_button)

    def abrir_caso(self, nir):
        try:
            # Iniciar sesión si no está iniciada
            self.iniciar_sesion()
            
            # Ingresar el NIR en el campo de texto
            input_nir = self.driver.find_element(By.XPATH, "//input[@id='filtroForm:j_idt41']")
            input_nir.clear()
            input_nir.send_keys(nir, Keys.RETURN)  # Presionar Enter para buscar el NIR

            # Esperar a que se carguen los resultados
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.ui-g-12.ui-md-6.ui-gl-6")))

        except Exception as e:
            print(f"Error al abrir el caso: {str(e)}")

def verificar_condiciones(row):
    columna_d = row[3].value  # Columna D (columna 4)
    columna_k = row[10].value  # Columna K (columna 11)
    
    # Verificar si la columna K está vacía y la columna D no es "No encontrado" ni nula
    if not columna_k and columna_d != "No encontrado" and columna_d is not None:
        return True
    else:
        return False


def consultar_nir(driver, nir, escritura_numero):
    try:
        input_nir = driver.find_element(By.XPATH, "//input[@id='filtroForm:j_idt41']")
        input_nir.clear()
        input_nir.send_keys(nir)
        driver.find_element(By.XPATH, "//button[@id='filtroForm:j_idt43']").click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.ui-g-12.ui-md-6.ui-gl-6")))

        container = driver.find_element(By.CSS_SELECTOR, "div.ui-g-12.ui-md-6.ui-gl-6")
        rechazado_texto = container.find_elements(By.XPATH, "//label[contains(text(), 'Documento en estado Ingreso Rechazado')]")
        
        if rechazado_texto:
            return f"Proceso Rechazado. Valide correcciones y envíe nuevamente."
        elif container.find_elements(By.XPATH, "//label[contains(text(), 'Se deben pagar todos los impuestos de registro para generar el recibo de pago')]") and container.find_elements(By.XPATH, "//label[contains(text(), 'Documento en estado Pago Pendiente')]"):
            return f"Aún no se ha cargado boleta de rentas para el caso."
        elif container.find_elements(By.XPATH, "//label[contains(text(), 'Se deben pagar todos los impuestos de registro para generar el recibo de pago')]") and container.find_elements(By.XPATH, "//label[contains(text(), 'Documento en estado Aprobación Pendiente')]"):
            return f"Caso en estado de aprobación PENDIENTE. Se debe cargar boleta de rentas al caso"
        elif container.find_elements(By.XPATH, "//label[contains(text(), 'Se deben pagar todos los impuestos de registro para generar el recibo de pago')]") and container.find_elements(By.XPATH, "//label[contains(text(), 'Documento en estado Aprobación N/A')]"):
            return f"Caso en estado de aprobación 'N/A'. Se debe cargar boleta de rentas al caso"
        elif container.find_elements(By.XPATH, "//div[contains(text(), 'PAGAR EN LÍNEA')]"):
            return f"Recibo de pago descargado y sin cancelar"
        elif container.find_elements(By.XPATH, "//span[@class='ui-button-text ui-c' and contains(text(), 'Visualizar y generar')]"):
            return f"Recibo de pago listo para descargar"
        elif container.find_elements(By.XPATH, "//div[@style='font-size: 11px; font-weight: 600; color: #4bb04f' and contains(text(), 'PAGO REALIZADO')]"):
            return f"Recibo de pago descargado y cancelado"
        else:
            return f"No se puede determinar el estado del recibo"
    except Exception as e:
        return f"Error durante la consulta: {str(e)}"

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
