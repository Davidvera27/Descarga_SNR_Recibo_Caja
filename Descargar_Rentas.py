import sys
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtGui import QIcon

class ConsultaAnexosInterfaz(QtWidgets.QWidget):
    def __init__(self, excel_file_path):
        super().__init__()
        self.setWindowTitle("Consulta de Anexos")
        self.setGeometry(100, 100, 800, 400)
        self.excel_file_path = excel_file_path
        self.proceso_en_curso = False
        self.icono_busqueda = QIcon("PYTHON\WebScraping\Excel\Descarga_SNR_Recibo_Caja\Icons\Search_icon.png")
        self.icono_ver = QIcon("PYTHON\WebScraping\Excel\Descarga_SNR_Recibo_Caja\Icons\ver_icon.png")

        # Estilos
        self.setStyleSheet("background-color: #f0f0f0;")

        # Tabla de anexos
        self.tabla_anexos = QtWidgets.QTableWidget(self)
        self.tabla_anexos.setGeometry(50, 50, 700, 300)
        self.tabla_anexos.setColumnCount(4)
        self.tabla_anexos.setHorizontalHeaderLabels(["ESCRITURA", "RADICADO", "ESTADO DE LIQUIDACIÓN", "VER"])
        self.tabla_anexos.verticalHeader().setVisible(False)
        self.tabla_anexos.horizontalHeader().setStyleSheet("QHeaderView::section { background-color: #d3d3d3; }")
        self.tabla_anexos.setStyleSheet("color: #333; font-size: 12px;")
        self.tabla_anexos.horizontalHeader().setStretchLastSection(True)

        # Configurar el botón de consulta
        self.boton_consultar = QtWidgets.QPushButton("ABRIR LIQUIDACIÓN DE RENTAS", self)
        self.boton_consultar.setGeometry(50, 360, 250, 30)
        self.boton_consultar.setIcon(self.icono_busqueda)  # Asignar el ícono de búsqueda al botón
        self.boton_consultar.setIconSize(QtCore.QSize(24, 24))  # Ajustar el tamaño del ícono
        self.boton_consultar.setStyleSheet("QPushButton { background-color: #ADD8E6; color: white; border: 1px solid #4682B4; border-radius: 5px; }"
                                           "QPushButton:hover { background-color: #87CEEB; }"
                                           "QPushButton:pressed { background-color: #4682B4; }")  # Establecer estilo CSS
        self.boton_consultar.clicked.connect(self.consultar_anexos)

        # Configurar el botón de reiniciar
        self.boton_reiniciar = QtWidgets.QPushButton("REINICIAR", self)
        self.boton_reiniciar.setGeometry(320, 360, 250, 30)
        self.boton_reiniciar.setIcon(self.icono_busqueda)  # Asignar el ícono de búsqueda al botón
        self.boton_reiniciar.setIconSize(QtCore.QSize(24, 24))  # Ajustar el tamaño del ícono
        self.boton_reiniciar.setStyleSheet("QPushButton { background-color: #ADD8E6; color: white; border: 1px solid #4682B4; border-radius: 5px; }"
                                            "QPushButton:hover { background-color: #87CEEB; }"
                                            "QPushButton:pressed { background-color: #4682B4; }")  # Establecer estilo CSS
        self.boton_reiniciar.clicked.connect(self.reiniciar_proceso)

        # Otras configuraciones
        self.chrome_options = Options()
        self.chrome_options.add_argument("--headless")
        self.driver = webdriver.Chrome(options=self.chrome_options)

        self.cargar_iconos()

    def cargar_iconos(self):
        # Aquí cargarías los iconos desde tus archivos locales o descargarías desde internet
        # Ejemplo de cómo se cargarían los iconos
        self.icono_ver = QIcon("ver.png")

    def ver_anexo(self):
        boton = self.sender()
        fila = self.tabla_anexos.indexAt(boton.pos()).row()
        radicado = self.tabla_anexos.item(fila, 1).text()
        proceso = self.tabla_anexos.item(fila, 2).text()
        url_consulta = f"https://mercurio.antioquia.gov.co/mercurio/servlet/ControllerMercurio?command=anexos&tipoOperacion=abrirLista&idDocumento={radicado}&tipDocumento=R&now=Date()&ventanaEmergente=S&origen=NTR&proceso={proceso}"
        
        # Abrir el navegador controlado por Selenium en modo visible para mostrar los anexos
        self.driver = webdriver.Chrome()
        self.driver.get(url_consulta)

    def consultar_anexos(self):
        radicados_a_consultar = self.buscar_nir_para_consulta(self.excel_file_path)
        self.proceso_en_curso = True

        for radicado, escritura_procesada in radicados_a_consultar:
            try:
                estado_liquidacion, enlaces_anexos = self.obtener_estado_y_enlaces_anexos(radicado)
                if enlaces_anexos:
                    # Agregar fila a la tabla visible en la interfaz gráfica
                    self.agregar_fila_tabla_anexos(escritura_procesada, radicado, estado_liquidacion)
            except Exception as e:
                print(f"Error al consultar radicado {radicado}: {e}")

        print("Proceso de consulta de anexos finalizado.")
        self.proceso_en_curso = False

    def buscar_nir_para_consulta(self, excel_file_path):
        wb = openpyxl.load_workbook(excel_file_path, data_only=True)
        sheet = wb["ESCRITURAS FISICAS (DIARIO)"]
        return [(row[2], row[1]) for row in sheet.iter_rows(min_row=2, max_col=4, max_row=sheet.max_row, values_only=True) if row[3] == "SIN INICIAR" and row[1]]

    def obtener_estado_y_enlaces_anexos(self, radicado):
        url_consulta = f"https://mercurio.antioquia.gov.co/mercurio/servlet/ControllerMercurio?command=anexos&tipoOperacion=abrirLista&idDocumento={radicado}&tipDocumento=R&now=Date()&ventanaEmergente=S&origen=NTR"
        self.driver.get(url_consulta)
        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//td[contains(text(), 'Total Anexos')]")))
        enlaces_anexos = self.driver.find_elements(By.XPATH, "//tr[@class='contenido']/td/a")
        cantidad_anexos = len(enlaces_anexos) if enlaces_anexos else 0
        estado_liquidacion = f"LIQUIDACIÓN LISTA ({cantidad_anexos})" if cantidad_anexos else "SIN ANEXOS"
        return estado_liquidacion, enlaces_anexos

    def agregar_fila_tabla_anexos(self, escritura_procesada, radicado, estado_liquidacion):
        fila = self.tabla_anexos.rowCount()
        self.tabla_anexos.insertRow(fila)
        self.tabla_anexos.setItem(fila, 0, QtWidgets.QTableWidgetItem(str(escritura_procesada)))
        self.tabla_anexos.setItem(fila, 1, QtWidgets.QTableWidgetItem(str(radicado)))
        self.tabla_anexos.setItem(fila, 2, QtWidgets.QTableWidgetItem(estado_liquidacion))

        # Configurar el botón VER en la columna 3
        boton_ver = QtWidgets.QPushButton(self)
        boton_ver.setText("VER")  # Establecer el texto del botón
        boton_ver.setIcon(self.icono_ver)  # Asignar el ícono VER al botón
        boton_ver.setIconSize(QtCore.QSize(24, 24))  # Ajustar el tamaño del ícono
        boton_ver.setStyleSheet("QPushButton { text-align: left; padding-left: 5px; padding-right: 5px; background-color: #ADD8E6; border: 1px solid #4682B4; border-radius: 5px; }"
                                "QPushButton:hover { background-color: #87CEEB; }"
                                "QPushButton:pressed { background-color: #4682B4; }")  # Establecer estilo CSS
        boton_ver.clicked.connect(self.ver_anexo)
        self.tabla_anexos.setCellWidget(fila, 3, boton_ver)  # Asignar el botón a la celda de la tabla

    def reiniciar_proceso(self):
        # Limpiar la tabla
        self.tabla_anexos.clearContents()
        self.tabla_anexos.setRowCount(0)
        # Volver a consultar anexos
        self.consultar_anexos()

if __name__ == "__main__":
    excel_file_path = r'C:\Users\DAVID\Desktop\DAVID\N-15\DAVID\LIBROS XLSM\HISTORICO.xlsm'
    app = QtWidgets.QApplication(sys.argv)
    ventana = ConsultaAnexosInterfaz(excel_file_path)
    ventana.show()
    sys.exit(app.exec_())
