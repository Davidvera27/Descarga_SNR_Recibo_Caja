import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem, QPushButton, QMessageBox
from PyQt5.QtGui import QIcon


class ResultadosVentana(QWidget):
    def __init__(self, escrituras, radicados, boletas, anexos):
        super().__init__()
        self.setWindowTitle("Resultados de la Consulta")
        self.layout = QVBoxLayout()

        self.tabla_resultados = QTableWidget()
        self.tabla_resultados.setRowCount(len(escrituras))
        self.tabla_resultados.setColumnCount(4)  # Añadir una columna para los anexos
        self.tabla_resultados.setHorizontalHeaderLabels(["ESCRITURA", "RADICADO", "BOLETA DE RENTAS", "VER ANEXO"])

        for i in range(len(escrituras)):
            self.tabla_resultados.setItem(i, 0, QTableWidgetItem(escrituras[i]))
            self.tabla_resultados.setItem(i, 1, QTableWidgetItem(str(radicados[i])))
            self.tabla_resultados.setItem(i, 2, QTableWidgetItem(boletas[i]))

            if anexos.get(radicados[i]):
                btn_ver_anexo = QPushButton("Ver Anexo")
                btn_ver_anexo.setIcon(QIcon("icono_anexo.png"))  # Establecer un ícono para el botón de ver anexo
                btn_ver_anexo.clicked.connect(lambda state, radicado=radicados[i], anexo=anexos.get(radicados[i]): self.abrir_anexo(radicado, anexo))
                self.tabla_resultados.setCellWidget(i, 3, btn_ver_anexo)
            else:
                self.tabla_resultados.setItem(i, 3, QTableWidgetItem("No disponible"))

        self.layout.addWidget(self.tabla_resultados)
        self.setLayout(self.layout)

    def abrir_anexo(self, radicado, anexo):
        try:
            # Configuración para ejecutar el navegador
            driver = webdriver.Chrome()
            
            # Abrir la página del anexo
            driver.get(anexo)
            
            # Mostrar un mensaje al usuario
            QMessageBox.information(self, "Anexo", "Anexo abierto. Por favor, haga clic en 'Cerrar' cuando haya terminado de visualizarlo.")

        except Exception as e:
            QMessageBox.warning(self, "Error", f"Error al abrir el anexo: {e}")


def buscar_nir_para_consulta():
    excel_file_path = r'C:\Users\DAVID\Desktop\DAVID\N-15\DAVID\LIBROS XLSM\HISTORICO.xlsm'
    wb = openpyxl.load_workbook(excel_file_path, data_only=True)
    sheet = wb["ESCRITURAS FISICAS (DIARIO)"]

    # Lista para almacenar los números de radicado a consultar
    radicados_a_consultar = []

    # Iterar sobre las filas de la columna "C" (índice 3) buscando las que tienen "SIN INICIAR" en la columna "D"
    for row in sheet.iter_rows(min_row=2, max_col=4, max_row=sheet.max_row, values_only=True):
        if row[3] == "SIN INICIAR" and row[1]:  # Verificar si la celda en la columna "B" no está vacía
            radicados_a_consultar.append(row[2])  # Agregar el valor de la columna "C" a la lista

    return radicados_a_consultar


def buscar_escritura(radicado):
    excel_file_path = r'C:\Users\DAVID\Desktop\DAVID\N-15\DAVID\LIBROS XLSM\HISTORICO.xlsm'
    wb = openpyxl.load_workbook(excel_file_path, data_only=True)
    sheet = wb["ESCRITURAS FISICAS (DIARIO)"]

    # Buscar la escritura correspondiente al radicado
    for row in sheet.iter_rows(min_row=2, max_col=3, max_row=sheet.max_row, values_only=True):
        if row[1] == radicado:
            return row[0]

    return "No disponible"


def obtener_anexos(radicados_a_consultar):
    try:
        anexos = {}  # Diccionario para almacenar los enlaces de los anexos

        # Configuración para ejecutar el navegador
        driver = webdriver.Chrome()

        for radicado in radicados_a_consultar:
            try:
                # Construir la URL para la consulta del radicado
                url_consulta = f"https://mercurio.antioquia.gov.co/mercurio/servlet/ControllerMercurio?command=anexos&tipoOperacion=abrirLista&idDocumento={radicado}&tipDocumento=R&now=Date()&ventanaEmergente=S&origen=NTR"

                # Abrir la página de consulta
                driver.get(url_consulta)

                # Esperar a que se cargue la tabla de anexos
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//td[contains(text(), 'Total Anexos')]"))
                )

                # Obtener el enlace del anexo si está presente
                enlaces_anexos = driver.find_elements(By.XPATH, "//tr[@class='contenido']/td/a")
                if enlaces_anexos:
                    anexos[radicado] = url_consulta  # Almacenar el enlace del anexo

            except Exception as e:
                print(f"Error al consultar radicado {radicado}: {e}")

        return anexos

    except Exception as e:
        print(f"Error en la consulta: {e}")
        return {}


def mostrar_resultados(escrituras, radicados, boletas, anexos):
    app = QApplication(sys.argv)
    ventana_resultados = ResultadosVentana(escrituras, radicados, boletas, anexos)
    ventana_resultados.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    radicados_a_consultar = buscar_nir_para_consulta()
    anexos = obtener_anexos(radicados_a_consultar)

    escrituras = [buscar_escritura(radicado) for radicado in radicados_a_consultar]
    radicados = radicados_a_consultar
    boletas = ["LIQUIDACIÓN DE RENTAS LISTA" if radicado in anexos else "NO HAY ANEXO DE RENTAS DISPONIBLE" for radicado in radicados]

    mostrar_resultados(escrituras, radicados, boletas, anexos)
