import sys
import threading
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PyQt5 import QtWidgets, QtCore
import os
from selenium.webdriver.common.keys import Keys


class ConsultaAnexosInterfaz(QtWidgets.QWidget):
    def __init__(self, excel_file_path):
        super().__init__()
        self.setWindowTitle("Consulta de Anexos")
        self.setGeometry(100, 100, 800, 400)
        self.excel_file_path = excel_file_path
        self.proceso_en_curso = False  # Bandera para indicar si el proceso de consulta está en curso

        self.tabla_anexos = QtWidgets.QTableWidget(self)
        self.tabla_anexos.setGeometry(50, 50, 700, 300)
        self.tabla_anexos.setColumnCount(4)
        self.tabla_anexos.setHorizontalHeaderLabels(
            ["ESCRITURA PROCESADA", "ESTADO DE LIQUIDACIÓN", "IMPRIMIR", "VER"])

        self.driver = None  # Inicializar el driver al iniciar la interfaz

    def imprimir_anexo(self):
        boton = self.sender()
        fila = self.tabla_anexos.indexAt(boton.pos()).row()
        radicado = self.tabla_anexos.item(fila, 0).text()
        print(f"Imprimir anexo del caso {radicado}")

    def ver_anexo(self):
        boton = self.sender()
        fila = self.tabla_anexos.indexAt(boton.pos()).row()
        radicado = self.tabla_anexos.item(fila, 0).text()
        proceso = self.tabla_anexos.item(fila, 1).text()  # Obtener el valor de la columna "PROCESO"
        url_consulta = f"https://mercurio.antioquia.gov.co/mercurio/servlet/ControllerMercurio?command=anexos&tipoOperacion=abrirLista&idDocumento={radicado}&tipDocumento=R&now=Date()&ventanaEmergente=S&origen=NTR&proceso={proceso}"
        
        # Inicializar un nuevo navegador Chrome en una nueva instancia
        options = webdriver.ChromeOptions()
        options.add_argument("--new-window")  # Abrir en una nueva ventana
        driver = webdriver.Chrome(options=options)
        
        try:
            # Abrir la URL en la nueva ventana
            driver.get(url_consulta)
            
            # Esperar hasta que el usuario cierre manualmente la ventana
            input("Presiona ENTER para cerrar el navegador...")
            
        finally:
            # Cerrar el navegador después de que el usuario presione ENTER
            driver.quit()


    def closeEvent(self, event):
        if self.proceso_en_curso:
            event.ignore()  # Ignorar el evento de cierre mientras el proceso de consulta esté en curso
        else:
            if self.driver:
                self.driver.quit()  # Cerrar el driver al cerrar la ventana

    def consultar_anexos(self, radicado):
        try:
            if not self.driver:
                self.driver = webdriver.Chrome()
            for handle in self.driver.window_handles:
                self.driver.switch_to.window(handle)
                if radicado in self.driver.current_url:
                    # Identificar y hacer clic en el elemento HTML correspondiente al anexo
                    WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@title='Ver Anexo']"))).click()
        except Exception as e:
            print(f"Error al consultar radicado {radicado}: {e}")


def buscar_nir_para_consulta(excel_file_path):
    wb = openpyxl.load_workbook(excel_file_path, data_only=True)
    sheet = wb["ESCRITURAS FISICAS (DIARIO)"]

    radicados_a_consultar = []

    for row in sheet.iter_rows(min_row=2, max_col=4, max_row=sheet.max_row, values_only=True):
        if row[3] == "SIN INICIAR" and row[1]:
            radicados_a_consultar.append(row[2])

    return radicados_a_consultar


def consultar_anexos(ventana):
    excel_file_path = ventana.excel_file_path
    radicados_a_consultar = buscar_nir_para_consulta(excel_file_path)

    ventana.proceso_en_curso = True  # Indicar que el proceso está en curso

    if ventana.driver is None:
        ventana.driver = webdriver.Chrome()

    for radicado in radicados_a_consultar:
        try:
            estado_liquidacion, enlaces_anexos = obtener_estado_y_enlaces_anexos(ventana.driver, radicado)

            if enlaces_anexos:
                agregar_fila_tabla_anexos(ventana.tabla_anexos, radicado, estado_liquidacion, ventana.imprimir_anexo, ventana.ver_anexo)

        except Exception as e:
            print(f"Error al consultar radicado {radicado}: {e}")

    print("Proceso de consulta de anexos finalizado.")
    ventana.proceso_en_curso = False  # Indicar que el proceso ha finalizado


    print("Proceso de consulta de anexos finalizado.")
    ventana.proceso_en_curso = False  # Indicar que el proceso ha finalizado
def obtener_estado_y_enlaces_anexos(driver, radicado):
    try:
        # Construir la URL para la consulta del radicado
        url_consulta = f"https://mercurio.antioquia.gov.co/mercurio/servlet/ControllerMercurio?command=anexos&tipoOperacion=abrirLista&idDocumento={radicado}&tipDocumento=R&now=Date()&ventanaEmergente=S&origen=NTR"

        # Abrir la página de consulta en una nueva ventana
        driver.execute_script("window.open();")
        driver.switch_to.window(driver.window_handles[-1])
        driver.get(url_consulta)

        # Esperar a que se cargue la tabla de anexos
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//td[contains(text(), 'Total Anexos')]")))

        # Encontrar todos los enlaces de anexos
        enlaces_anexos = driver.find_elements(By.XPATH, "//tr[@class='contenido']/td/a")

        # Verificar si se encontraron enlaces de anexos
        if enlaces_anexos:
            cantidad_anexos = len(enlaces_anexos)
            estado_liquidacion = f"LIQUIDACIÓN LISTA ({cantidad_anexos})"
            return estado_liquidacion, enlaces_anexos
        else:
            return "SIN ANEXOS", None

    except Exception as e:
        print(f"Error al obtener anexos del radicado {radicado}: {e}")
        return "ERROR", None



def agregar_fila_tabla_anexos(tabla_anexos, radicado, estado_liquidacion, imprimir_anexo_func, ver_anexo_func):
    # Agregar una fila a la tabla de anexos
    fila = tabla_anexos.rowCount()
    tabla_anexos.insertRow(fila)
    tabla_anexos.setItem(fila, 0, QtWidgets.QTableWidgetItem(str(radicado)))  # ESCRITURA PROCESADA
    tabla_anexos.setItem(fila, 1, QtWidgets.QTableWidgetItem(estado_liquidacion))  # ESTADO DE LIQUIDACIÓN

    # Botón de imprimir anexo
    boton_imprimir = QtWidgets.QPushButton("IMPRIMIR")
    boton_imprimir.clicked.connect(imprimir_anexo_func)
    tabla_anexos.setCellWidget(fila, 2, boton_imprimir)

    # Botón de ver anexo
    boton_ver = QtWidgets.QPushButton("VER")
    boton_ver.clicked.connect(ver_anexo_func)
    tabla_anexos.setCellWidget(fila, 3, boton_ver)


if __name__ == "__main__":
    # Ruta del archivo Excel
    excel_file_path = r'C:\Users\DAVID\Desktop\DAVID\N-15\DAVID\LIBROS XLSM\HISTORICO.xlsm'

    try:
        # Iniciar la aplicación Qt
        app = QtWidgets.QApplication(sys.argv)

        # Crear la ventana de la interfaz gráfica
        ventana = ConsultaAnexosInterfaz(excel_file_path)
        ventana.show()

        # Consultar los anexos
        consultar_anexos(ventana)

        # Salir de la aplicación Qt
        sys.exit(app.exec_())

    finally:
        # Cerrar el navegador si está abierto
        if ventana.driver is not None:
            ventana.driver.quit()