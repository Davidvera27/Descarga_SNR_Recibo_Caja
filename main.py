from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import os
import time
import openpyxl
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QPushButton, QWidget, QVBoxLayout

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Descarga de Documentos")

        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["ESCRITURA", "ESTADO DEL RECIBO DE CAJA", "BOTÓN DE DESCARGA"])

        self.download_all_button = QPushButton("DESCARGAR DISPONIBLES")
        self.download_all_button.clicked.connect(self.download_all)

        layout = QVBoxLayout()
        layout.addWidget(self.table)
        layout.addWidget(self.download_all_button)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # Iniciar sesión y cargar archivo Excel al iniciar la aplicación
        self.driver = self.iniciar_sesion("JULIAN.ZAPATA", "Notaria15")
        if self.driver:
            self.load_excel("C:/Users/DAVID/Desktop/DAVID/N-15/DAVID/LIBROS XLSM/HISTORICO.xlsm")

    def iniciar_sesion(self, usuario, contraseña):
        try:
            # Iniciar el navegador Chrome
            driver = webdriver.Chrome()
            
            # Abrir la página de inicio de sesión
            driver.get("https://radicacion.supernotariado.gov.co/app/inicio.dma")

            # Introducir el nombre de usuario
            usr_input = driver.find_element(By.ID, "formLogin:usrlogin")
            usr_input.send_keys(usuario)

            # Introducir la contraseña
            pwd_input = driver.find_element(By.ID, "formLogin:j_idt8")
            pwd_input.send_keys(contraseña)

            # Hacer clic en el botón de inicio de sesión
            driver.find_element(By.ID, "formLogin:j_idt11").click()

            # Esperar a que se cargue completamente la página de bienvenida y aceptarla
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "infoBienvenida")))
            driver.find_element(By.ID, "formInfoBienvenida:id-acepta-bienvenida").click()

            # Devolver el controlador del navegador
            return driver

        except Exception as e:
            print("Error:", e)
            return None

    def load_excel(self, archivo_excel):
        wb = openpyxl.load_workbook(archivo_excel, data_only=True)
        sheet = wb["PRINCIPAL"]

        # Iterar sobre las filas
        for row in sheet.iter_rows(min_row=2, max_col=11, values_only=False):
            if row[10].value == "DESCARGADA":
                nir = row[3].value  # Valor de la columna "D"
                if nir is not None:  # Verificar si el valor no es None
                    self.search_documents(nir)

    def search_documents(self, nir):
        try:
            # Navegar a la página de búsqueda
            self.driver.get("https://radicacion.supernotariado.gov.co/app/external/documentary-manager.dma")

            # Esperar un momento para que la página de búsqueda se cargue completamente
            self.driver.implicitly_wait(3)

            # Enviar el NIR al formulario de búsqueda
            nir_field = self.driver.find_element(By.ID, "formFilterDocManager:j_idt48")
            nir_field.clear()
            nir_field.send_keys(nir)

            # Hacer clic en el botón de búsqueda
            buscar_btn = self.driver.find_element(By.ID, "formFilterDocManager:j_idt79")
            buscar_btn.click()

            # Esperar a que se cargen los resultados de la búsqueda
            WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.ID, "formDocsManager")))

            # Verificar si el recibo de caja está disponible
            recibo_caja_disponible = False
            try:
                # Intentar encontrar el enlace del recibo de caja
                recibo_caja_link = self.driver.find_element(By.ID, "formDocsManager:j_idt84:0:j_idt158")
                recibo_caja_disponible = True
            except NoSuchElementException:
                pass

            # Verificar si ya existe una fila para este NIR en la tabla
            row_idx = -1
            for idx in range(self.table.rowCount()):
                if self.table.item(idx, 0).text() == str(nir):
                    row_idx = idx
                    break

            # Si no se encuentra una fila para este NIR, agregar una nueva fila a la tabla
            if row_idx == -1:
                row_idx = self.table.rowCount()
                self.table.insertRow(row_idx)
                self.table.setItem(row_idx, 0, QTableWidgetItem(str(nir)))  # Escritura

            # Actualizar el estado del recibo de caja en la fila correspondiente
            estado_recibo_caja = "RECIBO DE CAJA LISTO PARA DESCARGAR" if recibo_caja_disponible else "RECIBO DE CAJA NO DISPONIBLE"
            self.table.setItem(row_idx, 1, QTableWidgetItem(estado_recibo_caja))  # Estado del recibo de caja

            # Si el recibo de caja está disponible y no se ha agregado el botón de descarga, agregarlo
            if recibo_caja_disponible and self.table.cellWidget(row_idx, 2) is None:
                button = QPushButton("DESCARGAR")
                button.clicked.connect(lambda _, nir=nir: self.download_document(nir))
                self.table.setCellWidget(row_idx, 2, button)

        except Exception as e:
            print(f"Error durante la búsqueda del NIR {nir}: {e}")

    def download_document(self, nir):
        try:
            # Realizar la búsqueda nuevamente para obtener el enlace de descarga
            self.search_documents(nir)

            # Hacer clic en el enlace del recibo de caja
            recibo_caja_link = self.driver.find_element(By.ID, "formDocsManager:j_idt84:0:j_idt158")
            recibo_caja_link.click()

            # Esperar a que se cargue el recibo de caja
            WebDriverWait(self.driver, 7).until(EC.presence_of_element_located((By.ID, "formDocsManager:j_idt207")))

            # Hacer clic en el botón de descarga
            descargar_btn = self.driver.find_element(By.ID, "formDocsManager:j_idt207")
            descargar_btn.click()

            # Esperar hasta que el archivo se descargue completamente
            tiempo_inicial = time.time()
            while True:
                if time.time() - tiempo_inicial > 60:  # Aumentar el tiempo de espera máximo a 60 segundos
                    print("Tiempo de espera para la descarga excedido.")
                    break
                archivos_en_descargas = [f for f in os.listdir("E:/Downloads") if f.endswith('.pdf')]
                if archivos_en_descargas and os.path.getsize(os.path.join("E:/Downloads", archivos_en_descargas[-1])) > 0:
                    nombre_archivo = archivos_en_descargas[-1]
                    nuevo_nombre = f"BR{nir}.pdf"
                    ruta_archivo_original = os.path.join("E:/Downloads", nombre_archivo)
                    ruta_archivo_nuevo = os.path.join("E:/Downloads", nuevo_nombre)
                    os.rename(ruta_archivo_original, ruta_archivo_nuevo)
                    print(f"Archivo PDF descargado y guardado correctamente como {nuevo_nombre}")
                    break
                time.sleep(1)

        except NoSuchElementException as e:
            print(f"El documento no está disponible para la escritura {nir}")

        except Exception as e:
            print(f"Error durante la descarga del archivo: {str(e)}")

    def download_all(self):
        # Iterar sobre las filas de la tabla
        for row_idx in range(self.table.rowCount()):
            # Obtener el NIR de la fila actual
            nir = self.table.item(row_idx, 0).text()
            # Verificar si el estado del recibo de caja es "RECIBO DE CAJA LISTO PARA DESCARGAR"
            estado_recibo_caja = self.table.item(row_idx, 1).text()
            if estado_recibo_caja == "RECIBO DE CAJA LISTO PARA DESCARGAR":
                self.download_document(nir)

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

