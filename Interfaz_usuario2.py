import sys
import time
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QLineEdit, QLabel
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl

class InterfazUsuario(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Inicio de Sesión Supernotariado')
        layout = QVBoxLayout()

        # Usuario
        self.lbl_usuario = QLabel('Usuario:')
        self.txt_usuario = QLineEdit()
        layout.addWidget(self.lbl_usuario)
        layout.addWidget(self.txt_usuario)

        # Contraseña
        self.lbl_contraseña = QLabel('Contraseña:')
        self.txt_contraseña = QLineEdit()
        self.txt_contraseña.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.lbl_contraseña)
        layout.addWidget(self.txt_contraseña)

        # Botón Autenticarse
        self.btn_autenticarse = QPushButton('Autenticarse')
        self.btn_autenticarse.clicked.connect(self.iniciar_sesion)
        layout.addWidget(self.btn_autenticarse)

        # Etiqueta y cuadro de texto para NIR
        self.lbl_nir = QLabel('NIR:')
        self.txt_nir = QLineEdit()
        self.lbl_nir.setVisible(False)
        self.txt_nir.setVisible(False)
        layout.addWidget(self.lbl_nir)
        layout.addWidget(self.txt_nir)

        # Botón Buscar para NIR
        self.btn_buscar_nir = QPushButton('Buscar')
        self.btn_buscar_nir.setVisible(False)
        self.btn_buscar_nir.clicked.connect(self.buscar_nir)
        layout.addWidget(self.btn_buscar_nir)

        # Etiqueta y cuadro de texto para Escritura
        self.lbl_escritura = QLabel('Escritura:')
        self.txt_escritura = QLineEdit()
        self.lbl_escritura.setVisible(False)
        self.txt_escritura.setVisible(False)
        layout.addWidget(self.lbl_escritura)
        layout.addWidget(self.txt_escritura)

        # Botón Buscar para Escritura
        self.btn_buscar_escritura = QPushButton('Buscar')
        self.btn_buscar_escritura.setVisible(False)
        self.btn_buscar_escritura.clicked.connect(self.buscar_escritura)
        layout.addWidget(self.btn_buscar_escritura)

        # Etiqueta y cuadro de texto para Certificado
        self.lbl_certificado = QLabel('Certificado:')
        self.txt_certificado = QLineEdit()
        self.lbl_certificado.setVisible(False)
        self.txt_certificado.setVisible(False)
        layout.addWidget(self.lbl_certificado)
        layout.addWidget(self.txt_certificado)

        # Botón Buscar para Certificado
        self.btn_buscar_certificado = QPushButton('Buscar')
        self.btn_buscar_certificado.setVisible(False)
        self.btn_buscar_certificado.clicked.connect(self.buscar_certificado)
        layout.addWidget(self.btn_buscar_certificado)

        self.setLayout(layout)

    def iniciar_sesion(self):
        usuario = self.txt_usuario.text()
        contraseña = self.txt_contraseña.text()

        # Abrir el navegador y navegar a la página de inicio de sesión
        try:
            self.driver = webdriver.Chrome()
            self.driver.get("https://radicacion.supernotariado.gov.co/app/inicio.dma")

            # Ingresar las credenciales y hacer clic en Autenticarse
            usr_input = self.driver.find_element(By.ID, "formLogin:usrlogin")
            usr_input.send_keys(usuario)

            pwd_input = self.driver.find_element(By.ID, "formLogin:j_idt8")
            pwd_input.send_keys(contraseña)

            # Desactivar el botón mientras se procesa
            self.btn_autenticarse.setEnabled(False)

            # Esperar a que el botón de autenticación esté habilitado
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, "formLogin:j_idt11")))
            autenticarse_btn = self.driver.find_element(By.ID, "formLogin:j_idt11")
            autenticarse_btn.click()

            # Esperar a que aparezca la ventana emergente
            info_bienvenida = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.ID, "infoBienvenida")))

            # Encontrar y hacer clic en el botón Aceptar
            boton_aceptar = info_bienvenida.find_element(By.ID, "formInfoBienvenida:id-acepta-bienvenida")
            boton_aceptar.click()

            # Redirigir directamente al Visor Documental
            self.driver.get("https://radicacion.supernotariado.gov.co/app/external/documentary-manager.dma")

            # Mostrar las etiquetas y cuadros de texto para la consulta
            self.lbl_nir.setVisible(True)
            self.txt_nir.setVisible(True)
            self.btn_buscar_nir.setVisible(True)
            self.lbl_escritura.setVisible(True)
            self.txt_escritura.setVisible(True)
            self.btn_buscar_escritura.setVisible(True)
            self.lbl_certificado.setVisible(True)
            self.txt_certificado.setVisible(True)
            self.btn_buscar_certificado.setVisible(True)

            # Simular espera mientras se procesa la autenticación (aquí deberías ajustar el tiempo según el sitio)
            time.sleep(5)

            # Una vez autenticado y en el Visor Documental, hacer algo más (por ejemplo, imprimir el título de la página)
            print(self.driver.title)

        except Exception as e:
            print("Error durante el inicio de sesión:", e)

    def buscar_nir(self):
        nir = self.txt_nir.text()
        print("Buscando NIR:", nir)
        try:
            # Ingresar el NIR en el campo de texto del sitio web
            input_nir = self.driver.find_element(By.ID, "formFilterDocManager:j_idt48")
            input_nir.clear()
            input_nir.send_keys(nir)

            # Hacer clic en el botón de búsqueda
            btn_buscar = self.driver.find_element(By.ID, "formFilterDocManager:j_idt79")
            btn_buscar.click()
            
            # Esperar a que se carguen los resultados de la búsqueda (puedes agregar más código aquí según sea necesario)
        except Exception as e:
            print("Error al buscar NIR en el sitio web:", e)

    def buscar_escritura(self):
        escritura = self.txt_escritura.text().strip()
        print("Buscando Escritura:", escritura)
        try:
            wb = openpyxl.load_workbook(r"C:\Users\DAVID\Desktop\DAVID\N-15\DAVID\LIBROS XLSM\HISTORICO.xlsm", data_only=True)
            sheet = wb["PRINCIPAL"]
            column_b_values = [str(cell.value).strip() for cell in sheet['B'] if cell.value]
            found = False
            for row_index, value in enumerate(column_b_values):
                if value == escritura:
                    nir = sheet.cell(row=row_index + 1, column=4).value
                    self.txt_nir.setText(str(nir))  # Set the NIR text
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

    def realizar_busqueda_sitio_web(self, nir):
        try:
            # Navigate to the webpage where the search will be performed
            self.driver.get("URL_DEL_SITIO_WEB")

            # Wait for the page to load and the NIR input field to be present
            wait = WebDriverWait(self.driver, 10)
            input_nir = wait.until(EC.presence_of_element_located((By.ID, "formFilterDocManager:j_idt48")))

            # Input the NIR into the text field
            input_nir.clear()
            input_nir.send_keys(nir)

            # Find and click the search button
            btn_buscar = self.driver.find_element(By.ID, "formFilterDocManager:j_idt79")
            btn_buscar.click()

            # Wait for the search results to load (you can add more code here as needed)

        except Exception as e:
            print("Error al interactuar con el sitio web:", e)

    def buscar_certificado(self):
        certificado = self.txt_certificado.text()
        print("Buscando Certificado:", certificado)
        # Aquí iría la lógica para buscar el certificado en el sitio web

def main():
    app = QApplication(sys.argv)
    ventana = InterfazUsuario()
    ventana.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
