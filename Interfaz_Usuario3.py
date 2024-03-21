import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QLineEdit, QVBoxLayout, QWidget, QHBoxLayout, QComboBox, QMessageBox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Interfaz de Usuario")

        # Crear los elementos de la interfaz de usuario
        self.title_label = QLabel("Visor Documental", self)
        self.subtitle_label = QLabel("Visualización de los documentos de cada proceso", self)

        self.home_button = QPushButton("Ir al inicio", self)
        self.refresh_button = QPushButton("Actualizar la página", self)
        
        self.username_label = QLabel("Usuario:", self)
        self.username_lineedit = QLineEdit(self)

        self.password_label = QLabel("Contraseña:", self)
        self.password_lineedit = QLineEdit(self)
        self.password_lineedit.setEchoMode(QLineEdit.Password)

        self.login_button = QPushButton("Iniciar Sesión", self)

        self.search_label = QLabel("Buscar por:", self)
        self.search_combobox = QComboBox(self)
        self.search_combobox.addItems(["NIR", "Documento"])

        self.search_lineedit_nir = QLineEdit(self)
        self.search_lineedit_doc = QLineEdit(self)
        self.search_lineedit_doc.hide()

        self.second_search_label = QLabel("Tipo de documento:", self)
        self.second_search_combobox = QComboBox(self)
        self.second_search_combobox.addItems(["Escritura", "Certificado"])
        self.second_search_combobox.hide()

        self.search_button = QPushButton("Buscar", self)

        # Crear el diseño de la interfaz de usuario
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.title_label)
        self.layout.addWidget(self.subtitle_label)

        self.layout.addWidget(self.username_label)
        self.layout.addWidget(self.username_lineedit)

        self.layout.addWidget(self.password_label)
        self.layout.addWidget(self.password_lineedit)

        self.layout.addWidget(self.login_button)

        self.layout.addWidget(self.home_button)
        self.layout.addWidget(self.refresh_button)

        self.layout.addWidget(self.search_label)
        self.search_layout = QHBoxLayout()
        self.search_layout.addWidget(self.search_combobox)
        self.search_layout.addWidget(self.search_lineedit_nir)
        self.search_layout.addWidget(self.search_lineedit_doc)
        self.layout.addLayout(self.search_layout)

        self.second_search_layout = QHBoxLayout()
        self.second_search_layout.addWidget(self.second_search_label)
        self.second_search_layout.addWidget(self.second_search_combobox)
        self.layout.addLayout(self.second_search_layout)

        self.layout.addWidget(self.search_button)

        self.central_widget = QWidget()
        self.central_widget.setLayout(self.layout)
        self.setCentralWidget(self.central_widget)

        # Conectar señales y slots
        self.search_combobox.currentIndexChanged.connect(self.toggle_second_search)
        self.login_button.clicked.connect(self.iniciar_sesion)

        # Inicializar el controlador de Chrome
        self.driver = None

    def toggle_second_search(self, index):
        if index == 0:  # NIR
            self.search_lineedit_nir.show()
            self.search_lineedit_doc.hide()
            self.second_search_combobox.hide()
        elif index == 1:  # Documento
            self.search_lineedit_nir.hide()
            self.search_lineedit_doc.show()
            self.second_search_combobox.show()

    def iniciar_sesion(self):
        usuario = self.username_lineedit.text()
        contraseña = self.password_lineedit.text()

        try:
            # Inicializa el controlador de Chrome si aún no se ha iniciado
            if not self.driver:
                options = webdriver.ChromeOptions()
                options.add_experimental_option('excludeSwitches', ['enable-logging'])
                self.driver = webdriver.Chrome(options=options)

            url = "https://radicacion.supernotariado.gov.co/app/inicio.dma"
            self.driver.get(url)

            campo_usuario = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "formLogin:usrlogin")))
            campo_usuario.send_keys(usuario)

            campo_contraseña = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "formLogin:j_idt8")))
            campo_contraseña.send_keys(contraseña)

            boton_autenticarse = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, "formLogin:j_idt11")))
            boton_autenticarse.click()

            ventana_emergente = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "infoBienvenida")))
            boton_aceptar = WebDriverWait(ventana_emergente, 10).until(EC.element_to_be_clickable((By.ID, "formInfoBienvenida:id-acepta-bienvenida")))
            boton_aceptar.click()

            time.sleep(2)  # Esperar a que la página se cargue completamente

            # Muestra información relevante del "Visor Documental" o cualquier otra sección
            # Puedes agregar el código necesario para interactuar con la interfaz después del inicio de sesión

            # Realizar búsqueda si se proporciona un término de búsqueda
            if self.search_combobox.currentText() == "NIR":
                search_term = self.search_lineedit_nir.text()
            elif self.search_combobox.currentText() == "Documento":
                search_term = self.search_lineedit_doc.text()

            if search_term:
                self.realizar_busqueda(search_term)

        except Exception as e:
            print("Error:", e)
            QMessageBox.critical(self, "Error", str(e))

    def realizar_busqueda(self, search_term):
        # Aquí puedes agregar el código para realizar la búsqueda en el sitio web
        # Utiliza self.driver para interactuar con la página web
        
        # Por ejemplo, puedes encontrar el campo de búsqueda y enviar el término de búsqueda
        campo_busqueda = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "campo_busqueda")))
        campo_busqueda.send_keys(search_term)

        # Luego, haz clic en el botón de búsqueda
        boton_busqueda = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, "boton_busqueda")))
        boton_busqueda.click()

        # Esperar a que aparezcan los resultados y procesarlos
        time.sleep(2)

        # Obtener los resultados de la búsqueda y mostrarlos en una ventana emergente
        resultados = self.driver.find_elements_by_class_name("resultado_busqueda")
        if resultados:
            # Construir el texto con los resultados de la búsqueda
            texto_resultados = "\n".join([resultado.text for resultado in resultados])
            QMessageBox.information(self, "Resultados de la búsqueda", texto_resultados)
        else:
            QMessageBox.information(self, "Resultados de la búsqueda", "No se encontraron resultados.")

    def closeEvent(self, event):
        # Cerrar el navegador al cerrar la ventana principal
        if self.driver:
            self.driver.quit()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

