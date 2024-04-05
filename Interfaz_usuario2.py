import sys
import time
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLineEdit, QLabel, QPushButton, QMessageBox
from PyQt5.QtCore import QFile, QTextStream
from PyQt5.QtGui import QIcon, QPalette, QPixmap, QBrush
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import subprocess
from Iniciar_Sesion import iniciar_sesion, cerrar_sesion, buscar_escritura, realizar_busqueda_sitio_web

class InterfazUsuario(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.driver = None  # Inicializar el atributo driver como None

    def initUI(self):
        self.setWindowTitle('Panel - Supernotariado')
        layout = QVBoxLayout()
        # Configurar el fondo de la ventana con una imagen
        palette = self.palette()
        palette.setBrush(QPalette.Background, QBrush(QPixmap("PYTHON\WebScraping\Excel\Descarga_SNR_Recibo_Caja\Icons\wallpaper3.jpg")))
        self.setPalette(palette)

        # Usuario
        self.lbl_usuario = QLabel('Usuario:')
        self.txt_usuario = QLineEdit()
        self.lbl_usuario.setObjectName("lblUsuario")  # Establecer un ID
        self.txt_usuario.setObjectName("txtUsuario")  # Establecer un ID
        layout.addWidget(self.lbl_usuario)
        layout.addWidget(self.txt_usuario)

        # Contraseña
        self.lbl_contraseña = QLabel('Contraseña:')
        self.txt_contraseña = QLineEdit()
        self.txt_contraseña.setEchoMode(QLineEdit.Password)
        self.lbl_contraseña.setObjectName("lblContraseña")  # Establecer un ID
        self.txt_contraseña.setObjectName("txtContraseña")  # Establecer un ID
        layout.addWidget(self.lbl_contraseña)
        layout.addWidget(self.txt_contraseña)

        # Botón Autenticarse
        self.btn_autenticarse = QPushButton('Autenticarse', self)
        self.btn_autenticarse.setIcon(QIcon("PYTHON/WebScraping/Excel/Descarga_SNR_Recibo_Caja/Icons/Authenticate_icon.png"))
        self.btn_autenticarse.clicked.connect(self.iniciar_sesion)
        self.btn_autenticarse.setObjectName("btnAutenticarse")  # Establecer un ID
        layout.addWidget(self.btn_autenticarse)
        
        # Botón Cerrar Sesión
        self.btn_cerrar_sesion = QPushButton('Cerrar Sesión', self)
        self.btn_cerrar_sesion.setIcon(QIcon("PYTHON/WebScraping/Excel/Descarga_SNR_Recibo_Caja/Icons/logout_icon.png"))
        self.btn_cerrar_sesion.clicked.connect(self.iniciar_o_cerrar_sesion)
        self.btn_cerrar_sesion.setEnabled(False)  # Inicialmente deshabilitado
        layout.addWidget(self.btn_cerrar_sesion)

        # Nuevos botones
        self.btn_descargar_recibo = QPushButton('Descargar recibo de caja')
        self.btn_descargar_recibo.setIcon(QIcon("PYTHON\WebScraping\Excel\Descarga_SNR_Recibo_Caja\Icons\Download_icon.png"))
        self.btn_descargar_recibo.clicked.connect(self.descargar_recibo_caja)
        layout.addWidget(self.btn_descargar_recibo)

        self.btn_consultar_estado_rel = QPushButton('Consultar estado en REL')
        self.btn_consultar_estado_rel.setIcon(QIcon("PYTHON\WebScraping\Excel\Descarga_SNR_Recibo_Caja\Icons\search_process_icon.png"))
        self.btn_consultar_estado_rel.clicked.connect(self.consultar_estado_rel)
        layout.addWidget(self.btn_consultar_estado_rel)

        self.btn_consultar_estado_rentas = QPushButton('Consultar estado de rentas')
        self.btn_consultar_estado_rentas.setIcon(QIcon("PYTHON\WebScraping\Excel\Descarga_SNR_Recibo_Caja\Icons\Rentas_icon.png"))
        self.btn_consultar_estado_rentas.clicked.connect(self.consultar_estado_rentas)
        layout.addWidget(self.btn_consultar_estado_rentas)

        # Etiqueta y cuadro de texto para NIR
        self.lbl_nir = QLabel('NIR:')
        self.txt_nir = QLineEdit()
        self.lbl_nir.setVisible(False)
        self.txt_nir.setVisible(False)
        layout.addWidget(self.lbl_nir)
        layout.addWidget(self.txt_nir)

        # Botón Buscar para NIR
        self.btn_buscar_nir = QPushButton('Buscar')
        self.btn_buscar_nir.setIcon(QIcon("PYTHON\WebScraping\Excel\Descarga_SNR_Recibo_Caja\Icons\Search_NIR_icon.png"))
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
        self.btn_buscar_escritura.setIcon(QIcon("PYTHON\WebScraping\Excel\Descarga_SNR_Recibo_Caja\Icons\search_number_icon.png"))
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
        # Conectar evento returnPressed a la función iniciar_sesion
        self.txt_usuario.returnPressed.connect(self.iniciar_sesion)
        self.txt_contraseña.returnPressed.connect(self.iniciar_sesion)
        # Aplicar el archivo CSS
        self.setStyleSheet(self.load_stylesheet())
        
        
            # Funciones de los nuevos botones y otras funciones

    def load_stylesheet(self):
        stylesheet = QFile("PYTHON\WebScraping\Excel\Descarga_SNR_Recibo_Caja\styles.css")
        if stylesheet.open(QFile.ReadOnly | QFile.Text):
            stream = QTextStream(stylesheet)
            return stream.readAll()
        else:
            QMessageBox.critical(self, "Error", "No se pudo cargar el archivo de estilos CSS.")
            return ""


    # Funciones de los nuevos botones

    def descargar_recibo_caja(self):
        try:
            subprocess.run(["python", "PYTHON/WebScraping/Excel/Descarga_SNR_Recibo_Caja/main.py"])
        except Exception as e:
            print("Error al ejecutar el script para descargar recibo de caja:", e)

    def consultar_estado_rel(self):
        try:
            subprocess.run(["python", "PYTHON/WebScraping/Excel/Descarga_SNR_Recibo_Caja/Descargar_ReciboPago.py"])
        except Exception as e:
            print("Error al ejecutar el script para consultar estado en REL:", e)

    def consultar_estado_rentas(self):
        try:
            subprocess.run(["python", "PYTHON/WebScraping/Excel/Descarga_SNR_Recibo_Caja/Descargar_Rentas.py"])
        except Exception as e:
            print("Error al ejecutar el script para consultar estado de rentas:", e)

    def iniciar_o_cerrar_sesion(self):
        if self.sender() == self.btn_autenticarse:
            self.iniciar_sesion()
        elif self.sender() == self.btn_cerrar_sesion:
            self.cerrar_sesion()

    def iniciar_sesion(self):
        usuario = self.txt_usuario.text()
        contraseña = self.txt_contraseña.text()
        self.driver = iniciar_sesion(usuario, contraseña)  # Guardar el controlador devuelto
        if self.driver is not None:
            self.btn_autenticarse.setEnabled(False)
            self.btn_cerrar_sesion.setEnabled(True)
            self.show_labels_and_textboxes()  # Mostrar elementos ocultos después del inicio de sesión exitoso

    def cerrar_sesion(self):
        cerrar_sesion(self.driver)
        self.btn_autenticarse.setEnabled(True)
        self.btn_cerrar_sesion.setEnabled(False)
        self.hide_labels_and_textboxes()  # Ocultar elementos después de cerrar sesión

    def hide_labels_and_textboxes(self):
        # Ocultar etiquetas y cuadros de texto
        self.lbl_nir.setVisible(False)
        self.txt_nir.setVisible(False)
        self.btn_buscar_nir.setVisible(False)
        self.lbl_escritura.setVisible(False)
        self.txt_escritura.setVisible(False)
        self.btn_buscar_escritura.setVisible(False)
        self.lbl_certificado.setVisible(False)
        self.txt_certificado.setVisible(False)
        self.btn_buscar_certificado.setVisible(False)



    def closeEvent(self, event):
        # Cierre del navegador Selenium cuando se cierra la aplicación PyQt5
        self.driver.quit()
        event.accept()

    def show_labels_and_textboxes(self):
        try:
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
        nir = buscar_escritura(escritura, self.txt_nir)
        if nir:
            realizar_busqueda_sitio_web(self.driver, nir)
            
    def realizar_busqueda_sitio_web(self, nir):
        if self.driver:
            realizar_busqueda_sitio_web(self.driver, nir)
        else:
            print("El controlador del navegador no está inicializado.")

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