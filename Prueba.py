import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QLineEdit, QLabel
from PyQt5.QtGui import QPixmap, QPalette, QBrush
from PyQt5.QtCore import Qt
import subprocess

class InterfazUsuario(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Inicio de Sesión Supernotariado')

        layout = QVBoxLayout(self)

        # Cargar la imagen de fondo
        background_img_path = "PYTHON/WebScraping/Excel/Descarga_SNR_Recibo_Caja/Icons/Fondo_interfaz.jpg"
        self.setFixedSize(800, 600)  # Fijar tamaño de la ventana
        palette = self.palette()
        palette.setBrush(QPalette.Background, QBrush(QPixmap(background_img_path)))
        self.setPalette(palette)

        # Añadir widgets al layout
        self.add_widgets(layout)

    def add_widgets(self, layout):
        # Usuario
        self.lbl_usuario = QLabel('Usuario:')
        self.txt_usuario = QLineEdit()
        layout.addWidget(self.lbl_usuario)
        layout.addWidget(self.txt_usuario)

        # Contraseña
        self.lbl_contraseña = QLabel('Contraseña:')
        self.txt_contraseña = QLineEdit()
        layout.addWidget(self.lbl_contraseña)
        layout.addWidget(self.txt_contraseña)

        # Botón Autenticarse
        self.btn_autenticarse = QPushButton('Autenticarse')
        self.btn_autenticarse.clicked.connect(self.iniciar_sesion)
        layout.addWidget(self.btn_autenticarse)

        # Nuevos botones
        buttons_info = [
            ("Descargar recibo de caja", self.descargar_recibo_caja),
            ("Consultar estado en REL", self.consultar_estado_rel),
            ("Consultar estado de rentas", self.consultar_estado_rentas)
        ]

        for btn_text, btn_func in buttons_info:
            btn = QPushButton(btn_text)
            btn.clicked.connect(btn_func)
            layout.addWidget(btn)

    def iniciar_sesion(self):
        usuario = self.txt_usuario.text()
        contraseña = self.txt_contraseña.text()
        # Tu código para iniciar sesión aquí

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


def main():
    app = QApplication(sys.argv)
    ventana = InterfazUsuario()
    ventana.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
