import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton
from subprocess import Popen, PIPE

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Interfaz de Usuario')
        layout = QVBoxLayout()

        button1 = QPushButton('DESCARGAR RECIBO DE CAJA', self)
        button1.clicked.connect(self.descargar_recibo_caja)
        layout.addWidget(button1)

        button2 = QPushButton('CONSULTAR RECIBO DE PAGO', self)
        button2.clicked.connect(self.Descargar_ReciboPago)
        layout.addWidget(button2)

        button3 = QPushButton('CONSULTAR LIQUIDACIÓN DE RENTAS', self)
        button3.clicked.connect(self.consultar_liquidacion_rentas)
        layout.addWidget(button3)

        self.setLayout(layout)

    def descargar_recibo_caja(self):
        self.run_module("C:/Users/DAVID/Desktop/DAVID/PROGRAMACIÓN/PYTHON/WebScraping/Excel/Descarga_SNR_Recibo_Caja/main.py")

    def Descargar_ReciboPago(self):
        self.run_module("C:/Users/DAVID/Desktop/DAVID/PROGRAMACIÓN/PYTHON/WebScraping/Excel/Descarga_SNR_Recibo_Caja/Descargar_ReciboPago.py")

    def consultar_liquidacion_rentas(self):
        process = Popen(["python", "C:/Users/DAVID/Desktop/DAVID/PROGRAMACIÓN/PYTHON/WebScraping/Excel/Descarga_SNR_Recibo_Caja/Descargar_Rentas.py"], stdout=PIPE, stderr=PIPE, stdin=PIPE)
        stdout, _ = process.communicate(input=b'\n')
        output = stdout.decode().strip()
        print(output)

    def run_module(self, module_path):
        process = Popen(["python", module_path], stdout=PIPE, stderr=PIPE)
        stdout, _ = process.communicate()
        output = stdout.decode().strip()
        print(output)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec_())
