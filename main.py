import openpyxl
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.chrome.options import Options
import inicio_sesion
import buscar_documentos
import descargar_archivos

def consultar_anexos():
    # Configurar el servicio del controlador de Chrome
    service = Service('PYTHON\\WebScraping\\chromedriver-win64\\chromedriver.exe')

    # Configurar las opciones del navegador
    options = Options()
    options.add_argument('--start-maximized')  # Para iniciar maximizado

    # Inicializar el navegador
    driver = WebDriver(service=service, options=options)

    try:
        usuario = "JULIAN.ZAPATA"
        contraseña = "Notaria15"
        inicio_sesion.iniciar_sesion(driver, usuario, contraseña)

        archivo_excel = "C:/Users/DAVID/Desktop/DAVID/N-15/DAVID/LIBROS XLSM/HISTORICO.xlsm"
        wb = openpyxl.load_workbook(archivo_excel, data_only=True)
        sheet = wb["PRINCIPAL"]

        # Iterar sobre las filas
        for row in sheet.iter_rows(min_row=2, max_col=11, values_only=False):
            if row[10].value == "DESCARGADA":
                nir = row[3].value  # Valor de la columna "D"
                if nir is not None:  # Verificar si el valor no es None
                    buscar_documentos.buscar_documentos(driver, nir)
                    descargar_archivos.descargar_archivo(driver, row)
                else:
                    print(f"La escritura {row[1].value} no contiene NIR")
    finally:
        # Cerrar el navegador al finalizar
        driver.quit()

# Llamar a la función para iniciar la consulta
consultar_anexos()
