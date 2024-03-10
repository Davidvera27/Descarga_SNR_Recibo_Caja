import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


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

def consultar_anexos():
    # Obtener la lista de radicados a consultar
    radicados_a_consultar = buscar_nir_para_consulta()

    # Inicializar el navegador
    driver = webdriver.Chrome()

    for radicado in radicados_a_consultar:
        try:
            # Construir la URL para la consulta del radicado
            url_consulta = f"https://mercurio.antioquia.gov.co/mercurio/servlet/ControllerMercurio?command=anexos&tipoOperacion=abrirLista&idDocumento={radicado}&tipDocumento=R&now=Date()&ventanaEmergente=S&origen=NTR"

            # Abrir la página de consulta
            driver.get(url_consulta)

            # Esperar a que se cargue la tabla de anexos
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//td[contains(text(), 'Total Anexos')]")))

            # Encontrar todos los enlaces de anexos y hacer clic en cada uno
            enlaces_anexos = driver.find_elements(By.XPATH, "//tr[@class='contenido']/td/a")
            for enlace in enlaces_anexos:
                enlace.click()
                # Agregar aquí el código para manejar la apertura de los anexos

        except Exception as e:
            print(f"Error al consultar radicado {radicado}: {e}")

    # Mostrar mensaje al finalizar
    print("Proceso de consulta de anexos finalizado.")

    # Mantener el navegador abierto para visualizar los anexos
    input("Presiona ENTER para cerrar el navegador...")
    driver.quit()

# Llamar a la función para iniciar la consulta
consultar_anexos()
