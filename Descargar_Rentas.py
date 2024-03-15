from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import os
import time

def descargar_archivo(driver, row):
    try:
        recibo_caja_link = driver.find_element(By.ID, "formDocsManager:j_idt84:0:j_idt158")
        recibo_caja_link.click()  # Hacer clic en el enlace del recibo de caja

        # Esperar a que se cargue el recibo de caja
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "formDocsManager:j_idt197")))

        # Hacer clic en el botón de descarga
        descargar_btn = driver.find_element(By.ID, "formDocsManager:j_idt197")
        descargar_btn.click()

        # Esperar hasta que el archivo se descargue completamente
        tiempo_inicial = time.time()
        while True:
            if time.time() - tiempo_inicial > 30:
                print("Tiempo de espera para la descarga excedido.")
                break
            archivos_en_descargas = [f for f in os.listdir("E:/Downloads") if f.endswith('.pdf')]
            if archivos_en_descargas:
                break
            time.sleep(1)

        # Obtener el nombre del último archivo descargado y renombrarlo si se descargó correctamente
        if archivos_en_descargas:
            nombre_archivo = archivos_en_descargas[-1]
            numero_escritura = row[1].value  # Obtener el número de escritura
            nuevo_nombre = f"{numero_escritura}.pdf"
            os.rename(f"E:/Downloads/{nombre_archivo}", f"E:/Downloads/{nuevo_nombre}")
            print(f"Archivo '{nombre_archivo}' descargado y renombrado como '{nuevo_nombre}'")
    except NoSuchElementException:
        print("No se encontró el elemento para descargar el archivo.")
