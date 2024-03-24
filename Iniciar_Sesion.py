import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PyQt5.QtWidgets import QMessageBox

def iniciar_sesion(usuario, contraseña):
    try:
        driver = webdriver.Chrome()
        driver.get("https://radicacion.supernotariado.gov.co/app/inicio.dma")

        usr_input = driver.find_element(By.ID, "formLogin:usrlogin")
        usr_input.send_keys(usuario)

        pwd_input = driver.find_element(By.ID, "formLogin:j_idt8")
        pwd_input.send_keys(contraseña)

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "formLogin:j_idt11"))).click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "infoBienvenida")))

        driver.get("https://radicacion.supernotariado.gov.co/app/external/documentary-manager.dma")

        QMessageBox.information(None, "Inicio de Sesión Exitoso", "¡Inicio de sesión exitoso!")
    except Exception as e:
        QMessageBox.critical(None, "Error", "Error durante el inicio de sesión: El sitio web no está disponible. Por favor, inténtelo de nuevo más tarde.")
    finally:
        if 'driver' in locals():
            driver.quit()