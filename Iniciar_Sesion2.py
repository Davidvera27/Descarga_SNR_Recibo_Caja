from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import getpass
import time

def iniciar_sesion(usuario, contraseña):
    driver = None
    try:
        # Inicializa el controlador de Chrome
        options = webdriver.ChromeOptions()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        driver = webdriver.Chrome(options=options)

        # URL del sitio web
        url = "https://radicacion.supernotariado.gov.co/app/inicio.dma"

        # Abre la página de inicio de sesión
        driver.get(url)

        # Espera hasta que el campo de usuario esté presente
        campo_usuario = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "formLogin:usrlogin")))
        campo_usuario.send_keys(usuario)

        # Espera hasta que el campo de contraseña esté presente
        campo_contraseña = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "formLogin:j_idt8")))
        
        # Ingresa la contraseña en el campo de texto
        campo_contraseña.send_keys(contraseña)

        # Haz clic en el botón "Autenticarse"
        boton_autenticarse = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "formLogin:j_idt11")))
        boton_autenticarse.click()

        # Espera hasta que la ventana emergente esté presente
        ventana_emergente = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "infoBienvenida")))

        # Haz clic en el botón "Aceptar" de la ventana emergente para cerrarla
        boton_aceptar = WebDriverWait(ventana_emergente, 10).until(EC.element_to_be_clickable((By.ID, "formInfoBienvenida:id-acepta-bienvenida")))
        boton_aceptar.click()
        
        # Esperar un breve tiempo para permitir que la página se cargue completamente
        time.sleep(2)

        # Hacer clic en el menú de hamburguesa
        menu_hamburguesa = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, "btn-menu")))
        menu_hamburguesa.click()

        # Esperar hasta que aparezca el menú desplegable
        opciones_menu = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "navigation")))

        # Encontrar y hacer clic en la opción "Visor Documental"
        link_visor_documental = WebDriverWait(opciones_menu, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Visor Documental')]")))
        link_visor_documental.click()

        # Espera a que el usuario presione Enter antes de cerrar el navegador
        input("Presiona ENTER para cerrar el navegador.")

    except Exception as e:
        print("Error:", e)

    finally:
        # Cierra el navegador
        if driver:
            driver.quit()

if __name__ == "__main__":
    try:
        # Solicitar al usuario que ingrese su nombre de usuario y contraseña
        usuario = input("Usuario: ")
        contraseña = getpass.getpass("Contraseña: ")

        # Inicia sesión con las credenciales proporcionadas
        iniciar_sesion(usuario, contraseña)
    
    except KeyboardInterrupt:
        print("\nProceso cancelado por el usuario.")
