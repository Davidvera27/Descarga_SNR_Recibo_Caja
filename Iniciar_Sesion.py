from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def iniciar_sesion(usuario, contraseña):
    # Inicializar el driver del navegador
    driver = webdriver.Chrome()  # Esto inicializará Chrome
    try:
        # Navegar a la página de inicio de sesión
        driver.get("https://radicacion.supernotariado.gov.co/app/inicio.dma")
        
        # Esperar a que se cargue la página de inicio de sesión
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "formLogin:usrlogin")))
        
        # Ingresar las credenciales en el formulario de inicio de sesión
        usr_input = driver.find_element(By.ID, "formLogin:usrlogin")
        usr_input.send_keys(usuario)
        
        pwd_input = driver.find_element(By.ID, "formLogin:j_idt8")
        pwd_input.send_keys(contraseña)
        
        # Hacer clic en el botón "Autenticarse"
        autenticarse_btn = driver.find_element(By.ID, "formLogin:j_idt11")
        autenticarse_btn.click()
        
        # Esperar a que se cargue la página de inicio después del inicio de sesión
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "id_of_an_element_on_the_next_page")))
        
        print("Inicio de sesión exitoso.")
    
    except Exception as e:
        print("Error durante el inicio de sesión:", e)
    
    finally:
        # Cerrar el navegador al finalizar
        driver.quit()

# Ejemplo de uso
if __name__ == "__main__":
    usuario = input("Ingresa tu usuario: ")
    contraseña = input("Ingresa tu contraseña: ")
    iniciar_sesion(usuario, contraseña)
