from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Constantes
URL_INICIO_SESION = "https://radicacion.supernotariado.gov.co/app/inicio.dma"
ID_USUARIO = "formLogin:usrlogin"
ID_CONTRASEÑA = "formLogin:j_idt8"
ID_BOTON_INICIAR_SESION = "formLogin:j_idt11"

def iniciar_sesion(driver, usuario, contraseña):
    # Navegar a la página de inicio de sesión
    driver.get(URL_INICIO_SESION)

    # Diligenciar el formulario de inicio de sesión
    usuario_field = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, ID_USUARIO)))
    usuario_field.send_keys(usuario)
    
    contraseña_field = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, ID_CONTRASEÑA)))
    contraseña_field.send_keys(contraseña)
    
    # Hacer clic en el botón "Autenticarse"
    iniciar_sesion_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, ID_BOTON_INICIAR_SESION)))
    iniciar_sesion_btn.click()

    # Esperar un momento para que la página se cargue completamente
    driver.implicitly_wait(3)
