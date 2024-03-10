import tkinter as tk
from tkinter import messagebox
from threading import Thread
import openpyxl
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.chrome.options import Options
import Descargar_Rentas as descargar_rentas
from inicio_sesion import iniciar_sesion
from buscar_documentos import buscar_documentos
from descargar_archivos import descargar_archivo




# Función principal para consultar anexos
def consultar_anexos():
    # Configurar el servicio del controlador de Chrome
    service = Service('PYTHON\\WebScraping\\chromedriver-win64\\chromedriver.exe')

    # Configurar las opciones del navegador
    options = Options()
    options.add_argument('--start-maximized')  # Para iniciar maximizado

    # Inicializar el navegador
    driver = WebDriver(service=service, options=options)

    try:
        usuario = usuario_entry.get()
        contraseña = contraseña_entry.get()
        iniciar_sesion(driver, usuario, contraseña)

        archivo_excel = "C:/Users/DAVID/Desktop/DAVID/N-15/DAVID/LIBROS XLSM/HISTORICO.xlsm"
        wb = openpyxl.load_workbook(archivo_excel, data_only=True)
        sheet = wb["PRINCIPAL"]

        for row in sheet.iter_rows(min_row=2, max_col=11, values_only=False):
            if row[10].value == "DESCARGADA":
                nir = row[3].value
                if nir is not None:
                    buscar_documentos(driver, nir)
                    descargar_archivo(driver, row)
                else:
                    print(f"La escritura {row[1].value} no contiene NIR")
    finally:
        driver.quit()

# Funciones para procesos futuros
#def proceso1():
    messagebox.showinfo("Mensaje", "Este es el proceso 1")

#def proceso2():
    messagebox.showinfo("Mensaje", "Este es el proceso 2")

#def proceso3():
    messagebox.showinfo("Mensaje", "Este es el proceso 3")

# Función para llamar a la función consultar_anexos del módulo Descargar_Rentas
def proceso4():
    # Llama a la función consultar_anexos del módulo Descargar_Rentas
    descargar_rentas.consultar_anexos()

# Crear la ventana principal
root = tk.Tk()
root.title("Interfaz de Usuario")

# Etiquetas y campos de entrada
tk.Label(root, text="Usuario:").grid(row=0, column=0, padx=5, pady=5)
usuario_entry = tk.Entry(root)
usuario_entry.grid(row=0, column=1, padx=5, pady=5)
tk.Label(root, text="Contraseña:").grid(row=1, column=0, padx=5, pady=5)
contraseña_entry = tk.Entry(root, show="*")
contraseña_entry.grid(row=1, column=1, padx=5, pady=5)

# Botón para iniciar el proceso
# Función para iniciar el proceso en un hilo aparte
def iniciar_proceso():
    Thread(target=consultar_anexos).start()
tk.Button(root, text="Iniciar Proceso", command=iniciar_proceso).grid(row=2, columnspan=2, padx=5, pady=5)

# Botones para procesos futuros
#tk.Button(root, text="Proceso 1", command=proceso1).grid(row=3, column=0, padx=5, pady=5)
#tk.Button(root, text="Proceso 2", command=proceso2).grid(row=3, column=1, padx=5, pady=5)
#tk.Button(root, text="Proceso 3", command=proceso3).grid(row=4, columnspan=2, padx=5, pady=5)

# Botón para llamar a la función de Descargar_Rentas
tk.Button(root, text="Consultar Anexos", command=proceso4).grid(row=5, columnspan=2, padx=5, pady=5)

# Mostrar la interfaz de usuario
root.mainloop()