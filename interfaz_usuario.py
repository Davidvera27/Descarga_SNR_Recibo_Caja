import tkinter as tk
from tkinter import messagebox, ttk
from threading import Thread
import openpyxl
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import subprocess

import Descargar_Rentas as descargar_rentas
from inicio_sesion import iniciar_sesion
from buscar_documentos import buscar_documentos
from descargar_archivos import descargar_archivo

def mostrar_resultados(resultados):
    ventana_resultados = tk.Toplevel()
    ventana_resultados.title("Resultados de la consulta")

    tabla_frame = ttk.Frame(ventana_resultados)
    tabla_frame.pack(expand=True, fill="both")
    tabla = ttk.Treeview(tabla_frame, columns=("#1", "#2", "#3", "#4", "#5", "#6", "#7", "#8", "#9", "#10", "#11", "#12"))  # Agregamos una nueva columna
    tabla.heading("#0", text="Acciones")
    tabla.heading("#1", text="Ver")
    tabla.heading("#2", text="Columna Disponible")
    tabla.heading("#3", text="Escritura")
    tabla.heading("#4", text="Estado del proceso")
    tabla.heading("#5", text="NIR")
    tabla.heading("#6", text="Vigencia de Rentas")
    tabla.heading("#7", text="Estado en el REL")
    tabla.heading("#8", text="Tramitadora")
    tabla.heading("#9", text="Número de paquete")
    tabla.heading("#10", text="Estado de liquidación")
    tabla.heading("#11", text="Recibo de pago (Estado)")
    tabla.heading("#12", text="Descarga exitosa")  # Nueva columna

    tabla.column("#0", width=100)

    scrollbar_x = ttk.Scrollbar(tabla_frame, orient="horizontal", command=tabla.xview)
    tabla.configure(xscrollcommand=scrollbar_x.set)
    scrollbar_x.pack(side="bottom", fill="x")

    for i, resultado in enumerate(resultados):
        tabla.insert("", "end", iid=i, values=(f"Ver Detalles {i}", *resultado))

    tabla.pack(expand=True, fill="both")

def consultar_anexos(usuario, contraseña):
    service = Service('PYTHON\\WebScraping\\chromedriver-win64\\chromedriver.exe')
    options = Options()
    options.add_argument('--start-maximized')
    driver = WebDriver(service=service, options=options)

    try:
        iniciar_sesion(driver, usuario, contraseña)

        archivo_excel = "C:/Users/DAVID/Desktop/DAVID/N-15/DAVID/LIBROS XLSM/HISTORICO.xlsm"
        wb = openpyxl.load_workbook(archivo_excel, data_only=True)
        sheet = wb["PRINCIPAL"]

        resultados = []

        for row in sheet.iter_rows(min_row=2, max_col=11, values_only=False):
            if row[10].value == "DESCARGADA":
                nir = row[3].value
                if nir is not None:
                    buscar_documentos(driver, nir)
                    resultado_descarga = descargar_archivo(driver, row)
                    descarga_exitosa = "Sí" if resultado_descarga else "No"  # Indicamos si la descarga fue exitosa
                    resultados.append([cell.value for cell in row] + [descarga_exitosa])  # Agregamos la nueva información a la lista de resultados
                else:
                    print(f"La escritura {row[1].value} no contiene NIR")

        mostrar_resultados(resultados)

    except Exception as e:
        messagebox.showinfo("Error", f"Error: {e}")

    finally:
        driver.quit()

def consultar_recibo_pago():
    subprocess.run(['python', 'PYTHON\\WebScraping\\Excel\\Descarga_SNR_Recibo_Caja\\Descargar_ReciboPago.py'])

def consultar_rentas():
    descargar_rentas.consultar_anexos()

root = tk.Tk()
root.title("Interfaz de Usuario")

tk.Label(root, text="Usuario:").grid(row=0, column=0, padx=5, pady=5)
usuario_entry = tk.Entry(root)
usuario_entry.grid(row=0, column=1, padx=5, pady=5)
tk.Label(root, text="Contraseña:").grid(row=1, column=0, padx=5, pady=5)
contraseña_entry = tk.Entry(root, show="*")
contraseña_entry.grid(row=1, column=1, padx=5, pady=5)

def iniciar_proceso():
    usuario = usuario_entry.get()
    contraseña = contraseña_entry.get()
    Thread(target=consultar_anexos, args=(usuario, contraseña)).start()

tk.Button(root, text="Descargar Recibo de caja", command=iniciar_proceso).grid(row=2, columnspan=2, padx=5, pady=5)

tk.Button(root, text="Consultar Recibo de Pago", command=consultar_recibo_pago).grid(row=3, columnspan=2, padx=5, pady=5)

def proceso4():
    consultar_rentas()

tk.Button(root, text="Consultar Liquidación de Rentas", command=proceso4).grid(row=4, columnspan=2, padx=5, pady=5)

root.mainloop()
