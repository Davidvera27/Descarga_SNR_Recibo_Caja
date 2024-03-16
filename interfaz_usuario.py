import tkinter as tk
import os

root = tk.Tk()

tk.Label(root, text="Usuario:").grid(row=0, column=0, padx=5, pady=5)
usuario_entry = tk.Entry(root)
usuario_entry.grid(row=0, column=1, padx=5, pady=5)
tk.Label(root, text="Contraseña:").grid(row=1, column=0, padx=5, pady=5)
contraseña_entry = tk.Entry(root, show="*")
contraseña_entry.grid(row=1, column=1, padx=5, pady=5)

def iniciar_proceso():
    os.system('start /wait cmd /c "python PYTHON\WebScraping\Excel\Descarga_SNR_Recibo_Caja\main.py"')

tk.Button(root, text="Descargar Recibo de caja", command=iniciar_proceso).grid(row=2, columnspan=2, padx=5, pady=5)

def consultar_recibo_pago():
    os.system('start /wait cmd /c "python PYTHON\WebScraping\Excel\Descarga_SNR_Recibo_Caja\Descargar_ReciboPago.py"')

tk.Button(root, text="Consultar Recibo de Pago", command=consultar_recibo_pago).grid(row=3, columnspan=2, padx=5, pady=5)

def proceso4():
    from Descargar_Rentas import consultar_anexos
    consultar_anexos()

tk.Button(root, text="Consultar Liquidación de Rentas", command=proceso4).grid(row=4, columnspan=2, padx=5, pady=5)

root.mainloop()
