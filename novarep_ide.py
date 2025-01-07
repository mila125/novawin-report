from tkinter import Tk, Label, Button, Entry, filedialog
import configparser
import threading
from novarep import main
from HK import hk_main
from DFT import dft_main
from BJHA import bjha_main
from BJHD import bjhd_main
from FFHAGRAPH import ffhagraph_main
from NKAGRAPH import nkagraph_main
from datetime import datetime

# Archivo de configuración
config_file = "config.ini"

# Función genérica para ejecutar módulos con rutas
def ejecutar_modulo(funcion):
    """Ejecutar una función pasando las rutas como argumentos."""
    ruta_qps = entry_qps.get()
    ruta_csv = entry_csv.get()
    ruta_novawin = entry_novawin.get()
    ruta_pdf = entry_pdf.get()
    print(f"Ejecutando {funcion.__name__} con rutas:")
    print(f"QPS: {ruta_qps}, CSV: {ruta_csv}, NovaWin: {ruta_novawin}, PDF: {ruta_pdf}")
    ventana.quit()
    ventana.destroy()
    funcion(ruta_qps, ruta_csv, ruta_novawin)

# Función genérica para seleccionar rutas
def seleccionar_ruta(entry, is_file=True):
    """Seleccionar archivo o carpeta y actualizar el Entry correspondiente."""
    try:
        ruta = (
            filedialog.askopenfilename(title="Seleccionar archivo") if is_file
            else filedialog.askdirectory(title="Seleccionar directorio")
        )
        if ruta:
            entry.delete(0, "end")
            entry.insert(0, ruta)
    except Exception as e:
        print(f"Error al seleccionar la ruta: {e}")

def seleccionar_ruta_async(entry, is_file=True):
    """Abrir el diálogo de selección en un hilo separado."""
    threading.Thread(target=seleccionar_ruta, args=(entry, is_file), daemon=True).start()

# Funciones para cargar y guardar configuración
def cargar_configuracion():
    config = configparser.ConfigParser()
    config.read(config_file)
    if "Rutas" in config:
        entry_qps.insert(0, config["Rutas"].get("ruta_qps", ""))
        entry_csv.insert(0, config["Rutas"].get("ruta_csv", ""))
        entry_novawin.insert(0, config["Rutas"].get("ruta_novawin", ""))
        entry_pdf.insert(0, config["Rutas"].get("ruta_pdf", ""))

def guardar_configuracion():
    config = configparser.ConfigParser()
    config["Rutas"] = {
        "ruta_qps": entry_qps.get(),
        "ruta_csv": entry_csv.get(),
        "ruta_novawin": entry_novawin.get(),
        "ruta_pdf": entry_pdf.get()
    }
    with open(config_file, "w") as configfile:
        config.write(configfile)
    print("Rutas guardadas en config.ini")

# Crear ventana principal
ventana = Tk()
ventana.title("Selector de Rutas")
ventana.geometry("1000x800")
ventana.resizable(False, False)

# Etiquetas y campos de entrada
Label(ventana, text="Ruta de archivos .qps:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
entry_qps = Entry(ventana, width=50)
entry_qps.grid(row=0, column=1, padx=10, pady=10)
Button(ventana, text="Seleccionar Archivo", command=lambda: seleccionar_ruta(entry_qps)).grid(row=0, column=2, padx=10, pady=10)

Label(ventana, text="Ruta de directorio CSV:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
entry_csv = Entry(ventana, width=50)
entry_csv.grid(row=1, column=1, padx=10, pady=10)
Button(ventana, text="Seleccionar Carpeta", command=lambda: seleccionar_ruta_async(entry_csv, False)).grid(row=1, column=2, padx=10, pady=10)

Label(ventana, text="Ruta de NovaWin:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
entry_novawin = Entry(ventana, width=50)
entry_novawin.grid(row=2, column=1, padx=10, pady=10)
Button(ventana, text="Seleccionar Archivo", command=lambda: seleccionar_ruta(entry_novawin)).grid(row=2, column=2, padx=10, pady=10)

Label(ventana, text="Ruta para guardar PDFs:").grid(row=3, column=0, padx=10, pady=10, sticky="w")
entry_pdf = Entry(ventana, width=50)
entry_pdf.grid(row=3, column=1, padx=10, pady=10)
Button(ventana, text="Seleccionar Carpeta", command=lambda: seleccionar_ruta_async(entry_pdf, False)).grid(row=3, column=2, padx=10, pady=10)

# Botones para guardar configuración y ejecutar funciones
Button(ventana, text="Guardar Configuración", command=guardar_configuracion).grid(row=4, column=1, pady=10)
Button(ventana, text="Ejecutar HK", command=lambda: ejecutar_modulo(hk_main)).grid(row=5, column=1, pady=10)
Button(ventana, text="Ejecutar DFT", command=lambda: ejecutar_modulo(dft_main)).grid(row=6, column=1, pady=10)
Button(ventana, text="Ejecutar BJHA", command=lambda: ejecutar_modulo(bjha_main)).grid(row=7, column=1, pady=10)
Button(ventana, text="Ejecutar BJHD", command=lambda: ejecutar_modulo(bjhd_main)).grid(row=8, column=1, pady=10)
Button(ventana, text="Ejecutar FFHAGRAPH", command=lambda: ejecutar_modulo(ffhagraph_main)).grid(row=9, column=1, pady=10)
Button(ventana, text="Ejecutar NKAGRAPH", command=lambda: ejecutar_modulo(nkagraph_main)).grid(row=10, column=1, pady=10)
Button(ventana, text="Ejecutar BET", command=lambda: ejecutar_modulo(main)).grid(row=11, column=1, pady=10)

# Cargar configuración inicial
cargar_configuracion()

# Iniciar ventana
ventana.mainloop()