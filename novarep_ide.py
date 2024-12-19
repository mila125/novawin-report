import os
from tkinter import Tk, Label, Button, Entry, filedialog
import configparser
from novarep import main
# Archivo de configuración
config_file = "config.ini"
def cargar_configuracion():
    """Cargar rutas desde un archivo de configuración."""
    config = configparser.ConfigParser()
    config.read(config_file)
    if "Rutas" in config:
        entry_qps.insert(0, config["Rutas"].get("ruta_qps", "").replace("/", "\\"))
        entry_csv.insert(0, config["Rutas"].get("ruta_csv", "").replace("/", "\\"))
        entry_novawin.insert(0, config["Rutas"].get("ruta_novawin", "").replace("/", "\\"))

def guardar_configuracion(ruta_qps, ruta_csv, ruta_novawin):
    """Guardar las rutas en un archivo de configuración."""
    config = configparser.ConfigParser()
    config["Rutas"] = {
        "ruta_qps": ruta_qps,
        "ruta_csv": ruta_csv,
        "ruta_novawin": ruta_novawin
    }
    with open(config_file, "w") as configfile:
        config.write(configfile)
    print("Rutas guardadas en config.ini")

def seleccionar_archivo(entry):
    """Abrir un cuadro de diálogo para seleccionar un archivo y mostrar la ruta en el Entry."""
    ruta_archivo = filedialog.askopenfilename()
    if ruta_archivo:
        entry.delete(0, "end")
        entry.insert(0, ruta_archivo)

def seleccionar_carpeta(entry):
    """Abrir un cuadro de diálogo para seleccionar una carpeta y mostrar la ruta en el Entry."""
    ruta_carpeta = filedialog.askdirectory()
    if ruta_carpeta:
        entry.delete(0, "end")
        entry.insert(0, ruta_carpeta)

def obtener_rutas():
    """Recuperar las rutas ingresadas o seleccionadas por el usuario."""
    ruta_qps = entry_qps.get()
    ruta_csv = entry_csv.get()
    ruta_novawin = entry_novawin.get()

    if not os.path.isdir(ruta_qps):
        print(f"Error: La ruta para archivos .qps no es válida: {ruta_qps}")
        return
    if not os.path.isdir(ruta_csv):
        print(f"Error: La ruta para guardar reportes no es válida: {ruta_csv}")
        return
    if not os.path.isfile(os.path.join(ruta_novawin, "NovaWin.exe")):
        print(f"Error: No se encontró NovaWin.exe en la ruta especificada: {ruta_novawin}")
        return

    guardar_configuracion(ruta_qps, ruta_csv, ruta_novawin)

    print("Rutas seleccionadas:")
    print(f"Ruta de archivos .qps: {ruta_qps}")
    print(f"Ruta de reportes: {ruta_csv}")
    print(f"Ruta de NovaWin: {ruta_novawin}")

    # Llamar a la función main de novarep.py con las rutas obtenidas
    main(ruta_qps, ruta_csv, ruta_novawin)

# Crear ventana principal
ventana = Tk()
ventana.title("Selector de Rutas")
ventana.geometry("600x300")

# Etiquetas y campos de entrada
Label(ventana, text="Ruta de archivos .qps:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
entry_qps = Entry(ventana, width=50)
entry_qps.grid(row=0, column=1, padx=10, pady=10)
Button(ventana, text="Seleccionar", command=lambda: seleccionar_carpeta(entry_qps)).grid(row=0, column=2, padx=10, pady=10)

Label(ventana, text="Ruta para guardar reportes:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
entry_csv = Entry(ventana, width=50)
entry_csv.grid(row=1, column=1, padx=10, pady=10)
Button(ventana, text="Seleccionar", command=lambda: seleccionar_carpeta(entry_csv)).grid(row=1, column=2, padx=10, pady=10)

Label(ventana, text="Ruta de NovaWin:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
entry_novawin = Entry(ventana, width=50)
entry_novawin.grid(row=2, column=1, padx=10, pady=10)
Button(ventana, text="Seleccionar", command=lambda: seleccionar_carpeta(entry_novawin)).grid(row=2, column=2, padx=10, pady=10)

# Botón para confirmar
Button(ventana, text="Confirmar", command=obtener_rutas).grid(row=3, column=1, pady=20)

# Cargar configuración existente si está disponible
cargar_configuracion()

# Ejecutar bucle principal
ventana.mainloop()



