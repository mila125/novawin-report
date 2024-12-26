from tkinter import Tk, Label, Button, Entry, filedialog
import configparser
import os
import threading
from novarep import main
from HK import hk_main
from DFT import dft_main
from BJHA import bjha_main
from BJHD import bjhd_main
from FFHAGRAPH import ffhagraph_main
from NKAGRAPH import nkagraph_main
# Archivo de configuración
config_file = "config.ini"
ruta_qps = "s"
ruta_csv = "s"
ruta_pdf = "s"
ruta_novawin = "s"
def seleccionar_archivo(entry):
 filename = filedialog.askopenfilename(
     parent=ventana,
     title="Examinar archivo"
     
 )
 print(filename)
 entry.insert(0, filename)

def cargar_configuracion():
    
    """Cargar rutas desde un archivo de configuración."""
    config = configparser.ConfigParser()
    config.read(config_file)
    if "Rutas" in config:
        entry_qps.insert(0, config["Rutas"].get("ruta_qps", "").replace("/", "\\"))
        entry_csv.insert(0, config["Rutas"].get("ruta_csv", "").replace("/", "\\"))
        entry_novawin.insert(0, config["Rutas"].get("ruta_novawin", "").replace("/", "\\"))
        entry_pdf.insert(0, config["Rutas"].get("ruta_pdf", "").replace("/", "\\"))
  

def guardar_configuracion():
    """Guardar las rutas en un archivo de configuración."""
    ruta_qps = entry_qps.get()
    ruta_csv = entry_csv.get()
    ruta_pdf = entry_pdf.get()
    ruta_novawin = entry_novawin.get()
    print(ruta_qps)
    print(ruta_csv)
    print(ruta_pdf)
    print(ruta_novawin)
    config = configparser.ConfigParser()
    config["Rutas"] = {
        "ruta_qps": ruta_qps,
        "ruta_csv": ruta_csv,
        "ruta_novawin": ruta_novawin,
        "ruta_pdf": ruta_pdf
    }
    with open(config_file, "w") as configfile:
        config.write(configfile)
    print("Rutas guardadas en config.ini")

def seleccionar_carpeta(entry):
    """Abrir un cuadro de diálogo para seleccionar una carpeta y mostrar la ruta en el Entry."""
    ruta_carpeta = filedialog.askdirectory()
    if ruta_carpeta:
        entry.delete(0, "end")
        entry.insert(0, ruta_carpeta)
        
def obtener_rutas_hk():
    ruta_qps = entry_qps.get()
    ruta_csv = entry_csv.get()
    ruta_novawin = entry_novawin.get()
    ruta_pdf = entry_pdf.get()
    print(ruta_qps)
    print(ruta_csv)
    print(ruta_pdf)
    print(ruta_novawin)
    ventana.quit()
    ventana.destroy()

    hk_main("a","b")
def obtener_rutas_dft():
    ruta_qps = entry_qps.get()
    ruta_csv = entry_csv.get()
    ruta_novawin = entry_novawin.get()
    ruta_pdf = entry_pdf.get()
    print(ruta_qps)
    print(ruta_csv)
    print(ruta_pdf)
    print(ruta_novawin)
    ventana.quit()
    ventana.destroy()
    dft_main("a","b")
def obtener_rutas_bjha():
    ruta_qps = entry_qps.get()
    ruta_csv = entry_csv.get()
    ruta_novawin = entry_novawin.get()
    ruta_pdf = entry_pdf.get()
    print(ruta_qps)
    print(ruta_csv)
    print(ruta_pdf)
    print(ruta_novawin)
    ventana.quit()
    ventana.destroy()
    bjha_main("a","b")
    
def obtener_rutas_bjhd():
    ruta_qps = entry_qps.get()
    ruta_csv = entry_csv.get()
    ruta_novawin = entry_novawin.get()
    ruta_pdf = entry_pdf.get()
    print(ruta_qps)
    print(ruta_csv)
    print(ruta_pdf)
    print(ruta_novawin)
    ventana.quit()
    ventana.destroy()
    bjhd_main("a","b")
    
def obtener_rutas_ffhagraph():
    ruta_qps = entry_qps.get()
    ruta_csv = entry_csv.get()
    ruta_novawin = entry_novawin.get()
    ruta_pdf = entry_pdf.get()
    print(ruta_qps)
    print(ruta_csv)
    print(ruta_pdf)
    print(ruta_novawin)
    ventana.quit()
    ventana.destroy()
    ffhagraph_main("a","b")
    
def obtener_rutas_nkagraph():
    ruta_qps = entry_qps.get()
    ruta_csv = entry_csv.get()
    ruta_novawin = entry_novawin.get()
    ruta_pdf = entry_pdf.get()
    print(ruta_qps)
    print(ruta_csv)
    print(ruta_pdf)
    print(ruta_novawin)
    ventana.quit()
    ventana.destroy()
    nkagraph_main("a","b")
    
def obtener_rutas():
    """Recuperar las rutas ingresadas o seleccionadas por el usuario."""
    ruta_qps = entry_qps.get()
    ruta_csv = entry_csv.get()
    ruta_novawin = entry_novawin.get()
    ruta_pdf = entry_pdf.get()

    #if not os.path.isdir(ruta_qps):
    #    print(f"Error: La ruta para archivos .qps no es válida: {ruta_qps}")
    #    return
    #if not os.path.isdir(ruta_csv):
    #    print(f"Error: La ruta para guardar reportes no es válida: {ruta_csv}")
    #    return
    #if not os.path.isfile(os.path.join(ruta_novawin, "NovaWin.exe")):
    #    print(f"Error: No se encontró NovaWin.exe en la ruta especificada: {ruta_novawin}")
    #    return
    #if not os.path.isdir(ruta_pdf):
    #    print(f"Error: La ruta para guardar pdf no es válida: {ruta_pdf}")
    #    return

def obtener_rutas_novawin():
    ruta_qps = entry_qps.get()
    ruta_csv = entry_csv.get()
    ruta_novawin = entry_novawin.get()
    ruta_pdf = entry_pdf.get()
    print(ruta_qps)
    print(ruta_csv)
    print(ruta_pdf)
    print(ruta_novawin)
    

    # Llamar al módulo novarep
    ventana.quit()
    ventana.destroy()
    print(ruta_qps)
    main(ruta_qps, ruta_csv,ruta_novawin)

    

# Crear ventana principal
ventana = Tk()
ventana.title("Selector de Rutas")
ventana.geometry("1100x1000")

# Deshabilitar redimensionamiento
ventana.resizable(False, False)

# Etiquetas y campos de entrada
Label(ventana, text="Ruta de archivos .qps:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
entry_qps = Entry(ventana, width=50)
entry_qps.grid(row=0, column=1, padx=10, pady=10)
# Botón para seleccionar archivo
Button(ventana, text="Seleccionar Archivo", command=seleccionar_archivo(entry_qps)).grid(row=0, column=2, padx=10, pady=10)


Label(ventana, text="Ruta para guardar reportes:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
entry_csv = Entry(ventana, width=50)
entry_csv.grid(row=1, column=1, padx=10, pady=10)
Button(ventana, text="Seleccionar Archivo", command=seleccionar_archivo(entry_csv)).grid(row=1, column=2, padx=10, pady=10)
Label(ventana, text="Ruta de NovaWin:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
entry_novawin = Entry(ventana, width=50)
entry_novawin.grid(row=2, column=1, padx=10, pady=10)
Button(ventana, text="Seleccionar Archivo", command=seleccionar_archivo(entry_novawin)).grid(row=2, column=2, padx=10, pady=10)
Label(ventana, text="Ruta para guardar PDFs:").grid(row=3, column=0, padx=10, pady=10, sticky="w")
entry_pdf = Entry(ventana, width=50)
entry_pdf.grid(row=3, column=1, padx=10, pady=10)
Button(ventana, text="Seleccionar Archivo", command=seleccionar_archivo(entry_pd)).grid(row=3, column=2, padx=10, pady=10)


# Botón para confirmar rutas
Button(ventana, text="Guardar configuracion", command=guardar_configuracion).grid(row=7, column=1, pady=10)

# Botón para ejecutar hk
Button(ventana, text="Ejecutar HK", command=obtener_rutas_hk).grid(row=8, column=1, pady=10)

# Botón para ejecutar dft
Button(ventana, text="Ejecutar DFT", command=obtener_rutas_dft).grid(row=9, column=1, pady=10)

# Botón para ejecutar bjha
Button(ventana, text="Ejecutar BJHA", command=obtener_rutas_bjha).grid(row=10, column=1, pady=10)

# Botón para ejecutar bjhd
Button(ventana, text="Ejecutar BJHD", command=obtener_rutas_bjhd).grid(row=11, column=1, pady=10)

# Botón para ejecutar ffhagraph
Button(ventana, text="Ejecutar FFHAGRAPH", command=obtener_rutas_ffhagraph).grid(row=12, column=1, pady=10)

# Botón para ejecutar nkagraph
Button(ventana, text="Ejecutar NKAGRAPH", command=obtener_rutas_nkagraph).grid(row=13, column=1, pady=10)

# Botón para ejecutar novawin
Button(ventana, text="Ejecutar NovaWin", command=obtener_rutas_novawin).grid(row=14, column=1, pady=10)




# Cargar configuración existente si está disponible
cargar_configuracion()

# Ejecutar bucle principal
ventana.mainloop()