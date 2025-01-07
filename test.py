from tkinter import Tk, Label, Button, Entry, filedialog

# Crear ventana principal
ventana = Tk()
ventana.title("Selector de Rutas")
ventana.geometry("1100x1000")

# Deshabilitar redimensionamiento
ventana.resizable(False, False)

# Función para seleccionar un directorio
def seleccionar_directorio(entry):
    ruta_directorio = filedialog.askdirectory(parent=ventana, title="Seleccionar directorio")
    if ruta_directorio:  # Si se ha seleccionado un directorio
        entry.delete(0, "end")  # Borra el contenido actual
        entry.insert(0, ruta_directorio)  # Inserta la nueva ruta del directorio

# Etiquetas y campos de entrada
Label(ventana, text="Selecciona un directorio:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
entry_directorio = Entry(ventana, width=50)
entry_directorio.grid(row=0, column=1, padx=10, pady=10)

# Botón para seleccionar un directorio
Button(ventana, text="Seleccionar Directorio", command=lambda: seleccionar_directorio(entry_directorio)).grid(row=0, column=2, padx=10, pady=10)

# Ejecutar bucle principal
ventana.mainloop()