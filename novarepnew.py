# -*- coding: utf-8 -*-
from pywinauto import Application
import time
from datetime import datetime
import os

def limpiar_caracteres(texto):
    """Eliminar caracteres no imprimibles de una cadena."""
    return ''.join(c for c in texto if c.isprintable())

def encontrar_primer_qps(directorio):
    """Buscar el primer archivo con extensión .qps en el directorio."""
    for archivo in os.listdir(directorio):
        if archivo.endswith(".qps"):
            return os.path.join(directorio, archivo)
    return None

# Solicitar rutas al usuario
path_qps = input("Por favor, ingresa ruta para extraer archivos: ").strip()
path_csv = input("Por favor, ingresa ruta donde se van a guardar los reportes: ").strip()
path_novawin = input("Por favor, ingresa ruta donde se encuentra NovaWin: ").strip()

# Validar existencia de directorios y archivo ejecutable
if not os.path.isdir(path_qps):
    raise FileNotFoundError(f"La ruta especificada para los archivos .qps no existe: {path_qps}")
if not os.path.isdir(path_csv):
    raise FileNotFoundError(f"La ruta especificada para guardar los reportes no existe: {path_csv}")
if not os.path.isfile(os.path.join(path_novawin, "NovaWin.exe")):
    raise FileNotFoundError(f"No se encontró NovaWin.exe en la ruta especificada: {path_novawin}")

# Ruta al ejecutable de NovaWin
novawin_path = os.path.join(path_novawin, "NovaWin.exe")

# Buscar el primer archivo .qps
file_to_open_nameonly = encontrar_primer_qps(path_qps)
if not file_to_open_nameonly:
    raise FileNotFoundError("No se encontró ningún archivo .qps en el directorio especificado.")

# Generar la ruta completa del archivo a abrir
file_to_open = limpiar_caracteres(file_to_open_nameonly)

print(f"Archivo .qps encontrado: {file_to_open}")

# Generar la ruta completa para el archivo exportado
fecha_str = datetime.now().strftime("%Y%m%d_%H-%M-%S")
exported_file_path = limpiar_caracteres(f"{path_csv}\\reporte_{fecha_str}.csv")

print(f"Archivo exportado será guardado en: {exported_file_path}")

# Iniciar la aplicación NovaWin
app = Application(backend="uia").start(novawin_path)

# Esperar a que la ventana principal se cargue
time.sleep(5)

# Conectar con la ventana principal
main_window = app.window(title_re=".*NovaWin.*")

# Seleccionar la opción de menú para abrir archivo
main_window.menu_select("File->Open")

# Esperar a que aparezca el cuadro de diálogo
time.sleep(2)

# Interactuar con el cuadro de diálogo 'Abrir'
open_dialog = app.window(class_name="#32770")
if open_dialog.exists():
    print("Cuadro de diálogo 'Abrir' localizado.")

    open_dialog.wait("exists ready", timeout=10)
    
    # Buscar el control de tipo 'Edit' para escribir la ruta del archivo
    edit_box = open_dialog.child_window(class_name="Edit")
    if edit_box.exists():
        edit_box.set_edit_text(file_to_open)
        print(f"Ruta del archivo escrita en el cuadro de texto: {file_to_open}")
    else:
        raise Exception("No se pudo localizar el control Edit para ingresar la ruta del archivo.")

    # Obtener todos los controles del cuadro de diálogo
    all_controls = open_dialog.children()

    # Imprimir detalles de todos los controles
    for i, control in enumerate(all_controls):
        try:
           control_type = control.control_type() if hasattr(control, 'control_type') else 'N/A'
           control_text = control.window_text()
           print(f"Control {i}: {control_text} - Tipo: {control_type}")
        except Exception as e:
           print(f"Error al procesar control {i}: {e}")

    # Seleccionar el botón específico usando índices
    open_button = open_dialog.child_window( class_name="Button",found_index=0) 
    open_button.click_input()
    
else:
    raise Exception("No se pudo localizar el cuadro de diálogo 'Abrir'.")

# Continuar con la selección de menús y exportación de datos
try:
    # Imprimir detalles de los controles y ventanas hijas de main_window
    all_controls = main_window.children()
    main_window.print_control_identifiers()
    # Buscar el control por su clase
    graph_view_window = main_window.child_window(class_name="TGraphViewWindow")

    # Verificar si se ha encontrado el control
    if graph_view_window.exists():
      print("Componente 'TGraphViewWindow' encontrado.")
      graph_view_window.print_control_identifiers()  # Imprime los identificadores del control
      # Realizar un clic derecho sobre el componente
      graph_view_window.right_click_input()
      time.sleep(1)
    else:
      print("No se encontró el componente 'TGraphViewWindow'.")
    # Imprimir los controles y sus propiedades

    context_menu = app.window(title_re=".*Context.*")
    tables_menu_item = context_menu.child_window(title="Tables", control_type="MenuItem")
    tables_menu_item.click_input()
    print("Menú 'Tables' seleccionado.")

    bet_menu_item = app.window(best_match="Tables").child_window(title="BET", control_type="MenuItem")
    bet_menu_item.click_input()
    print("Submenú 'BET' seleccionado.")

    # Verificar todos los controles del menú antes de intentar interactuar con ellos
    # Esperar a que el menú "Single Point Surface Area" se haga disponible
    single_point_menu_item = bet_menu_item.child_window(title="Single Point Surface Area", control_type="MenuItem")
    single_point_menu_item.wait('exists', timeout=10)  # Esperar hasta que el control esté disponible

    if single_point_menu_item.exists():
     single_point_menu_item.click_input()
     print("Opción 'Single Point Surface Area' seleccionada.")
    else:
     print("No se encontró la opción 'Single Point Surface Area'.")
    time.sleep(2)
    context_menu = app.window(title_re=".*Context.*")
    savecsv_menu_item = context_menu.child_window(title="Export to .CSV", control_type="MenuItem")
    savecsv_menu_item.click_input()

    csv_dialog = app.window(title_re=".*Name  File.*")
    csv_dialog.child_window(auto_id="1148", control_type="Edit").type_keys(exported_file_path, with_spaces=True)
    csv_dialog.child_window(title="Guardar", control_type="Button").click_input()
    print("Archivo exportado exitosamente.")

except Exception as e:
    print("Error durante la exportación:", e)

    
