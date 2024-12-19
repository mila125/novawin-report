# -*- coding: utf-8 -*-
from pywinauto import Application, findwindows
import time
from datetime import datetime
import os
import traceback

def limpiar_caracteres(texto):
    """Eliminar caracteres no imprimibles de una cadena."""
    return ''.join(c for c in texto if c.isprintable())

def encontrar_primer_qps(directorio):
    """Buscar el primer archivo con extensión .qps en el directorio."""
    for archivo in os.listdir(directorio):
        if archivo.endswith(".qps"):
            return os.path.join(directorio, archivo)
    return None

def inicializar_novawin(novawin_path):
    """Inicia la aplicación NovaWin."""
    try:
        app = Application(backend="uia").start(novawin_path)
        time.sleep(5)
        main_window = app.window(title_re=".*NovaWin.*")
        return app, main_window
    except Exception as e:
        print(f"Error al iniciar NovaWin: {e}")
        raise

def seleccionar_menu(window, ruta_menu):
    """Selecciona una opción de menú."""
    try:
        window.menu_select(ruta_menu)
        time.sleep(2)
    except Exception as e:
        print(f"Error al seleccionar menú '{ruta_menu}': {e}")
        raise

def interactuar_con_cuadro_dialogo(dialog, archivo):
    """Interactúa con el cuadro de diálogo para abrir o guardar archivos."""
    try:
        edit_box = dialog.child_window(class_name="Edit")
        edit_box.set_edit_text(archivo)
        open_button = dialog.child_window(class_name="Button", found_index=0)
        open_button.click_input()
    except Exception as e:
        print(f"Error al interactuar con el cuadro de diálogo: {e}")
        raise

def exportar_reporte(main_window, ruta_exportacion,app):
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


    # Seleccionar "Single Point Surface Area"
    bet_menu_item = app.window(best_match="BET")
    bet_menu_item.print_control_identifiers(depth=2)  # Aumenta el nivel de profundidad si es necesario
        
    single_point_menu_item = bet_menu_item.child_window(title="Single Point Surface Area", control_type="MenuItem")
    single_point_menu_item.click_input()

    print("Se seleccionó 'Single Point Surface Area' exitosamente.")

    time.sleep(2)
    
    secondary_window2 = app.window(title_re=f".*tab:Single Point Surface Area: file_to_open_nameonly.*")
    main_window.right_click_input()  # Click derecho en la ventana de reporte
    time.sleep(1)

    # Buscar el item "Export to .CSV" dentro del menú
    savecsv_menu_item = context_menu.child_window(title="Export to .CSV", control_type="MenuItem")
    savecsv_menu_item.click_input()
    print("Se seleccionó 'Export to .CSV' exitosamente.")

    time.sleep(2)  # Ajusta según sea necesario
    
    csv_dialog = app.window(class_name="#32770")
    #csv_dialog.wait('visible', timeout=10)  # Esperar hasta que sea visible

    print("llego hasta aqui")
    edit_box = csv_dialog.child_window( control_type="Edit",found_index=0) #auto_id="1148"
    if edit_box.exists():
        print("Existe el dialogo de texto")
        edit_box.type_keys(exported_file_path, with_spaces=True)
    else:
       raise Exception("Campo de texto para la ruta no encontrado.")

    # Intentar localizar el botón con título "Guardar" o "Save"
    csv_button = csv_dialog.child_window(control_type="Button", title="Guardar") \
    if csv_dialog.child_window(control_type="Button", title="Guardar").exists() \
    else csv_dialog.child_window(control_type="Button", title="Save")
    if csv_button.exists():
      print("Existe el boton")
      csv_button.click_input()
    else:
       raise Exception("Botón 'Guardar' no encontrado.")
       print("Archivo exportado exitosamente.")
    
    
    
    #"""Exporta el reporte desde NovaWin."""
    #try:
    #    main_window.right_click_input()
    #    time.sleep(1)
        
        # Interactuar con menú contextual
    #    context_menu = main_window.child_window(title_re=".*Context.*")
    #    export_option = context_menu.child_window(title="Export to .CSV", control_type="MenuItem")
    #    export_option.click_input()

        # Guardar el archivo
    #    csv_dialog = main_window.child_window(class_name="#32770")
    #    interactuar_con_cuadro_dialogo(csv_dialog, ruta_exportacion)
    #    print("Archivo exportado exitosamente.")
   except Exception as e:
        print(f"Error durante la exportación: {e}")
        traceback.print_exc()

# --- INICIO DEL SCRIPT ---
def main(path_qps,path_csv,path_novawin):
 try:
     ## Solicitar rutas al usuario
     #path_qps = input("Ruta para extraer archivos .qps: ").strip()
     #path_csv = input("Ruta para guardar reportes: ").strip()
     #path_novawin = input("Ruta de NovaWin: ").strip()
 
     # Validar rutas
    # if not os.path.isdir(path_qps):
    #     raise FileNotFoundError(f"La ruta para archivos .qps no existe: {path_qps}")
    #if not os.path.isdir(path_csv):
    #    raise FileNotFoundError(f"La ruta para guardar reportes no existe: {path_csv}")
     novawin_exe = os.path.join(path_novawin, "NovaWin.exe")
     if not os.path.isfile(novawin_exe):
         raise FileNotFoundError(f"No se encontró NovaWin.exe en: {path_novawin}")

     # Buscar archivo .qps
     archivo_qps = encontrar_primer_qps(path_qps)
     if not archivo_qps:
         raise FileNotFoundError("No se encontró ningún archivo .qps.")

     #Generar ruta del reporte
     timestamp = datetime.now().strftime("%Y%m%d_%H-%M-%S")
     ruta_reporte = os.path.join(path_csv, f"reporte_{timestamp}.csv")

     # Iniciar NovaWin
     app, main_window = inicializar_novawin(novawin_exe)

     # Abrir archivo .qps
     seleccionar_menu(main_window, "File->Open")
     dialog = app.window(class_name="#32770")
     interactuar_con_cuadro_dialogo(dialog, archivo_qps)

     # Exportar reporte
     exportar_reporte(main_window, ruta_reporte,app)

 except Exception as general_error:
     print(f"Se produjo un error: {general_error}")
     traceback.print_exc()
