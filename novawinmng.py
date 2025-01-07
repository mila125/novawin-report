from pywinauto import Application
import time
import os
import traceback
import pandas as pd
import configparser

def manejar_novawin(path_novawin, archivo_qps):
    try:
        # Invertir las barras en la ruta del archivo
        archivo_qps = archivo_qps.replace("/", "\\")  # Reemplazar barras normales por barras invertidas
        
        # O usar normpath para normalizar la ruta según el sistema operativo
        archivo_qps = os.path.normpath(archivo_qps)

        # Inicializar NovaWin
        app, main_window = inicializar_novawin(path_novawin)

        # Interactuar con NovaWin
        seleccionar_menu(main_window, "File->Open")
        dialog = app.window(class_name="#32770")
        interactuar_con_cuadro_dialogo(dialog, archivo_qps)

        return app, main_window

    except Exception as e:
        print(f"Error al manejar NovaWin: {e}")
        traceback.print_exc()
        raise

def inicializar_novawin(path_novawin):
    try:
        app = Application(backend="uia").start(path_novawin)
        time.sleep(5)  # Esperar que se cargue NovaWin
        main_window = app.window(title_re=".*NovaWin.*")
        return app, main_window
    except Exception as e:
        print(f"Error al inicializar NovaWin: {e}")
        raise

def seleccionar_menu(window, ruta_menu):
    try:
        window.menu_select(ruta_menu)
        time.sleep(2)
    except Exception as e:
        print(f"Error al seleccionar menú '{ruta_menu}': {e}")
        raise

def interactuar_con_cuadro_dialogo(dialog, archivo):
    try:
        edit_box = dialog.child_window(class_name="Edit")
        edit_box.set_edit_text(archivo)
        open_button = dialog.child_window(class_name="Button", found_index=0)
        open_button.click_input()
    except Exception as e:
        print(f"Error al interactuar con el cuadro de diálogo: {e}")
        raise

def exportar_reporte(main_window, path_csv, app):
    try:
        # Buscar ventana de contexto y realizar exportación
        main_window.right_click_input()
        time.sleep(1)

        context_menu = app.window(title_re=".*Context.*")
        savecsv_menu_item = context_menu.child_window(title="Export to .CSV", control_type="MenuItem")
        savecsv_menu_item.click_input()

        time.sleep(2)
        csv_dialog = app.window(class_name="#32770")
        edit_box = csv_dialog.child_window(control_type="Edit", found_index=0)
        edit_box.set_edit_text(path_csv)
        save_button = csv_dialog.child_window(title="Guardar", control_type="Button")
        save_button.click_input()

        print("Archivo exportado exitosamente.")
    except Exception as e:
        print(f"Error durante la exportación: {e}")
        traceback.print_exc()
        raise

def leer_csv_y_crear_dataframe(ruta_csv):
    try:
        return pd.read_csv(ruta_csv)
    except FileNotFoundError:
        print(f"Archivo CSV no encontrado: {ruta_csv}")
        raise
    except Exception as e:
        print(f"Error al leer CSV: {e}")
        raise

def guardar_dataframe_en_ini(df, archivo_ini):
    try:
        config = configparser.ConfigParser()
        for columna in df.columns:
            config[columna] = {f"fila_{i}": str(valor) for i, valor in enumerate(df[columna])}
        with open(archivo_ini, 'w') as archivo:
            config.write(archivo)
        print(f"DataFrame guardado en {archivo_ini}")
    except Exception as e:
        print(f"Error al guardar INI: {e}")
        raise