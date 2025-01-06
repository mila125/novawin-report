from pywinauto import Application, findwindows
import time
from datetime import datetime
import os
import traceback
import threading
import subprocess
import pandas as pd
import configparser
def generar_nombre_unico(ruta):
    """
    Genera un nombre único para el archivo añadiendo un índice si ya existe un archivo con el mismo nombre.
    """
    directorio, nombre_archivo = os.path.split(ruta)
    nombre, extension = os.path.splitext(nombre_archivo)
    indice = 1

    while os.path.exists(ruta):
        ruta = os.path.join(directorio, f"{nombre}_{indice}{extension}")
        indice += 1

    return ruta


def guardar_dataframe_en_ini(df, archivo_ini):
    """
    Guarda un DataFrame en formato .ini.
    
    Args:
        df (pd.DataFrame): DataFrame a guardar.
        archivo_ini (str): Ruta del archivo .ini de destino.
    """
    config = configparser.ConfigParser()
    
    for columna in df.columns:
        # Crear una sección para cada columna
        config[columna] = {}
        for i, valor in enumerate(df[columna]):
            # Agregar cada fila como un par clave-valor
            config[columna][f"fila_{i}"] = str(valor)
    
    with open(archivo_ini, 'w') as archivo:
        config.write(archivo)
    print(f"DataFrame guardado exitosamente en {archivo_ini}")
def leer_csv_y_crear_dataframe(ruta_csv):
    """
    Lee un archivo CSV y lo convierte en un DataFrame de pandas.

    Args:
        ruta_csv (str): Ruta al archivo CSV.

    Returns:
        pd.DataFrame: DataFrame creado a partir del CSV.
    """
    try:
        # Leer el archivo CSV
        df = pd.read_csv(ruta_csv)
        print("DataFrame creado exitosamente.")
        return df
    except FileNotFoundError:
        print(f"Error: El archivo {ruta_csv} no existe.")
    except Exception as e:
        print(f"Error al leer el archivo CSV: {e}")

def crear_ventana_novarep_ide():
    """Simula la creación de la ventana principal de novarep_ide."""
    try:
        print("Recreando ventana novarep_ide...")
        # Aquí deberías colocar el código para inicializar o mostrar novarep_ide
        # Por ejemplo: app_novarep = Application(backend="uia").start(path_novarep)
        # print("Ventana novarep_ide creada.")
    except Exception as e:
        print(f"Error al crear la ventana novarep_ide: {e}")
        raise
# Inicializar NovaWin
def inicializar_novawin(novawin_path):
    try:
        app = Application(backend="uia").start(novawin_path)
        time.sleep(5)  # Esperar a que se cargue la ventana principal
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

def exportar_reporte(main_window, ruta_exportacion, app):
    try:
        # Imprimir detalles de los controles y ventanas hijas de main_window
        all_controls = main_window.children()
        main_window.print_control_identifiers()

        # Buscar el control por su clase
        graph_view_window = main_window.child_window(class_name="TGraphViewWindow")

        if graph_view_window.exists():
            print("Componente 'TGraphViewWindow' encontrado.")
            graph_view_window.print_control_identifiers()
            graph_view_window.right_click_input()
            time.sleep(1)
        else:
            print("No se encontró el componente 'TGraphViewWindow'.")

        context_menu = app.window(title_re=".*Context.*")
        tables_menu_item = context_menu.child_window(title="Tables", control_type="MenuItem")
        tables_menu_item.click_input()
        print("Menú 'Tables' seleccionado.")

        bet_menu_item = app.window(best_match="Tables").child_window(title="DFT method", control_type="MenuItem")
        bet_menu_item.click_input()
        print("Submenú 'HK method' seleccionado.")

        bet_menu_item = app.window(best_match="HK method")
        bet_menu_item.print_control_identifiers(depth=2)
        single_point_menu_item = bet_menu_item.child_window(title=" Pore Size Distribution", control_type="MenuItem")
        single_point_menu_item.click_input()
        print("Se seleccionó ' Pore Size Distribution' exitosamente.")

        time.sleep(2)
        secondary_window2 = app.window(title_re=f".*tab: Pore Size Distribution: file_to_open_nameonly.*")
        main_window.right_click_input()
        time.sleep(1)

        savecsv_menu_item = context_menu.child_window(title="Export to .CSV", control_type="MenuItem")
        savecsv_menu_item.click_input()
        print("Se seleccionó 'Export to .CSV' exitosamente.")

        time.sleep(2)
        csv_dialog = app.window(class_name="#32770")

        print("llegó hasta aquí")
        edit_box = csv_dialog.child_window(control_type="Edit", found_index=0)
        if edit_box.exists():
            print("Existe el cuadro de texto para la ruta")
            ruta_exportacion = generar_nombre_unico(ruta_exportacion)
            edit_box.type_keys(ruta_exportacion, with_spaces=True)
        else:
            raise Exception("Campo de texto para la ruta no encontrado.")

        csv_button = csv_dialog.child_window(control_type="Button", title="Guardar") \
            if csv_dialog.child_window(control_type="Button", title="Guardar").exists() \
            else csv_dialog.child_window(control_type="Button", title="Save")
        
        if csv_button.exists():
            print("Existe el botón")
            csv_button.click_input()
            print("Archivo exportado exitosamente.")
        else:
            raise Exception("Botón 'Guardar' no encontrado.")


    except Exception as e:
        print(f"Error durante la exportación: {e}")
        traceback.print_exc()

# Lógica de NovaWin, con exportación que señala el evento
def manejar_novawin( path_novawin, archivo_qps, path_csv):
    try:

        # Iniciar NovaWin
        app,main_window = inicializar_novawin(path_novawin)

        # Interactuar con NovaWin para abrir archivo y exportar reporte
        seleccionar_menu(main_window, "File->Open")
        dialog = app.window(class_name="#32770")
        interactuar_con_cuadro_dialogo(dialog, archivo_qps)

        # Exportar reporte
        exportar_reporte(main_window, path_csv, app)

        print(f"Archivo exportado a: {path_csv}")
        
        # Intentar cerrar la ventana principal de forma elegante
        try:
            main_window.close()
            print("Ventana principal de NovaWin cerrada exitosamente.")
        except Exception as e:
            print(f"No se pudo cerrar la ventana principal de forma elegante: {e}. Intentando forzar el cierre.")
            app.kill()  # Forzar el cierre del proceso de NovaWin

    except Exception as e:
        print(f"Error al manejar NovaWin: {e}")
        traceback.print_exc()

def dft_main(path_qps, path_csv, path_novawin):
    print("Hola desde novarep:main")
    print(path_qps)
    try:
  
        # Ejecutar NovaWin en un hilo
        hilo_novawin = threading.Thread(
            target=manejar_novawin,
            args=( path_novawin, path_qps, path_csv)
        )
        hilo_novawin.daemon = True
        hilo_novawin.start()

        # Opcional: esperar a que todos los hilos terminen
        hilo_novawin.join()
        
        # Crear DataFrame a partir del archivo exportado
        dataframe = leer_csv_y_crear_dataframe(path_csv)
        print(dataframe.head())  # Imprime las primeras filas para verificar
        guardar_dataframe_en_ini(dataframe, "dataframe.ini")
        # Cerrar NovaWin
        dataframe.to_excel("reporte.xlsx", index=False, engine='openpyxl')
        
        subprocess.run(["python", "novarep_ide.py"])
    except Exception as general_error:
        print(f"Se produjo un error: {general_error}")
        traceback.print_exc()

