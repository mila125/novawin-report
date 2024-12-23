from pywinauto import Application, findwindows
import time
from datetime import datetime
import os
import traceback
from HK import hk_main  # Asegúrate de que el módulo HK esté disponible
import threading


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

        bet_menu_item = app.window(best_match="Tables").child_window(title="BET", control_type="MenuItem")
        bet_menu_item.click_input()
        print("Submenú 'BET' seleccionado.")

        bet_menu_item = app.window(best_match="BET")
        bet_menu_item.print_control_identifiers(depth=2)
        single_point_menu_item = bet_menu_item.child_window(title="Single Point Surface Area", control_type="MenuItem")
        single_point_menu_item.click_input()
        print("Se seleccionó 'Single Point Surface Area' exitosamente.")

        time.sleep(2)
        secondary_window2 = app.window(title_re=f".*tab:Single Point Surface Area: file_to_open_nameonly.*")
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

def ejecutar_hk_main_en_hilo(ruta_exportacion, ruta_pdf, evento):
    try:
        # Esperar a que el evento sea señalado
        print("Esperando a que se complete la exportación...")
        evento.wait()  # Bloquea hasta que el evento sea activado
        print(f"Ejecutando hk_main con ruta_exportacion: {ruta_exportacion}, ruta_pdf: {ruta_pdf}")
        hk_main(ruta_exportacion, ruta_pdf)
        print("hk_main finalizado.")
    except Exception as e:
        print(f"Error en hk_main: {e}")
        traceback.print_exc()

# Lógica de NovaWin, con exportación que señala el evento
def manejar_novawin(evento, path_novawin, archivo_qps, ruta_reporte):
    try:
        novawin_exe = os.path.join(path_novawin, "NovaWin.exe")
        if not os.path.isfile(novawin_exe):
            raise FileNotFoundError(f"No se encontró NovaWin.exe en: {path_novawin}")

        # Iniciar NovaWin
        app, main_window = inicializar_novawin(novawin_exe)

        # Interactuar con NovaWin para abrir archivo y exportar reporte
        seleccionar_menu(main_window, "File->Open")
        dialog = app.window(class_name="#32770")
        interactuar_con_cuadro_dialogo(dialog, archivo_qps)

        # Exportar reporte
        exportar_reporte(main_window, ruta_reporte, app)

        print(f"Archivo exportado a: {ruta_reporte}")

        # Señalizar que el archivo está listo
        evento.set()  # Activar el evento para indicar que la exportación ha terminado

    except Exception as e:
        print(f"Error al manejar NovaWin: {e}")
        traceback.print_exc()

def main(path_qps, path_csv, path_novawin, ruta_pdf):
    try:
        # Buscar archivo .qps
        archivo_qps = next((os.path.join(path_qps, f) for f in os.listdir(path_qps) if f.endswith(".qps")), None)
        if not archivo_qps:
            raise FileNotFoundError("No se encontró ningún archivo .qps.")

        # Generar ruta del reporte
        timestamp = datetime.now().strftime("%Y%m%d_%H-%M-%S")
        ruta_reporte = os.path.join(path_csv, f"reporte_{timestamp}.csv")

        # Crear un evento de sincronización
        evento_csv_generado = threading.Event()

        # Ejecutar NovaWin en un hilo
        hilo_novawin = threading.Thread(
            target=manejar_novawin,
            args=(evento_csv_generado, path_novawin, archivo_qps, ruta_reporte)
        )
        hilo_novawin.daemon = True
        hilo_novawin.start()

        # Ejecutar hk_main en otro hilo, pero solo después del evento
        hilo_hk_main = threading.Thread(
            target=ejecutar_hk_main_en_hilo,
            args=(ruta_reporte, ruta_pdf, evento_csv_generado)
        )
        hilo_hk_main.daemon = True
        hilo_hk_main.start()

        print("Tareas de NovaWin y hk_main iniciadas en hilos separados.")

        # Opcional: esperar a que ambos hilos terminen
        hilo_novawin.join()
        hilo_hk_main.join()

    except Exception as general_error:
        print(f"Se produjo un error: {general_error}")
        traceback.print_exc()