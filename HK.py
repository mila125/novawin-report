import openpyxl
import os
import csv
import traceback
from datetime import datetime
from pywinauto import Application, findwindows
import time
from novawinmng import manejar_novawin, leer_csv_y_crear_dataframe,agregar_csv_a_plantilla_excel, guardar_dataframe_en_ini
from pywinauto.keyboard import send_keys

def generar_nombre_unico(base_path):
    # Normalizar las barras a formato Unix (/)
    base_path = base_path.replace("\\", "/")
    
    if not base_path.endswith("hk.csv"):
        base_path += "hk.csv"
    
    counter = 1
    while os.path.exists(base_path):
        name, ext = os.path.splitext(base_path)
        base_path = f"{name}_{counter}{ext}"
        counter += 1
    
    # Normalizar las barras de regreso a formato Windows (\)
    return base_path.replace("/", "\\")
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

        HK_method_menu_item = app.window(best_match="Tables").child_window(title="HK method", control_type="MenuItem")
        HK_method_menu_item.click_input()
        print("Submenú 'HK method' seleccionado.")

        HK_method_menu_item = app.window(best_match="HK method")
        HK_method_menu_item.print_control_identifiers(depth=2)
        Pore_Size_Distribution_menu_item = HK_method_menu_item.child_window(title=" Pore Size Distribution", control_type="MenuItem")
        Pore_Size_Distribution_menu_item.click_input()
        print("Se seleccionó ' Pore Size Distribution' exitosamente.")

        time.sleep(2)
        secondary_window2 = app.window(title_re=f".*tab:Pore Size Distribution: file_to_open_nameonly.*")
        main_window.right_click_input()
        time.sleep(1)

        savecsv_menu_item = context_menu.child_window(title="Export to .CSV", control_type="MenuItem")
        savecsv_menu_item.click_input()
        print("Se seleccionó 'Export to .CSV' exitosamente.")

        time.sleep(2)
        csv_dialog = app.window(class_name="#32770")

        print("llegó hasta aquí")
        ruta_exportacion = generar_nombre_unico(ruta_exportacion)
        
         # Enfocar el cuadro de texto con Alt + M
        send_keys('%m')  # % representa la tecla Alt en pywinauto
        time.sleep(2)
        send_keys(ruta_exportacion)  # % representa la tecla Alt en pywinauto
        # Esperar hasta que el cuadro de texto esté enfocado
        edit_box = csv_dialog.child_window(control_type="Edit", found_index=0)
        if edit_box.exists(timeout=5):
           print("Existe el cuadro de texto para la ruta")
           
           edit_box.type_keys("hola", with_spaces=True)
        else:
           raise Exception("Campo de texto para la ruta no encontrado.")

        csv_button = csv_dialog.child_window(control_type="Button", title="Guardar") \
            if csv_dialog.child_window(control_type="Button", title="Guardar").exists() \
            else csv_dialog.child_window(control_type="Button", title="Save")
        
        if csv_button.exists():
            print("Existe el botón")
            csv_button.click_input()
            print("Archivo exportado exitosamente.")
            return ruta_exportacion
        else:
            raise Exception("Botón 'Guardar' no encontrado.")

    except Exception as e:
        print(f"Error durante la exportación: {e}")
        traceback.print_exc()

def agregar_csv_a_excel(ruta_csv, ruta_excel):
    try:
        if not os.path.exists(ruta_excel):
            # Crear un nuevo archivo de Excel si no existe
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Reporte"
            wb.save(ruta_excel)
            print(f"Archivo Excel creado: {ruta_excel}")

        # Abrir el archivo de Excel existente
        wb = openpyxl.load_workbook(ruta_excel)
        if "Reporte" not in wb.sheetnames:
            ws = wb.create_sheet(title="Reporte")
        else:
            ws = wb["Reporte"]

        # Leer el archivo CSV
        with open(ruta_csv, "r", encoding="utf-8") as file:
            reader = csv.reader(file)
            for row in reader:
                ws.append(row)  # Agregar cada fila al Excel

        wb.save(ruta_excel)
        print(f"Reporte guardado exitosamente en: {ruta_excel}")

    except Exception as e:
        print(f"Error al agregar datos al archivo Excel: {e}")
        traceback.print_exc()

def hk_main(path_qps, path_csv, path_novawin):
    print("Inicio de hk_main")
    try:
        # Inicializar y manejar NovaWin
        app, main_window = manejar_novawin(path_novawin, path_qps)

        # Exportar reporte
        ruta_csv=exportar_reporte(main_window, path_csv, app)
        if not os.path.exists(ruta_csv):
            raise FileNotFoundError(f"Archivo exportado no encontrado en: {ruta_csv}")
        print(f"Archivo exportado exitosamente en: {ruta_csv}")
        # Crear DataFrame y guardar
        dataframe = leer_csv_y_crear_dataframe(ruta_csv)
        print(dataframe)
        agregar_csv_a_plantilla_excel(ruta_csv, path_csv)
        guardar_dataframe_en_ini(dataframe, path_csv+"dataframe.ini")
    
        #dataframe.to_excel("reporte.xlsx", index=False, engine="openpyxl")

        print("Proceso completado exitosamente.")

    except Exception as e:
        print(f"Error en hk_main: {e}")
        traceback.print_exc()
