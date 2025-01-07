import openpyxl
import os
import csv
import traceback
from datetime import datetime
from pywinauto import Application, findwindows
import time
from novawinmng import manejar_novawin, leer_csv_y_crear_dataframe, guardar_dataframe_en_ini

def generar_nombre_unico(ruta_base):
    base, ext = os.path.splitext(ruta_base)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{base}_{timestamp}{ext}"

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
        secondary_window2 = app.window(title_re=f".*tab: Pore Size Distribution: file_to_open_nameonly.*")
        main_window.right_click_input()
        time.sleep(1)

        savecsv_menu_item = context_menu.child_window(title="Export to .CSV", control_type="MenuItem")
        savecsv_menu_item.click_input()
        print("Se seleccionó 'Export to .CSV' exitosamente.")

        time.sleep(2)
        csv_dialog = app.window(class_name="#32770")

        print("Identificadores de controles de csv_dialog:")
        csv_dialog.print_control_identifiers()
        edit_box = csv_dialog.child_window(control_type="Edit", found_index=0)
        if edit_box.exists():
            print("Existe el cuadro de texto para la ruta")
            # Generar ruta con nombre único
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            ruta_exportacion = os.path.join(
                ruta_exportacion, f"qps_{timestamp}.csv"
            )
            ruta_exportacion = ruta_exportacion.replace("/", "\\")  # Corregir barras
            #edit_box.set_edit_text("hola")
            edit_box.type_keys("hola", with_spaces=True)
        else:
            raise Exception("Campo de texto para la ruta no encontrado.")
        print(ruta_exportacion)
        csv_button = csv_dialog.child_window(control_type="Button", title="Guardar") \
            if csv_dialog.child_window(control_type="Button", title="Guardar").exists() \
            else csv_dialog.child_window(control_type="Button", title="Save")
        
        if csv_button.exists():
            print("Existe el botón")
            csv_button.click_input()
            print(f"Archivo exportado exitosamente a: {ruta_exportacion}")
        else:
            raise Exception("Botón 'Guardar' no encontrado.")

        # Agregar contenido del CSV al archivo Excel
        agregar_csv_a_excel(ruta_exportacion, "reporte.xlsx")

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
        exportar_reporte(main_window, path_csv, app)
        print(f"Reporte exportado a: {path_csv}")

        # Crear DataFrame y guardar
        dataframe = leer_csv_y_crear_dataframe(path_csv)
        guardar_dataframe_en_ini(dataframe, "dataframe.ini")
        dataframe.to_excel("reporte.xlsx", index=False, engine="openpyxl")

        print("Proceso completado exitosamente.")

    except Exception as e:
        print(f"Error en hk_main: {e}")
        traceback.print_exc()
