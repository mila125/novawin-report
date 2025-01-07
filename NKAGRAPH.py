from pywinauto import Application, findwindows
import time
import os
import traceback
from novawinmng import manejar_novawin, leer_csv_y_crear_dataframe, guardar_dataframe_en_ini
 
def generar_nombre_unico(ruta_exportacion):
    # Obtener el directorio, el nombre base y la extensión
    directorio, nombre_archivo = os.path.split(ruta_exportacion)
    base, extension = os.path.splitext(nombre_archivo)
    
    # Agregar prefijo "nkagraph_" al nombre del archivo si no lo tiene
    if not base.startswith("nkagraph_"):
        base = f"hk_{base}"
    
    ruta_exportacion = os.path.join(directorio, f"{base}{extension}")

    # Generar un nombre único si el archivo ya existe
    contador = 1
    while os.path.exists(ruta_exportacion):
        ruta_exportacion = os.path.join(directorio, f"{base}_{contador}{extension}")
        contador += 1
    
    return ruta_exportacion
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

        Fractal_Dimension_Methods_menu_item = app.window(best_match="Tables").child_window(title="Fractal Dimension Methods", control_type="MenuItem")
        Fractal_Dimension_Methods_menu_item.click_input()
        print("Submenú 'Fractal Dimension Methods' seleccionado.")

        BJHD_menu_item = app.window(best_match="Fractal Dimension Methods")
        BJHD_menu_item.print_control_identifiers(depth=2)
        Adsorption_menu_item = BJHD_menu_item.child_window(title="NK Method Fractal Dimension(Adsorption )", control_type="MenuItem")
        Adsorption_menu_item.click_input()
        print("Se seleccionó NK Method Fractal Dimension(Adsorption )' exitosamente.")

        time.sleep(2)
        secondary_window2 = app.window(title_re=f".*tab: NK Method Fractal Dimension(Adsorption ): file_to_open_nameonly.*")
        main_window.right_click_input()
        time.sleep(1)

        savecsv_menu_item = context_menu.child_window(title="Export to .CSV", control_type="MenuItem")
        savecsv_menu_item.click_input()
        print("Se seleccionó 'Export to .CSV' exitosamente.")

        time.sleep(2)
        csv_dialog = app.window(class_name="#32770")

     
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

def nkagraph_main(path_qps, path_csv, path_novawin):
    print("Inicio de bjhd_main")
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
        print(f"Error en dft_main: {e}")
        traceback.print_exc()
