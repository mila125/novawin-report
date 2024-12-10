from pywinauto import Application
import time

# Ruta al ejecutable de NovaWin
novawin_path = r"C:\Quantachrome Instruments\NovaWin\NovaWin.exe"
# Ruta del archivo que deseas abrir
file_to_open = r"C:\Users\6lady\OneDrive\Escritorio\test.qps"
# Ruta donde deseas exprortar los archivos planos
exported_file_path = r"C:\Users\6lady\OneDrive\Escritorio\export.csv"
# Iniciar la aplicación
app = Application(backend="uia").start(novawin_path)

# Esperar a que la ventana principal se cargue
time.sleep(5)  # Ajusta este tiempo según sea necesario

# Conectar con la ventana principal de NovaWin
main_window = app.window(title_re=".*NovaWin.*")

# Navegar al menú 'Single Point Surface Area'
# Esto dependerá del diseño de la aplicación. Asegúrate de identificar los controles.
main_window.menu_select("File->Open")  #Tables->BET->Single Point Surface Area

# Esperar a que se abra el diálogo de selección de archivo
time.sleep(2)

# Conectar con el cuadro de diálogo de selección de archivo
open_dialog = app.window(title_re=".*Abrir.*")

# Escribir la ruta del archivo
open_dialog.Edit.type_keys(file_to_open, with_spaces=True)


# Seleccionar el botón específico usando índices
open_button = open_dialog.child_window(title="Abrir", control_type="Button", found_index=3)
open_button.click_input()

# Confirmar que el archivo se abre correctamente
print("Archivo abierto exitosamente en NovaWin.")
# Realizar la acción para guardar como .csv

# Buscar la ventana por su título
secondary_window = app.window(title_re=".*Quantachrome™ NovaWin - [graph:Isotherm :   Linear: test.qps].*")

# Verificar si la ventana está visible
if main_window.exists():
    main_window.set_focus()
    
    # Imprimir controles para verificar la jerarquía
    main_window.print_control_identifiers()

    # Realizar clic derecho
    main_window.right_click_input(coords=(658, 331))  # Ajusta las coordenadas según sea necesario
    
    # Esperar un segundo para que aparezca el menú
    time.sleep(1)

    # Buscar el menú contextual y el item "Tables"
    try:
        context_menu = app.window(title_re=".*Context.*")  # Modificar título si es necesario
        tables_menu_item = context_menu.child_window(title="Tables", control_type="MenuItem")
        tables_menu_item.click_input()
        print("Se seleccionó el menú 'Tables'.")
        # Enumerar subitems del menú "Tables"
        
        #tables_menu_item.print_control_identifiers()  # Ayuda a identificar "BET"
        #main_window.print_control_identifiers(depth=2)  # Ajusta la profundidad si es necesario

        app_top_window = app.window(best_match="Tables")  # Reemplaza "Tables" si es necesario
        app_top_window.print_control_identifiers()

 
        time.sleep(1)
        # Expandir el submenú "BET"
        bet_menu_item = app_top_window.child_window(title="BET", control_type="MenuItem")
        bet_menu_item.click_input()
        print("Se seleccionó 'BET' exitosamente.")
        time.sleep(1)

        # Seleccionar "Single Point Surface Area"
        bet_menu_item = app.window(best_match="BET")
        bet_menu_item.print_control_identifiers(depth=2)  # Aumenta el nivel de profundidad si es necesario
        
        single_point_menu_item = bet_menu_item.child_window(title="Single Point Surface Area", control_type="MenuItem")
        single_point_menu_item.click_input()

        print("Se seleccionó 'Single Point Surface Area' exitosamente.")
        secondary_window2 = app.window(title_re=".*tab:Single Point Surface Area: test.qps.*")
        main_window.right_click_input()  # Click derecho en la ventana de reporte
        time.sleep(1)

        # Buscar el menú contextual y el item "Export to .CSV"
       
        context_menu.print_control_identifiers()  # Verificar qué controles hay en el menú contextual

        # Buscar el item "Export to .CSV" dentro del menú
        savecsv_menu_item = context_menu.child_window(title="Export to .CSV", control_type="MenuItem")
        savecsv_menu_item.click_input()
        print("Se seleccionó 'Export to .CSV' exitosamente.")

        time.sleep(2)  # Ajusta según sea necesario

        # Verificar todas las ventanas activas
        for window in app.windows():
          print(window)
          print(window.window_text())

        # Intentar conectar con el diálogo de guardar
        csv_dialog = app.window(title_re=".*Name  File.*")
        csv_dialog.wait('visible', timeout=10)  # Esperar hasta que sea visible
        csv_dialog.print_control_identifiers()

        # Buscar el botón Guardar
        csv_button = csv_dialog.child_window(title="Guardar", control_type="Button")
        csv_button.click_input()
        print("Archivo exportado exitosamente.")

  
        
    except Exception as e:
       print("Error al seleccionar:", e)

else:
    print("La ventana '.*Quantachrome™ NovaWin - [graph:Isotherm :   Linear: test.qps].*")

    
