from pywinauto import Application
import time
from datetime import datetime
import os

# Función para buscar el primer archivo .qps
def encontrar_primer_qps(directorio):
    # Recorrer el directorio y sus subdirectorios
    for archivo in os.listdir(directorio):
        # Verificar si el archivo tiene extensión .qps
        if archivo.endswith(".qps"):
            # Retornar la ruta completa del primer archivo encontrado
            return os.path.join(directorio, archivo)
    # Si no se encuentra ningún archivo .qps
    return None

#En Python, puedes leer una cadena de texto desde el teclado utilizando la función input(). Aquí tienes un ejemplo básico:

path_qps = input("Por favor, ingresa ruta para extraer archivos : ")

path_csv = input("Por favor, ingresa ruta donde se van aguardar los reportes : ")

path_novawin = input("Por favor, ingresa  donde se encuentra NovaWin : ")

# Ruta al ejecutable de NovaWin
novawin_path = path_novawin+r"\NovaWin.exe"  #C:\Quantachrome Instruments\NovaWin\
# Ruta del archivo que deseas abrir

# Buscar el primer archivo .qps en el directorio
file_to_open_nameonly = encontrar_primer_qps(path_qps)

if file_to_open_nameonly:
    print(f"Se encontró el archivo .qps: {file_to_open_nameonly}")
    # Aquí puedes abrir el archivo o continuar con el procesamiento
else:
    print("No se encontró ningún archivo .qps en el directorio especificado.")
file_to_open = os.path.join(path_qps, file_to_open_nameonly)  # Esto genera la ruta completa correctamente #file_to_open =f"{path_qps}\\{file_to_open}"  #C:\Users\6lady\OneDrive\Escritorio\    path_qps+"\"file_to_open   

print(f"Asi se quedo file to open : {file_to_open}")
# Ruta donde deseas exprortar los archivos planos
# Obtener la fecha y hora actual
ahora = datetime.now()

# Convertir la fecha y hora a un string con un formato específico
fecha_str = ahora.strftime("%Y%m%d_%H-%M-%S")  # Formato válido para nombres de archivo

print("Fecha actual:", fecha_str)

# Concatenar la ruta, fecha/hora, y extensión del archivo
exported_file_path = f"{path_csv}\\FFHA_{fecha_str}.csv"

print("Ruta completa del archivo a guardar:", exported_file_path)

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
secondary_window = app.window(title_re=f".*Quantachrome™ NovaWin - [graph:Isotherm :   Linear: {file_to_open_nameonly}].*")

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
    
        app_top_window = app.window(best_match="Tables")  # Reemplaza "Tables" si es necesario
        app_top_window.print_control_identifiers()

        time.sleep(1)
        # Expandir el submenú "Fractal Dimension Methods"
        fdm_menu_item = app_top_window.child_window(title="Fractal Dimension Methods", control_type="MenuItem")
        fdm_menu_item.click_input()
        print("Se seleccionó 'Fractal Dimension Methods' exitosamente.")
        time.sleep(1)

        # Seleccionar "NK Method Fractal Dimension(Absorption)"
        fdm_menu_item = app.window(best_match="Fractal Dimension Methods")
        fdm_menu_item.print_control_identifiers(depth=2)  # Aumenta el nivel de profundidad si es necesario
        
        ffha_menu_item = fdm_menu_item.child_window(title="FHH Method Fractal Dimension(Adsorption )", control_type="MenuItem")
        ffha_menu_item.click_input()

        print("Se seleccionó 'FHH Method Fractal Dimension(Adsorption )' exitosamente.")
        secondary_window2 = app.window(title_re=f".*tab:FHH Method Fractal Dimension(Adsorption ): file_to_open_nameonly.*")
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

        # Localizar el campo de texto dentro del ComboBox
        edit_box = csv_dialog.child_window(auto_id="1148", control_type="Edit")


        # Escribir la ruta del archivo
        edit_box.type_keys(exported_file_path, with_spaces=True)

        # Escribir la ruta del archivo
        csv_dialog.Edit.type_keys(exported_file_path, with_spaces=True) #save_dialog

        # Buscar el botón Guardar
        csv_button = csv_dialog.child_window(title="Guardar", control_type="Button")
        csv_button.click_input()
        print("Archivo exportado exitosamente.")

  
        
    except Exception as e:
       print("Error al seleccionar:", e)

else:
    print(f"La ventana '.*Quantachrome™ NovaWin - [graph:Isotherm :   Linear: {file_to_open_nameonly}].*")

