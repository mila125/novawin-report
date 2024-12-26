import matplotlib.pyplot as plt
import pandas as pd

def nkagraph_main(file_path,output_pdf_path):
 print("Hola desde NKAGRAPH:main")
 try:
     # Leer los datos desde el archivo CSV
     datos = pd.read_csv(ruta_csv)

     # Verificar si las columnas necesarias están presentes
     if 'Radius of curvature' in datos.columns and 'Vapor-Liquid Intrface Area' in datos.columns:
         radius_of_curvature = datos['Radius of curvature']
         vapor_liquid_interface_area = datos['Vapor-Liquid Intrface Area']

         # Crear la gráfica
         plt.figure(figsize=(8, 6))
         plt.plot(radius_of_curvature, vapor_liquid_interface_area, marker='o', linestyle='-', color='green', label='Adsorption - NK')

         # Configurar etiquetas, título y leyenda
         plt.xlabel('Radius of Curvature')
         plt.ylabel('Vapor-Liquid Interface Area')
         plt.title('Adsorption - NK Method')
         plt.grid(True)
         plt.legend()

         # Guardar y mostrar la gráfica
         plt.savefig(ruta_salida_pdf, format='pdf')
         plt.show()

         print(f"Gráfica guardada correctamente como '{ruta_salida_pdf}'.")

     else:
         print("Error: El archivo CSV no contiene las columnas necesarias ('Radius of curvature', 'Vapor-Liquid Intrface Area').")

 except FileNotFoundError:
     print("Error: No se encontró el archivo CSV en la ruta especificada.")
 except Exception as e:
     print(f"Se produjo un error: {e}")
