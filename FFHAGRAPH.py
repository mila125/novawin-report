import pandas as pd
import matplotlib.pyplot as plt

def main(file_path,output_pdf_path):
 try:
     # Cargar los datos desde el archivo CSV
     data = pd.read_csv(input_file)

     # Verificar si las columnas necesarias existen
     if 'log(log(P/Po))' in data.columns and 'log(Vads)' in data.columns:
         # Extraer las columnas necesarias
         x = data['log(log(P/Po))']
         y = data['log(Vads)']

         # Crear la gráfica
         plt.figure(figsize=(8, 6))
         plt.plot(x, y, 'o-', color='blue', label='Adsorption Data (FFH Method)')
         plt.xlabel("log(log(P/Po))", fontsize=12)
         plt.ylabel("log(Vads)", fontsize=12)
         plt.title("FFH Adsorption Isotherm", fontsize=14)
         plt.grid(True)
         plt.legend()

         # Guardar la gráfica como PDF
         plt.savefig(output_file)
         plt.show()

         print(f"Gráfica guardada exitosamente en: {output_file}")
     else:
         print("Error: Las columnas 'log(log(P/Po))' y 'log(Vads)' no están en el archivo CSV.")

 except FileNotFoundError:
     print(f"Error: No se pudo encontrar el archivo en la ruta: {input_file}")
 except Exception as e:
     print(f"Ocurrió un error: {e}")
