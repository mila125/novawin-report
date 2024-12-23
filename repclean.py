import pandas as pd
import os
import numpy as np  # Para cálculos matemáticos como raíz cuadrada

# Función para buscar el primer archivo .csv
def encontrar_primer_csv(directorio):
    for archivo in os.listdir(directorio):
        if archivo.endswith(".csv"):
            return os.path.join(directorio, archivo)
    return None
def main(path_csv,output_file):
 # Solicitar la ruta al usuario
 path_csv = input("Por favor, ingresa la ruta donde están los reportes: ")

 if not os.path.isdir(path_csv):
     print("La ruta ingresada no es válida.")
     exit()

 file = encontrar_primer_csv(path_csv)

 if file is None:
     print("No se encontró ningún archivo con extensión .csv en el directorio.")
     exit()

 # Leer el archivo CSV
 try:
     df = pd.read_csv(file)
 except Exception as e:
     print(f"Error al leer el archivo {file}: {e}")
     exit()

 columnas_necesarias = ["Relative Pressure", "Surf. Area"]

 if not all(col in df.columns for col in columnas_necesarias):
     print("El archivo no contiene las columnas necesarias.")
     exit()

 df_filtrado = df[columnas_necesarias]

 try:
     rango_min = float(input("Ingresa el valor mínimo del rango para 'Relative Pressure': "))
     rango_max = float(input("Ingresa el valor máximo del rango para 'Relative Pressure': "))
 except ValueError:
     print("Por favor, ingresa valores numéricos válidos para el rango.")
     exit()

 # Filtrar filas basadas en el rango de la columna 'Relative Pressure'
 df_filtrado_rango = df_filtrado.loc[
     (df_filtrado["Relative Pressure"] >= rango_min) & 
     (df_filtrado["Relative Pressure"] <= rango_max)
 ]

 # Calcular el promedio de la columna 'Surf. Area'
 promedio_surf_area = df_filtrado_rango["Surf. Area"].mean()

 # Crear una nueva columna con el promedio solo en la primera fila
 df_filtrado_rango["Promedio Surf. Area"] = None  # Inicializar la columna con valores vacíos
 df_filtrado_rango.at[0, "Promedio Surf. Area"] = promedio_surf_area  # Escribir el promedio solo en la primera fila

 # Número de elementos (tamaño de la muestra)
 n = len(df_filtrado_rango)

 if n == 0:
     print("No hay datos dentro del rango especificado.")
     exit()

 # Calcular la desviación estándar individual y el error estándar
 df_filtrado_rango["Desviación Estándar Individual"] = np.sqrt(
     (df_filtrado_rango["Surf. Area"] - promedio_surf_area) ** 2
 )

 df_filtrado_rango["Error Estándar Individual"] = (
     df_filtrado_rango["Desviación Estándar Individual"] / np.sqrt(n)
 )

 # Agregar el promedio como una nueva columna
 #df_filtrado_rango["Promedio Surf. Area"] = promedio_surf_area

 # Guardar los resultados en un nuevo archivo CSV
 output_file = os.path.join(path_csv, "filtrado_con_desviacion_error_promedio.csv")

 try:
     df_filtrado_rango.to_csv(output_file, index=False)
     print(f"Datos con promedio, desviación estándar y error estándar guardados en: {output_file}")
 except Exception as e:
     print(f"Error al guardar el archivo: {e}")



