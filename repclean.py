import pandas as pd
import os

# Función para buscar el primer archivo .csv
def encontrar_primer_csv(directorio):
    # Recorrer el directorio y sus subdirectorios
    for archivo in os.listdir(directorio):
        if archivo.endswith(".csv"):  # Verificar si el archivo tiene extensión .csv
            return os.path.join(directorio, archivo)
    return None  # Si no se encuentra ningún archivo .csv

# Solicitar la ruta al usuario
path_csv = input("Por favor, ingresa la ruta donde están los reportes: ")

# Validar que la ruta existe
if not os.path.isdir(path_csv):
    print("La ruta ingresada no es válida.")
    exit()

# Encontrar el primer archivo .csv
file = encontrar_primer_csv(path_csv)

# Verificar que se encontró un archivo
if file is None:
    print("No se encontró ningún archivo con extensión .csv en el directorio.")
    exit()

# Leer el archivo CSV
try:
    df = pd.read_csv(file)
except Exception as e:
    print(f"Error al leer el archivo {file}: {e}")
    exit()

# Seleccionar solo las columnas necesarias
columnas_necesarias = ["Relative Pressure", "Surf. Area"]

# Verificar si las columnas existen en el DataFrame
if not all(col in df.columns for col in columnas_necesarias):
    print("El archivo no contiene las columnas necesarias.")
    exit()

df_filtrado = df[columnas_necesarias]

# Solicitar al usuario el rango de valores para filtrar
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

# Generar un nombre para el archivo de salida
output_file = os.path.join(path_csv, "limpio_filtrado_con_promedio_una_vez.csv")

# Guardar el resultado en un nuevo archivo
try:
    df_filtrado_rango.to_csv(output_file, index=False)
    print(f"Datos filtrados con el promedio escrito solo una vez en la columna guardados en: {output_file}")
except Exception as e:
    print(f"Error al guardar el archivo filtrado: {e}")



