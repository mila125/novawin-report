import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Ruta del archivo con datos exportados para desorción
file_path = input("Por favor, ingresa la ruta del archivo con datos de desorción (CSV o TXT): ")
output_pdf_path = input("Por favor, ingresa la ruta donde deseas guardar el gráfico de desorción (incluye el nombre y .pdf): ")

try:
    # Leer el archivo exportado
    if file_path.endswith('.csv'):
        data = pd.read_csv(file_path)
    elif file_path.endswith('.txt'):
        data = pd.read_csv(file_path, sep='\t')  # Suponiendo separadores de tabulación
    else:
        raise ValueError("El archivo debe ser .csv o .txt")

    # Mostrar los primeros datos para verificar
    print("Datos cargados correctamente:")
    print(data.head())

    # Verificar si las columnas necesarias existen (ajusta estos nombres si es necesario)
    if 'Relative Pressure' in data.columns and 'Volume @ STP' in data.columns:
        # Calcular dV(r) para desorción (cambio en volumen entre puntos consecutivos)
        data['dV(r)'] = data['Volume @ STP'].diff().fillna(0)

        # Calcular el diámetro de poro (o cualquier otro parámetro que se derive de las fórmulas de BJH)
        # Usualmente se utiliza la ecuación de BJH para obtener dV(r) en función de la presión relativa y otros parámetros.
        # Por simplicidad, en este ejemplo no se hace un cálculo detallado de BJH

        # Crear el gráfico de desorción
        plt.figure(figsize=(10, 6))
        plt.plot(data['Relative Pressure'], data['dV(r)'], marker='o', linestyle='-', color='red', label='Desorción')

        # Configurar el gráfico
        plt.title('Gráfico de Desorción - Método BJH', fontsize=16)
        plt.xlabel('Presión Relativa (P/P₀)', fontsize=14)
        plt.ylabel('dV(r)', fontsize=14)
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.legend()

        # Guardar el gráfico como PDF
        plt.savefig(output_pdf_path, format='pdf')
        print(f"Gráfico de desorción guardado exitosamente en: {output_pdf_path}")

        # Mostrar el gráfico
        plt.show()
    else:
        print("Las columnas 'Relative Pressure' y 'Volume @ STP' no están en el archivo.")
except FileNotFoundError:
    print("Archivo no encontrado. Verifica la ruta e inténtalo de nuevo.")
except Exception as e:
    print(f"Error al procesar los datos: {e}")
