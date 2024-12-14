import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Ruta del archivo con datos
file_path = input("Por favor, ingresa la ruta del archivo con datos DFT (CSV o TXT): ")
output_pdf_path = input("Por favor, ingresa la ruta completa para guardar el gráfico en PDF (incluye el nombre y .pdf): ")

try:
    # Leer el archivo con datos
    if file_path.endswith('.csv'):
        data = pd.read_csv(file_path)
    elif file_path.endswith('.txt'):
        data = pd.read_csv(file_path, sep='\t')  # Suponiendo separadores de tabulación
    else:
        raise ValueError("El archivo debe ser .csv o .txt")

    # Mostrar los primeros datos para verificar
    print("Datos cargados correctamente:")
    print(data.head())

    # Verificar si las columnas necesarias están presentes
    if 'Relative Pressure' in data.columns and 'Volume @ STP' in data.columns:
        # Extraer los datos
        x = data['Relative Pressure']
        y = data['Volume @ STP']

        # Aplicar la Transformada de Fourier Discreta (DFT)
        dft = np.fft.fft(y)  # Transformada
        frequencies = np.fft.fftfreq(len(y), d=(x.iloc[1] - x.iloc[0]))  # Frecuencias asociadas

        # Graficar la magnitud del DFT
        plt.figure(figsize=(10, 6))
        plt.plot(frequencies, np.abs(dft), color='blue', label='DFT Magnitud')
        plt.title('Transformada de Fourier Discreta (DFT)', fontsize=16)
        plt.xlabel('Frecuencia', fontsize=14)
        plt.ylabel('Magnitud', fontsize=14)
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.legend()

        # Guardar el gráfico como PDF
        plt.savefig(output_pdf_path, format='pdf')
        print(f"Gráfico guardado exitosamente en: {output_pdf_path}")

        # Mostrar el gráfico
        plt.show()
    else:
        print("Las columnas 'Relative Pressure' y 'Volume @ STP' no están en el archivo.")
except FileNotFoundError:
    print("Archivo no encontrado. Verifica la ruta e inténtalo de nuevo.")
except Exception as e:
    print(f"Error al procesar los datos: {e}")

