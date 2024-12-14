import pandas as pd
import matplotlib.pyplot as plt

# Ruta del archivo con datos exportados
file_path = input("Por favor, ingresa la ruta del archivo con datos exportados (CSV o TXT): ")
output_pdf_path = input("Por favor, ingresa la ruta completa donde deseas guardar el gráfico en PDF (incluye el nombre y .pdf): ")

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

    # Verificar si las columnas "Pressure" y "Volume" existen
    if 'Relative Pressure' in data.columns and 'Volume @ STP' in data.columns:
        # Crear el gráfico
        plt.figure(figsize=(10, 6))
        plt.plot(data['Relative Pressure'], data['Volume @ STP'], marker='o', linestyle='-', color='blue', label='Volume')

        # Configurar el gráfico
        plt.title('Gráfico de Volumen', fontsize=16)
        plt.xlabel('Presión (Pressure)', fontsize=14)
        plt.ylabel('Volumen (Volume)', fontsize=14)
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.legend()

        # Guardar el gráfico como PDF
        plt.savefig(output_pdf_path, format='pdf')
        print(f"Gráfico guardado exitosamente en: {output_pdf_path}")

        # Mostrar el gráfico
        plt.show()
    else:
        print("Las columnas 'Pressure' y 'Volume' no están en el archivo.")
except FileNotFoundError:
    print("Archivo no encontrado. Verifica la ruta e inténtalo de nuevo.")
except Exception as e:
    print(f"Error al procesar los datos: {e}")
