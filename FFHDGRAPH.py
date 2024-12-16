import pandas as pd
import matplotlib.pyplot as plt


# Ruta del archivo con datos exportados
csv_path = input("Por favor, ingresa la ruta del archivo con datos exportados (CSV o TXT): ")
output_pdf_path = input("Por favor, ingresa la ruta completa donde deseas guardar el gráfico en PDF (incluye el nombre y .pdf): ")
try:
    # Cargar datos desde el archivo CSV
    data = pd.read_csv(csv_path)
    print("Datos cargados correctamente:")
    print(data.head())

    # Verificar si las columnas correctas existen
    if 'log(log(P/Po))' in data.columns and 'log(Vads)' in data.columns:
        # Extraer columnas
        x = data['log(log(P/Po))']
        y = data['log(Vads)']

        # Crear la gráfica
        plt.figure(figsize=(10, 6))
        plt.plot(x, y, marker='o', linestyle='-', color='blue')
        plt.title("Método FFH Desorption: log(log(P/Po)) vs log(Vads)", fontsize=14)
        plt.xlabel("log(log(P/Po))", fontsize=12)
        plt.ylabel("log(Vads)", fontsize=12)
        plt.grid(True, linestyle='--', alpha=0.6)

        # Guardar el gráfico como PDF
        plt.savefig(output_pdf_path, format='pdf')
        print(f"Gráfico guardado exitosamente en: {output_pdf_path}")

        # Mostrar el gráfico
        plt.show()

    else:
        print("Error: Las columnas 'log(log(P/Po))' y 'log(Vads)' no están en el archivo.")
except FileNotFoundError:
    print("Error: No se encontró el archivo.")
except Exception as e:
    print(f"Ocurrió un error: {e}")

