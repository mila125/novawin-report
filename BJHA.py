import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


def bjha_main(file_path,output_pdf_path):
 print("Hola desde BJHA:main")
 try:
     # Leer el archivo CSV
     data = pd.read_csv(file_path)
    
     # Verificar si las columnas necesarias están presentes
     if 'Relative Pressure' in data.columns and 'Volume @ STP' in data.columns:
         # Cálculo de dV(r) (distribución del volumen)
        
         # Paso 1: Calcular la presión relativa P/P0 y volumen adsorbido V(P/P0)
         P = data['Relative Pressure']
         V = data['Volume @ STP']
        
         # Paso 2: Calcular el cambio en el volumen (dV)
         dV = np.gradient(V)  # Calculando la derivada del volumen respecto a la presión
        
         # Paso 3: Calcular el tamaño del poro (r) utilizando la fórmula de Kelvin (aproximación)
         sigma = 0.02  # Constante de adsorción (este valor depende de las condiciones experimentales)
         P0 = 1  # Suponiendo presión de saturación P0 = 1 atm para simplificación
         r = (-4 * sigma * np.log(P / P0)) / (np.gradient(P))  # Aproximación de los radios de poro
        
         # Paso 4: Graficar dV(r)
         plt.figure(figsize=(10, 6))
         plt.plot(r, dV, marker='o', label="dV(r) - Absorción", color='blue')
        
         # Configurar el gráfico
         plt.title('Distribución de Tamaño de Poros - Método BJH (dV(r))', fontsize=16)
         plt.xlabel('Radio de Poro (r)', fontsize=14)
         plt.ylabel('dV(r)', fontsize=14)
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
