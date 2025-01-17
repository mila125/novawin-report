#Graficos a partir de los dataframes

import openpyxl
import os
import csv
import traceback
from datetime import datetime
from pywinauto import Application, findwindows
import time
import threading
from novawinmng import manejar_novawin, leer_csv_y_crear_dataframe,agregar_csv_a_plantilla_excel, guardar_dataframe_en_ini,generar_nombre_unico,agregar_dataframe_a_excel_sin_borrar,agregar_dataframe_a_nueva_hoja
from pywinauto.keyboard import send_keys
from openpyxl import Workbook
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from openpyxl import load_workbook
from pandas import ExcelWriter
from openpyxl import load_workbook
from pandas import ExcelWriter


def agregar_dataframe_a_nueva_hoja(archivo_excel, dataframe, nombre_hoja):
    # Cargar el archivo Excel existente
    book = load_workbook(archivo_excel)
    
    # Usar el modo 'append' para evitar sobrescribir el archivo
    with pd.ExcelWriter(archivo_excel, engine='openpyxl', mode='a', if_sheet_exists='new') as writer:
        # Pasar el objeto book al escritor
        writer._book = book  # Nota: Usamos '_book' en lugar de 'book'
        
        # Escribir el DataFrame en la nueva hoja
        dataframe.to_excel(writer, sheet_name=nombre_hoja, index=False)
            
def BET_BI(df, ruta_excel, Rango_de_Absorpcion, Rango_de_Desorpcion):

    # Verificar si las columnas necesarias existen en el DataFrame
    if 'Volume @ STP' not in df.columns or '1 / [ W((P/Po) - 1) ]' not in df.columns:
        print("Las columnas necesarias no están en el DataFrame.")
        return
    try:
        # Convertir Rango_de_Desorpcion a entero y filtrar
        num_filas = int(Rango_de_Desorpcion)
        ultimo_rango_df = df.tail(num_filas)
        print("Últimos elementos filtrados según Rango de Desorción:")
        print(ultimo_rango_df)
    except ValueError:
        print("El valor de Rango_de_Desorpcion no es válido para realizar el filtrado.")
    
    # Filtrar los rangos y mostrar los datos
    filtered_df_a = ultimo_rango_df[
        (ultimo_rango_df['Relative Pressure'] >= 0.45) & (ultimo_rango_df['Relative Pressure'] <= 0.55)
    ]
    print("Filtrado en rango 0.45-0.55:")
    print(filtered_df_a)

    # Calcular el promedio de la columna 'Volume @ STP'
    if not ultimo_rango_df['Volume @ STP'].isnull().all():  # Verificar si la columna no está vacía
        promedio_volume_stp_a = filtered_df_a['Volume @ STP'].mean()
        print(f"Promedio de 'Volume @ STP': {promedio_volume_stp_a}")
    else:
        print("La columna 'Volume @ STP' está vacía o contiene solo valores nulos.")

    filtered_df_b = ultimo_rango_df[
        (ultimo_rango_df['Relative Pressure'] >= 0.6) & (ultimo_rango_df['Relative Pressure'] <= 0.75)
    ]
    print("Filtrado en rango 0.6-0.75:")
    print(filtered_df_b)

    # Calcular el promedio de la columna 'Volume @ STP'
    if not filtered_df_b['Volume @ STP'].isnull().all():  # Verificar si la columna no está vacía
        promedio_volume_stp_b = filtered_df_b['Volume @ STP'].mean()
        print(f"Promedio de 'Volume @ STP': {promedio_volume_stp_b}")
    else:
        print("La columna 'Volume @ STP' está vacía o contiene solo valores nulos.")

    # Dividir los promedios
    div = promedio_volume_stp_a / promedio_volume_stp_b
    print("La división es: " + str(round(div, 2)))  # Convierte el número a cadena y lo redondea
    if (div >= 1):
        return True
    else:
        return False
    # Filtrar los últimos N elementos según el valor de Rango_de_Desorpcion
   
        
def BET_P(df, ruta_excel, Rango_de_Absorpcion, Rango_de_Desorpcion):

    # Verificar si las columnas necesarias existen en el DataFrame
    if 'Volume @ STP' not in df.columns or '1 / [ W((P/Po) - 1) ]' not in df.columns:
        print("Las columnas necesarias no están en el DataFrame.")
        return
    try:
        # Convertir Rango_de_Desorpcion a entero y filtrar
        num_filas = int(Rango_de_Desorpcion)
        ultimo_rango_df = df.tail(num_filas)
        print("Últimos elementos filtrados según Rango de Desorción:")
        print(ultimo_rango_df)
    except ValueError:
        print("El valor de Rango_de_Desorpcion no es válido para realizar el filtrado.")

    # Verificar si las columnas necesarias existen en el DataFrame
    if 'Volume @ STP' not in df.columns or '1 / [ W((P/Po) - 1) ]' not in df.columns:
        print("Las columnas necesarias no están en el DataFrame.")
        return
    
     # Filtrar los rangos y mostrar los datos
    filtered_df_a = ultimo_rango_df[
        (ultimo_rango_df['Relative Pressure'] >= 0.6) & (ultimo_rango_df['Relative Pressure'] <= 0.75)
    ]
    print("Filtrado en rango 0.6-0.75:")
    print(filtered_df_a)

    # Calcular el promedio de la columna 'Volume @ STP'
    if not ultimo_rango_df['Volume @ STP'].isnull().all():  # Verificar si la columna no está vacía
        promedio_volume_stp_a = filtered_df_a['Volume @ STP'].mean()
        print(f"Promedio de 'Volume @ STP': {promedio_volume_stp_a}")
    else:
        print("La columna 'Volume @ STP' está vacía o contiene solo valores nulos.")
    
    try:
        # Convertir Rango_de_Absorpcion a entero y filtrar
        num_filas = int(Rango_de_Absorpcion)
        
        # Obtener los primeros "num_filas" elementos
        primeros_rango_df = df.head(num_filas)
        print("Primeros elementos filtrados según Rango de Rango_de_Absorpcion:")
        print(primeros_rango_df)
    except ValueError:
        print("El valor de Rango_de_Absorpcion no es válido para realizar el filtrado.")
        return
   
    filtered_df_b = ultimo_rango_df[
        (ultimo_rango_df['Relative Pressure'] >= 0.6) & (ultimo_rango_df['Relative Pressure'] <= 0.75)
    ]
    print("Filtrado en rango 0.6-0.75:")
    print(filtered_df_b)

    # Calcular el promedio de la columna 'Volume @ STP'
    if not filtered_df_b['Volume @ STP'].isnull().all():  # Verificar si la columna no está vacía
        promedio_volume_stp_b = filtered_df_b['Volume @ STP'].mean()
        print(f"Promedio de 'Volume @ STP': {promedio_volume_stp_b}")
    else:
        print("La columna 'Volume @ STP' está vacía o contiene solo valores nulos.")

    # Dividir los promedios
    div = promedio_volume_stp_a / promedio_volume_stp_b
    print("La división es: " + str(round(div, 2)))  # Convierte el número a cadena y lo redondea
    if (div >= 1):
        return True
    else:
        return False
    # Filtrar los últimos N elementos según el valor de Rango_de_Desorpcion
def BET_C(df, ruta_excel, Rango_de_Absorpcion, Rango_de_Desorpcion):


    # Verificar si las columnas necesarias existen en el DataFrame
    if 'Volume @ STP' not in df.columns or '1 / [ W((P/Po) - 1) ]' not in df.columns:
        print("Las columnas necesarias no están en el DataFrame.")
        return
    try:
        # Convertir Rango_de_Desorpcion a entero y filtrar
        num_filas = int(Rango_de_Desorpcion)
        ultimo_rango_df = df.tail(num_filas)
        print("Últimos elementos filtrados según Rango de Desorción:")
        print(ultimo_rango_df)
    except ValueError:
        print("El valor de Rango_de_Desorpcion no es válido para realizar el filtrado.")

    # Verificar si las columnas necesarias existen en el DataFrame
    if 'Volume @ STP' not in df.columns or '1 / [ W((P/Po) - 1) ]' not in df.columns:
        print("Las columnas necesarias no están en el DataFrame.")
        return
    
     # Filtrar los rangos y mostrar los datos
    filtered_df_a = ultimo_rango_df[
        (ultimo_rango_df['Relative Pressure'] >= 0.95) & (ultimo_rango_df['Relative Pressure'] <= 1.0)
    ]
    print("Filtrado en rango 0.95-1.0:")
    print(filtered_df_a)
   
    try:
        # Convertir Rango_de_Absorpcion a entero y filtrar
        num_filas = int(Rango_de_Absorpcion)
        
        # Obtener los primeros "num_filas" elementos
        primeros_rango_df = df.head(num_filas)
        print("Primeros elementos filtrados según Rango de Rango_de_Absorpcion:")
        print(primeros_rango_df)
    except ValueError:
        print("El valor de Rango_de_Absorpcion no es válido para realizar el filtrado.")
        return
   
    filtered_df_b = ultimo_rango_df[
        (ultimo_rango_df['Relative Pressure'] >= 0.95) & (ultimo_rango_df['Relative Pressure'] <= 1.0)
    ]
    print("Filtrado en rango 0.95-1.0:")
    print(filtered_df_b)

    # Calcular el promedio de la columna 'Volume @ STP'
    if not filtered_df_b['Volume @ STP'].isnull().all():  # Verificar si la columna no está vacía
        promedio_volume_stp_ab = filtered_df_b['Volume @ STP'].mean()
        promedio_volume_stp_ab = promedio_volume_stp_ab + filtered_df_a['Volume @ STP'].mean()
        print(f"Promedio de 'Volume @ STP de valores de Absorbcion y Desorbcion es ': {promedio_volume_stp_ab}")
    else:
        print("La columna 'Volume @ STP' está vacía o contiene solo valores nulos.")

    if (promedio_volume_stp_ab >= 1):
        return True
    else:
        return False

def rangos_dft_main(archivo_ruta_completa, archivo_planilla):
    """
    Procesa un archivo para clasificar los datos en 5 dataframes según rangos definidos
    en la columna 'Half pore width', calcula los valores de 'Cumulative Pore Volume'
    de la última fila de cada rango, y escribe estos valores con sus etiquetas en la hoja 'DFT'.
    
    Args:
    archivo_ruta_completa (str): Ruta completa del archivo que contiene los datos.
    archivo_planilla (str): Tipo de archivo ('csv' o 'excel').
    
    Returns:
    dict: Un diccionario con los 5 dataframes correspondientes a cada rango y los valores
          'Cumulative Pore Volume' para la última fila de cada dataframe.
    """
    print("Inicio de tests_main")
    print(f"Archivo ruta completa: {archivo_ruta_completa}")
    
    # Carga del archivo
    if archivo_planilla.lower() == 'csv':
        raise ValueError("Para escribir en la hoja de cálculo, utiliza un archivo Excel.")
    elif archivo_planilla.lower() in ['excel', 'xlsx']:
        df = pd.read_excel(archivo_ruta_completa, sheet_name="DFT")
    else:
        raise ValueError("Formato de archivo no soportado. Usa 'excel'.")
    
    print("Archivo cargado exitosamente.")
    
    # Filtrar datos en diferentes rangos
    rangos = {
        "MICRO": (0, 10.5),
        "UNCLASS": (10.49, 18.9),
        "BOTLE": (18.89, 24.8),
        "PLATE": (24.79, 109.7),
        "CYL": (109.69, float('inf'))
    }
    
    dataframes = {}
    cumulative_pore_volumes = {}
    
    for nombre, (min_val, max_val) in rangos.items():
        # Filtrar datos dentro del rango
        rango_df = df[(df['Half pore width'] > min_val) & (df['Half pore width'] <= max_val)]
        dataframes[nombre] = rango_df
        
        # Obtener el valor de 'Cumulative Pore Volume' de la última fila
        if not rango_df.empty:
            cumulative_pore_volumes[f"RANGO_CV_{nombre}"] = rango_df['Cumulative Pore Volume'].iloc[-1]
        else:
            cumulative_pore_volumes[f"RANGO_CV_{nombre}"] = None
        
        print(f"Datos filtrados para {nombre}: {len(rango_df)} filas.")
    
    # Escribir los resultados en la hoja 'DFT'
    with pd.ExcelWriter(archivo_ruta_completa, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        # Crear un dataframe con los valores y etiquetas
        resultados = pd.DataFrame({
            "Etiqueta": ["MICRO", "UNCLASS", "BOTLE", "PLATE", "CYL"],
            "Valor": [
                cumulative_pore_volumes["RANGO_CV_MICRO"],
                cumulative_pore_volumes["RANGO_CV_UNCLASS"],
                cumulative_pore_volumes["RANGO_CV_BOTLE"],
                cumulative_pore_volumes["RANGO_CV_PLATE"],
                cumulative_pore_volumes["RANGO_CV_CYL"]
            ]
        })
        
        # Escribir en la hoja 'DFT', comenzando en una celda específica (por ejemplo, B2)
        resultados.to_excel(writer, sheet_name="DFT", index=False, startrow=1, startcol=1, header=True)
    
    print("Valores escritos en la hoja 'DFT'.")
    return dataframes, cumulative_pore_volumes