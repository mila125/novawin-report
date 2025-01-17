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

def draw_comparison_bar_chart(bjhd, bjha):
    # Filtrar las columnas necesarias para ambos dataframes
    radius_bjhd = bjhd['Radius']
    dV_logr_bjhd = bjhd['dV(logr)']
    
    radius_bjha = bjha['Radius']
    dV_logr_bjha = bjha['dV(logr)']
    
    # Asegurarse de que ambas series tengan el mismo tamaño
    min_length = min(len(radius_bjhd), len(radius_bjha))  # Tomar la longitud mínima
    radius_bjhd = radius_bjhd[:min_length]
    dV_logr_bjhd = dV_logr_bjhd[:min_length]
    radius_bjha = radius_bjha[:min_length]
    dV_logr_bjha = dV_logr_bjha[:min_length]

    # Crear el gráfico de barras
    plt.figure(figsize=(12, 6))
    
    # Dibujar las barras para bjhd y bjha
    bar_width = 0.5  # Aumentar el ancho de las barras
    index = range(len(radius_bjhd))
    
    plt.bar(
        index, dV_logr_bjhd, bar_width, 
        label='Desorbcion', edgecolor='blue', fill=False, linewidth=1.5
    )
    plt.bar(
        [i + bar_width for i in index], dV_logr_bjha, bar_width, 
        label='Absorpcion', edgecolor='green', fill=False, linewidth=1.5
    )

    # Etiquetas y título
    plt.xlabel('Radius (Å)')
    plt.ylabel('dV(logr) (cc/Å/g)')
    plt.title('BJH Absorpcion y Desorbcion ')
    
    # Configuración de la cuadrícula
    plt.xticks([i + bar_width / 2 for i in index], radius_bjhd, rotation=90)
    plt.legend()
    plt.grid(axis='y', linestyle='--', alpha=0.6)
    
    # Mostrar el gráfico
    plt.tight_layout()
    plt.show()
    
def draw_DFT(df):
    # Filtrar las columnas necesarias
    half_pore_width = df['Half pore width']
    dVr = df['dV(r)']
    
    # Crear el gráfico de barras
    plt.figure(figsize=(12, 6))
    bar_width = 0.8  # Aumentar el ancho de las barras
    
    plt.bar(half_pore_width, dVr, width=bar_width, color='green', alpha=0.7, label='dV(r)')
    
    # Etiquetas y título
    plt.xlabel('Half pore width (Å)')
    plt.ylabel('dV(r) (cc/Å/g)')
    plt.title('DFT Analysis: Half Pore Width vs dV(r)')
    plt.legend()
    
    # Configuración de la cuadrícula
    plt.grid(axis='y', linestyle='--', alpha=0.6)
    
    # Mostrar el gráfico
    plt.tight_layout()
    plt.show()

def draw_HK(df):
    # Filtrar las columnas necesarias
    columna_dV = df['dV()']
    columna_half_pore_width = df['Half pore width']
    
    # Crear el gráfico de barras
    plt.figure(figsize=(14, 6))
    x_positions = range(len(columna_dV))  # Posiciones en el eje x
    plt.bar(x_positions, columna_dV, color='green', alpha=0.7, label='dV(r)')

    # Añadir etiquetas de texto en las barras
    #for i, width in enumerate(columna_half_pore_width):
    #    plt.text(i, columna_dV[i] + 0.05, f'{width:.2f}', ha='center', fontsize=8, color='blue')
    
    # Etiquetas y título
    plt.xlabel('Half pore width , A')
    plt.ylabel('dV cc A g')
    plt.title(' gráfico HK ')
    plt.xticks(x_positions, x_positions, rotation=90)  # Mostrar las filas como índices
    plt.legend()

    # Mostrar el gráfico
    plt.tight_layout()
    plt.show()

def graphs_main(archivo_planilla):
    print("Inicio de graphs_main")
    print(archivo_planilla)
       
    # Crear un diccionario para almacenar los valores extraídos
    valores_nkaffha = {}
    
    # Leer los datos de la hoja 'NKA'
    df_nka = pd.read_excel(archivo_planilla, sheet_name='NKA')
    valores_nka = df_nka.tail(2)  # Obtener los dos últimos valores
    valores_nkaffha['NKA'] = valores_nka
    
    # Leer los datos de la hoja 'FFHA'
    df_ffha = pd.read_excel(archivo_planilla, sheet_name='FFHA')
    valores_ffha = df_ffha.tail(2)  # Obtener los dos últimos valores
    valores_nkaffha['FFHA'] = valores_ffha
    
    # Crear un nuevo DataFrame combinando los valores extraídos
    nuevo_df = pd.concat([
        valores_nka.reset_index(drop=True),
        valores_ffha.reset_index(drop=True)
    ], keys=['NKA', 'FFHA'], names=['Sheet', 'Index'])
    
    # Guardar en una nueva hoja llamada 'NKAFFHA_valores'
    with pd.ExcelWriter(archivo_planilla, mode='a', engine='openpyxl') as writer:
        nuevo_df.to_excel(writer, sheet_name='NKAFFHA_valores', index=False)
    print("Se ha creado la hoja 'NKAFFHA_valores' con los datos de las hojas 'NKA' y 'FFHA'.")
    
    # Leer el archivo Excel
    df = pd.read_excel(archivo_planilla, sheet_name='HK')
    draw_HK(df)
    
    # Verificar las primeras filas del archivo para asegurarse de que las columnas existen
    print(df.head())
        # Leer el archivo Excel
    df = pd.read_excel(archivo_planilla, sheet_name='DFT')
    draw_DFT(df)
    
    # Verificar las primeras filas del archivo para asegurarse de que las columnas existen
    print(df.head())

    # Leer el archivo Excel
    df_bjha = pd.read_excel(archivo_planilla, sheet_name='BJHA')
    df_bjhd = pd.read_excel(archivo_planilla, sheet_name='BJHD')
   
    draw_comparison_bar_chart(df_bjhd, df_bjha)
    # Verificar las primeras filas del archivo para asegurarse de que las columnas existen
    print(df.head())

