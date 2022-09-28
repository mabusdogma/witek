#! /usr/bin/env python3

import os
import pandas as pd
import openpyxl
import time
import warnings

#desactiva advertencias
warnings.simplefilter("ignore")

#archivos de origen y destino, destino lleva una v detras del nombre
print ("\nEste script copia un archivo de Excel a solo valores, para procesar rápidamente")
print("Por favor, arrastre hasta aqui el archivo o escriba la ruta completa")
print('Ejemplo:', r'C:\Users...')
origen = input('\n\n')

#si se arrastra archivo desde Windows, quitar las comillas al inicio y al final
origen = origen.replace('"', '')

#si se arrastra archivo desde WSL, truncar la parte inicial de la dirección y cambiar las barras
if origen.find(r'wsl'):
    origen = origen.replace(r'\\wsl.localhost\Ubuntu', '')
    origen = origen.replace('\\', '/')

#concatena ruta y muestra nombre del archivo destino
destino = str(os.path.splitext(origen)[0]) + 'v' + str(os.path.splitext(origen)[1])       
print('')

#tiempo desde donde se cuenta la conversion
startTime = time.time()

#abre archivo origen y asigna variable las hojas
xl = pd.read_excel(origen, header=None, index_col=None, sheet_name=None)
sheets = xl.keys()

#calcula cuantas hojas lleva y cuantas en total
res = pd.ExcelFile(origen)
total = len(res.sheet_names)

#crea archivo destino y cambiando el nombre de la primera hoja por defecto (Sheet)
wb = openpyxl.Workbook()
ws = wb.active
ws.title = res.sheet_names[0]
wb.save(destino)   

#rellenar hoja por hoja en destino, solo valores
actual=1

with pd.ExcelWriter(destino, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
    for sheet in sheets:
        #crea barra de progreso
        print("\r[", actual, "/", total, "]", end='\r')  
        #crea o reemplaza la hoja en destino
        xl[sheet].to_excel(writer, sheet_name=sheet, index=False, header=False)
        actual += 1
print('Archivo copiado correctamente!\n')
executionTime = (time.time() - startTime)
print(f'Tiempo de ejecución: {executionTime:.2f} segundos')