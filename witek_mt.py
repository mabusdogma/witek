#!/usr/bin/env python3

import os
import pandas as pd
import openpyxl
import concurrent.futures
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
if os.name == 'nt':
    origen = origen.replace('"', '')
    import win32api,win32process,win32con
    win32process.SetPriorityClass(win32api.GetCurrentProcess(), win32process.HIGH_PRIORITY_CLASS) 
else:
    origen = origen.replace(r'\\wsl.localhost\Ubuntu', '')
    origen = origen.replace('\\', '/')
    os.nice(-18)
    
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

#crea archivo destino y primera hoja
sheet = res.sheet_names[0]
xl[sheet].to_excel(destino, engine="xlsxwriter", sheet_name=sheet, index=False, header=False)
itersheets = iter(sheets)
next(itersheets)  

#rellenar hoja por hoja en destino, solo valores
with pd.ExcelWriter(destino, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
    with concurrent.futures.ThreadPoolExecutor() as executor:
        #crea o reemplaza la hoja en destino
        futures = [executor.submit(xl[sheet].to_excel(writer, sheet_name=sheet, index=False, header=False)) 
                   for sheet in itersheets]    

#finalizar     
print('Archivo copiado correctamente!\n')
executionTime = (time.time() - startTime)
print(f'Tiempo de ejecución: {executionTime:.2f} segundos')