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
previo = input('\n\n')
startTime = time.time()

#si se arrastra archivo, quitar las comillas al inicio y al final
origen = previo.replace('"', '')
 
#concatena ruta y muestra nombre del archivo destino
destino = str(os.path.splitext(origen)[0]) + 'v' + str(os.path.splitext(origen)[1])       
print ("\nArchivo destino: ")
print (destino)
print('')

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
print('Espere...')
with pd.ExcelWriter(destino, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
    with concurrent.futures.ThreadPoolExecutor() as executor:
        #crea o reemplaza la hoja en destino
        futures = [executor.submit(xl[sheet].to_excel(writer, sheet_name=sheet, index=False, header=False)) for sheet in sheets]      

#finalizar     
print('Archivo copiado correctamente!\n')
executionTime = (time.time() - startTime)
print(f'Tiempo de ejecución: {executionTime:.2f} segundos')