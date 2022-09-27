import os
import pandas as pd
import warnings

warnings.simplefilter("ignore")
#archivos de origen y destino, destino lleva una v detras del nombre
print ("\nEste script copia un archivo de Excel a solo valores, para procesar rápidamente")
print("Por favor, arrastre hasta aqui el archivo o escriba la ruta completa")
print('Ejemplo:', r'C:\Users...')
previo = input('\n\n')

#si se arrastra archivo, quitar las comillas al inicio y al final
origen = previo.replace('"', '')


#revisa si el archivo existe, para crearlo o sustituirlo
if os.path.isfile(origen):
    pass
else:
    exit("\nNo se encuentra el archivo seleccionado...\n")
 
#concatena ruta y muestra nombre del archivo destino
destino = str(os.path.splitext(origen)[0]) + 'v' + str(os.path.splitext(origen)[1])       
print ("\nArchivo destino: ")
print (destino)
print('')

#abre archivo origen y asigna variable las hojas
xl = pd.read_excel(origen, header=None, index_col=None, sheet_name=None)
sheets = xl.keys()
res = pd.ExcelFile(origen)

writer = pd.ExcelWriter(destino, engine='xlsxwriter')
writer.save()
print('Espere...')

with pd.ExcelWriter(destino, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
    for sheet in sheets:
        xl[sheet].to_excel(writer, sheet_name=sheet, index=False, header=False)
print('Archivo copiado correctamente!\n')