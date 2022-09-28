# witek
Creates a COPY of Excel files to values-only.

Crea una COPIA de archivos de Excel a solo valores (menor tamaño de archivo, se trabaja más facilmente).
La copia se generará en la misma carpeta, con el mismo nombre y una 'v' antes de la extension (se puede buscar con *v.xls o *v.xlsx dependiendo del formato).

Funciona en Windows, Linux y Mac, no necesita Microsoft Excel para funcionar.

Versión normal (witek) y de hilos multiples (witek_mt), su velocidad dependerá del procesador. Se pueden comparar ambas versiones (muestran el tiempo total de conversión).

Requerimientos:

- Python 3.0 o mayor.
- Módulos time, os, pandas, warnings, openpyxl (y current.futures, en el caso de la versión con hilos múltiples).
