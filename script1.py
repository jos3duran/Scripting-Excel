import os
import glob
import xlrd
from tkinter import filedialog

def leer_archivo(ruta):
    """
    Lee un archivo .xls y devuelve un diccionario con las peliculas y el top 5 por cada archivo

    Args:
        ruta (str): Ruta al archivo .xls.

    Returns:
        dict: Diccionario con las siguientes claves:
            * "peliculas": Diccionario con las peliculas como claves y la asistencia total como valores.
            * "top_5": Diccionario con los nombres de los archivos como claves y una lista con el top 5 de cada archivo con su asistencia como valores.
    """
    if not os.path.exists(ruta):
        print(f"Error: El archivo '{ruta}' no existe.")
        return None

    wb=xlrd.open_workbook(ruta)
    ws=wb.sheet_by_index(0)

    filas=ws.nrows

    peliculas={}

    for i in range(filas):
        value = ws.cell_value(i, 1)
        ast=ws.cell_value(i,11)

        if value and value != "Title":
            try:
                asistencia = int(ast)
            except ValueError:
                if isinstance(ast, str):
                    continue
                else:
                    print(f"Error: valor inesperado en la fila {i+1} para la película: {value}")
                    continue

            peliculas[value]=peliculas.get(value,0)+asistencia

    top_5 = sorted(peliculas.items(), key=lambda x: x[1], reverse=True)[:5]
    
    return {"peliculas": peliculas, "top_5": top_5}

def leer_directorio():
    """
    Lee todos los archivos .xls en un directorio y devuelve un diccionario con el top 5 regional de las películas

    Args:
        ruta_directorio (str): Ruta al directorio que contiene los archivos .xls.

    Returns:
        None
    """
    
    #Usuario seleccionara folder con los archivos
    ruta_directorio  = filedialog.askdirectory(title="Seleccione la carpeta de archivos .xls")

    # Si el usuario cancela la selección o no selecciona un folder, no se hace nada
    if not ruta_directorio :
        return 

    # Validar que la carpeta contenga al menos un archivo .xls
    archivos_excel = [archivo for archivo in os.listdir(ruta_directorio) if archivo.endswith(".xls")]

    if not archivos_excel:
        print("Error: No se encontraron archivos .xls en la carpeta seleccionada.")
        return

    peliculas_globales = {}  # Diccionario global para acumular asistencia
    top_5_archivos = {}  # Diccionario para almacenar el top 5 de cada archivo

    for archivo in archivos_excel:
        ruta_archivo = os.path.join(ruta_directorio, archivo)
        resultado_archivo = leer_archivo(ruta_archivo)  # Llama a leer_archivo

        # Acumula asistencia en peliculas_globales
        for titulo, asistencia in resultado_archivo["peliculas"].items():
            peliculas_globales[titulo] = peliculas_globales.get(titulo, 0) + asistencia

        # Almacena el top 5 individual
        top_5_archivos[archivo] = resultado_archivo["top_5"]

            
    # Ordena peliculas_globales por asistencia
    peliculas_globales = sorted(peliculas_globales.items(), key=lambda x: x[1], reverse=True)
    
    # Escribe el resultado en un archivo
    escribir_archivo("top_5_peliculas.txt", peliculas_globales[:5], top_5_archivos)

    return None

def escribir_archivo(nombre_documento, top_5_global, top_5_archivos):
    """
    Escribe en un archivo .txt el resultado del analisis

    Args:
        nombre_documento (str): nombre del documento de texto plano a crear.
        top_5_global (dict): Diccionario con el top 5 regional de peliculas como claves y la asistencia como valor
        top_5_archivos (dict): Diccionario con los nombres de los archivos como claves y una lista con el top 5 de cada archivo con su asistencia como valores. 

    Returns:
        None
    """
    with open(nombre_documento, "w") as archivo:
        # Write global top 5 movies
        archivo.write("Top 5 Peliculas Nivel Regional\n")
        archivo.write("{:<30}{:<20}{:>10}\n".format('Posición', 'Titulo', 'Asistencia'))
        archivo.write('-' * 67 + '\n')

        for i, (titulo, asistencia) in enumerate(top_5_global, 1):
            archivo.write(f"{i:2} {titulo:>40}{asistencia:>15}\n")

        # Write top 5 movies per file
        archivo.write("\n\n")
        archivo.write("Top 5 por pais:\n")
        for nombre_archivo, top_5 in top_5_archivos.items():
            archivo.write(f"\n**{nombre_archivo}**\n")
            archivo.write("{:<30}{:<20}{:>10}\n".format('Posición', 'Titulo', 'Asistencia'))
            archivo.write('-' * 67 + '\n')

            for i, (titulo, asistencia) in enumerate(top_5, 1):
                archivo.write(f"{i:2} {titulo:>40}{asistencia:>15}\n")
                
    return None

if __name__ == "__main__":
    leer_directorio()