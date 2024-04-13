from tkinter import filedialog
import pandas as pd
import os

ruta_directorio = filedialog.askdirectory(title="Seleccione la carpeta de archivos .xls")

excel_files = [os.path.join(ruta_directorio, archivo) for archivo in os.listdir(ruta_directorio) if archivo.endswith(".xls")]

if not excel_files:
    print("No se encontraron archivos .xls en el directorio seleccionado.")
    exit()

books = [pd.read_excel(excel, skiprows=2, usecols=['Title', 'Circuit', 'Week\nAdm']) for excel in excel_files]

book = pd.concat(books, axis=0)

attendance_per_cinema = book.groupby(['Title', 'Circuit'])['Week\nAdm'].sum().reset_index()

total_per_movie = attendance_per_cinema.groupby('Title')['Week\nAdm'].sum()

attendance_per_cinema['%'] = attendance_per_cinema.apply(lambda row: round(row['Week\nAdm'] / total_per_movie[row['Title']] * 100, 2), axis=1)

# Exportar los resultados a un archivo de texto con el formato de espacios debidamente ajustado
with open('asistencia_por_cine.txt', 'w') as file:
    for title, group in attendance_per_cinema.groupby('Title'):
        file.write(title + '\n')
        file.write("{:<30}{:<20}{}\n".format('circuit', 'asistentes', '%'))
        file.write("\n")
        for _, row in group.iterrows():
            file.write("{:<30}{:<20}{:.0f}\n".format(row['Circuit'], row['Week\nAdm'], row['%']))
        file.write("\n") 
        file.write("{:<30}{:<20}{:.2f}\n\n".format('total', group['Week\nAdm'].sum(), 100.00))
        file.write("-"*70 + "\n")
