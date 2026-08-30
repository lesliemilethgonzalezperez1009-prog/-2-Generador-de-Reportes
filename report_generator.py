import csv

from openpyxl import Workbook
from openpyxl.styles import Font

book = Workbook()
sheet = book.active

sheet.merge_cells('A1:D1')
sheet['A1'] = 'ANALISIS DE DATOS CSV'


with open('ventas.csv', 'r') as cvs_file:
    cvs_reader = csv.reader(cvs_file)
    next(cvs_reader)

    for line in cvs_reader:
        print(f"Producto: {line[0]} | Categoria: {line[1]} | Unidades: {line[2]} | Precio unitario: {line[3]} ")
        sheet.append([line[0], line[1], line[2], line[3]]) 

book.save('Reporte.xlsx')