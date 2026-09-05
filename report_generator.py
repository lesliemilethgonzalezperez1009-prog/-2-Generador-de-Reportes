import csv

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.styles import Alignment
total_unidades= 0
total_ventas = 0

book = Workbook()
sheet = book.active
fuente = Font(name='Arial', size=12, bold=True)

sheet.merge_cells('A1:D1')
sheet['A1'] = 'ANALISIS DE DATOS CSV'
sheet['A1'].font = fuente


with open('ventas.csv', 'r') as cvs_file:
    cvs_reader = csv.reader(cvs_file)
    next(cvs_reader)
    sheet.append(['Producto', 'Categoria', 'Unidades', 'Precio unitario'])

    for line in cvs_reader:
        print(f"Producto: {line[0]} | Categoria: {line[1]} | Unidades: {line[2]} | Precio unitario: {line[3]} ")
        sheet.append([line[0], line[1], line[2], line[3]]) 
        total_unidades+= int(line[2])
        ventas  = int(line[3]) * float(line[2])
        total_ventas = total_ventas + ventas
    sheet.append(['Total', '', total_unidades, total_ventas])
    for cell in sheet[2]:
        cell.font = fuente
        cell.alignment = Alignment(horizontal='left')
    Ultima_fila = sheet.max_row
    for cell in sheet[Ultima_fila]:
        cell.font = fuente
        cell.alignment = Alignment(horizontal='left')

    for columna in ['A', 'B', 'C', 'D']:
        sheet.column_dimensions[columna].width = 20 
        

book.save('Reporte.xlsx')