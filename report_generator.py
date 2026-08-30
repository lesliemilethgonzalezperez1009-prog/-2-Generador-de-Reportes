import csv

with open('ventas.csv', 'r') as cvs_file:
    cvs_reader = csv.reader(cvs_file)
    next(cvs_reader)

    for line in cvs_reader:
        print(f"Producto: {line[0]} | Categoria: {line[1]} | Unidades: {line[2]} | Precio unitario: {line[3]} ")

  #  with open ('DATA', mode = 'w', newline= "",encoding = 'UTF-8') as file:
   #     csv_writer = csv.writer(file ,delimiter= "|")
#
 ##          csv_writer.writerow(line)
 # print ('TODO LISTO')
        