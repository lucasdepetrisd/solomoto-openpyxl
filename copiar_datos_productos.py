# Copiar filas de un archivo a otro

import math
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

p = load_workbook('../productos (1).xlsx')
pt = load_workbook('../null_5780278647463692669.xlsx')
wp = p.active
wpt = pt.active

p_row_count = wp.max_row
pt_row_count = wpt.max_row
print("Nro filas en productos (1).xlsx: " + str(p_row_count))
# print("Nro filas en null_5780278647463692669.xlsx: " + str(pt_row_count))
codp_list = []
codpt_list = []

def mensaje(prod, y, msg):
    if prod[0] is None:
        prod[0] = "Sin Codigo"

    if msg == "append":
        print("Appended: " + prod[0] + " a fila " + str(y))
    elif msg == "copy":
        print("Copied: " + prod[0] + " a fila " + str(y))

def progreso(progreso, total):
    porcentaje = 100 * (progreso / float(total))
    bar = '█' * int(porcentaje) + '-' * (100 - int(porcentaje))
    print(f"\r|{bar}| {porcentaje:.2f}%", end="\r")
    # print("Progreso: " + "%.2f" % (x*100/total) + "%")

def copy_row(x, y):
    tuplapt = tuple(wpt['A' + str(y+2): 'AE' + str(y+2)])
    for z in range(31):
        valor = tuplapt[0][z].internal_value
        wp.cell(x+2, 18+z).value = valor

def copy_product(prod, pt_row):
    # print(prod)
    wpt.cell(pt_row, 1).value = prod[0]
    wpt.cell(pt_row, 2).value = prod[1]
    wpt.cell(pt_row, 3).value = prod[2]
    wpt.cell(pt_row, 9).value = prod[3]
    wpt.cell(pt_row, 13).value = prod[4]


for x in range(2, p_row_count+1, 1):
    codp = wp.cell(x, 8).value
    codbarrap = wp.cell(x, 9).value
    nombp = wp.cell(x, 10).value
    provp = wp.cell(x, 11).value
    preciop = wp.cell(x, 12).value
    codp_list.append([codp, codbarrap, nombp, provp, preciop])
    # print(str(x))
    # print(str(codp))

for x in range(2, pt_row_count+1, 1):
    codpt = wpt.cell(x, 1).value
    codpt_list.append(codpt)

lencodp_list = len(codp_list)
lencodpt_list = len(codpt_list)

fp = open(r'../codigosp.txt', 'w')

for x, prod in enumerate(codp_list):
    for y, codpt in enumerate(codpt_list):
        if prod[0] == codpt:
            copy_product(prod, y + 2)
            # mensaje(prod, y, "add")
            fp.write("(%s), " % prod[0] + ", " + str(y+1))
            break
        elif y == lencodpt_list-1:
            copy_product(prod, lencodpt_list)
            # mensaje(prod, y, "append")
            codpt_list.append(prod[0])
            lencodpt_list += 1
            fp.write("(%s), " % prod[0] + ", " + str(y+1))
            break
    progreso(x, lencodp_list)

pt.save('../null_5780278647463692669.xlsx')