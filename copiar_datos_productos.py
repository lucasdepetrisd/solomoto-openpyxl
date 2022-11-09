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
            fp.write("(%s), " % prod[0] + ", " + str(y+1))
            break
        elif y == lencodpt_list-1:
            copy_product(prod, lencodpt_list)
            codpt_list.append(prod[0])
            lencodpt_list += 1
            fp.write("(%s), " % prod[0] + ", " + str(y+1))
            break

# pt.save('../null_5780278647463692669.xlsx')

# 	# print(str(nom) + "\t" + str(cat) + "\t" + str(cuil) + "\t" + str(dom) + "\t" + str(cel))
# 	for y in range(filastot, 1, -1):
# 		# print(str(y))
# 		nom2 = wp.cell(y, 2).value
# 		cat2 = wp.cell(y, 7).value
# 		cuil2 = wp.cell(y, 8).value
# 		dom2 = wp.cell(y, 16).value
# 		cel2 = wp.cell(y, 20).value
# 		if (x != y):
# 			if (nom1 == nom2 and cat1 == cat2 and cuil1 == cuil2 and dom1 == dom2 and cel1 == cel2):
# 				print("Se encontro duplicado en filas " + str(x) + " y " + str(y))
# 				eliminar(y)
# 				encontrados += 1
# print("Se encontraron " + str(encontrados) + " duplicados")
