# Lista a txt

fp = open(r'./codigosp.txt', 'w')
for codigos in codp_list:
	fp.write("%s, " % codigos)
print("Archivo P listo.")

fp = open(r'./codigospt.txt', 'w')
for codigos in codpt_list:
	fp.write("%s, " % codigos)
print("Archivo PT listo.")