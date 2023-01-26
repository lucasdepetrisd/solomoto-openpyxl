def mensaje(prod, y, msg):
    if prod[0] is None:
        prod[0] = "Sin Codigo"

    if msg == "append":
        print("Appended: " + prod[0] + " a fila " + str(y))
    elif msg == "copy":
        print("Copied: " + prod[0] + " a fila " + str(y))

def progreso(progreso: int, total: int, mensaje = ""):
    porcentaje = 100 * (progreso / float(total))
    bar = '█' * int(porcentaje) + '-' * (100 - int(porcentaje))  
    print(f"\r|{bar}| {porcentaje:.2f}% | {mensaje}", end="\r")
def fin_progreso():
    print(f"\r|{'█'*100}| {100:.2f}%", end="\r")      