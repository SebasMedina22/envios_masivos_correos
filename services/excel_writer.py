import openpyxl
from datetime import datetime

def marcar_como_enviado(ruta_archivo, correo_destino, nombre_hoja):
    workbook = openpyxl.load_workbook(ruta_archivo)
    hoja = workbook[nombre_hoja]

    for fila in hoja.iter_rows(min_row=2):  
        correo = fila[1].value

        if correo == correo_destino:
            fila[3].value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Actualiza SIEMPRE
            break

    workbook.save(ruta_archivo)

