import openpyxl

def leer_datos_excel(ruta_archivo, nombre_hoja):
    workbook = openpyxl.load_workbook(ruta_archivo)
    hoja = workbook[nombre_hoja]
    
    datos = []
    for fila in hoja.iter_rows(min_row=2):  # Empieza desde la fila 2
        nombre = fila[0].value
        correo = fila[1].value
        fecha_expiracion = fila[2].value
        enviado = fila[3].value if len(fila) > 3 else None

        datos.append({
            "nombre": nombre,
            "correo": correo,
            "fecha_expiracion": fecha_expiracion,
            "enviado": enviado
        })
    
    return datos

