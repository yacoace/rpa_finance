import os
import datetime
import logging
import shutil
import pandas as pd
from openpyxl.worksheet.table import Table, TableStyleInfo


logging.basicConfig(filename="C:\\Users\\dayanna_vidaurre\\log_scrip_rim_pendiente_v1.log",
                    format='%(asctime)s %(message)s',
                    filemode='a')

logger = logging.getLogger()

# Función para encontrar ruta donde se encuentra el One Drive Personal
def buscar_directorio(target, ruta):
    try:
        for elemento in os.listdir(ruta):
            ruta_completa = os.path.join(ruta, elemento)
            if os.path.isdir(ruta_completa):
                try:
                    contenido = os.listdir(ruta_completa)
                    if target in contenido:
                        return ruta_completa
                    resultado_recursivo = buscar_directorio(target, ruta_completa)
                    if resultado_recursivo:
                        return resultado_recursivo
                except PermissionError:
                    pass
        return None
    except PermissionError:
        return None

# Copia y pega aquí el contenido del script Ejecución.py
ruta_base = "C:\\Users"
carpeta_objetivo = "OneDrive - SGS"

ruta_carpeta_encontrada = buscar_directorio(carpeta_objetivo, ruta_base)

if ruta_carpeta_encontrada:
    ruta_archivo = os.path.join(ruta_carpeta_encontrada,carpeta_objetivo, "RIM_GENERAL")
    print(f"Ruta del archivo: {ruta_archivo}")
    logger.critical(f"Ruta del archivo: {ruta_archivo}")
else:
    print(f"No se encontró la carpeta {carpeta_objetivo}")
    logger.critical(f"No se encontró la carpeta {carpeta_objetivo}")


# Carpeta y nombre de archivo a buscar
carpeta = '//Pedb062/sites/ENV/PEENV03/Coll/StatusReports'
nombre_archivo = 'StatRep_PendientesAnalisis_'
destination_file = os.path.join(ruta_archivo, "StatRep_PendientesAnalisis.xlsx").replace("\\", "/")

# Obtener la fecha y hora actual
hora_actual = datetime.datetime.now() 

# Obtener la fecha actual
fecha_actual = hora_actual.date()

# Encontrar el rango de tiempo adecuado
hora_actual = hora_actual.time()
rango_seleccionado = None

rangos_horarios = [
    (datetime.time(2, 28), datetime.time(2, 34)),
    (datetime.time(5, 28), datetime.time(5, 34)),
    (datetime.time(8, 28), datetime.time(8, 34)),
    (datetime.time(11, 28), datetime.time(11, 34)),
    (datetime.time(14, 28), datetime.time(14, 34)),
    (datetime.time(17, 28), datetime.time(17, 34)),
    (datetime.time(20, 28), datetime.time(20, 34)),
    (datetime.time(23, 28), datetime.time(23, 34)),
]

for rango in rangos_horarios:
    if hora_actual < rango[0]:
        rango_seleccionado = rango
        break

rango_seleccionado = rangos_horarios[rangos_horarios.index(rango_seleccionado) - 1]

# Obtener la hora de inicio y fin del rango
hora_inicio, hora_fin = rango_seleccionado

# Lista para almacenar los archivos encontrados
archivos_encontrados = []

logger.critical(f"hora inicio {hora_inicio} hora fin {hora_fin}")
# Iterar a través de los archivos en la carpeta


for archivo in os.listdir(carpeta):
    if archivo.startswith(nombre_archivo):
        ruta_completa = os.path.join(carpeta, archivo)
        
        if os.path.isfile(ruta_completa):
           
            fecha_modificacion = datetime.datetime.fromtimestamp(os.path.getmtime(ruta_completa))
            
            if fecha_modificacion.date() == fecha_actual and hora_inicio <= fecha_modificacion.time() <= hora_fin:
                
                archivos_encontrados.append((archivo, os.path.getsize(ruta_completa)))

# Encontrar el archivo más pesado
# logger.critical("archivos_encontrados: " + ''.join(archivos_encontrados))

archivo_mas_pesado = max(archivos_encontrados, key=lambda x: x[1], default=None)
logger.critical(f"archivos_mas_pesado {archivo_mas_pesado}" )

if archivo_mas_pesado:
    nombre_archivo_mas_pesado, tamaño_mas_pesado = archivo_mas_pesado
    pendiente_file = os.path.join("//Pedb062/sites/ENV/PEENV03/Coll/StatusReports",nombre_archivo_mas_pesado).replace("\\", "/")
    # Preprocesamiento de los datos generados del RIM
    df = pd.read_excel(pendiente_file, skiprows= 2)
    columns_to_drop = [20]
    df = df.iloc[:, 1:].drop(columns=df.columns[columns_to_drop]).drop(index=0)
    # Verifica si el archivo de destino existe y crea el directorio si es necesario
    destination_dir = os.path.dirname(destination_file)
    os.makedirs(destination_dir, exist_ok=True)

    # Guarda los datos combinados directamente en el archivo final diario
    with pd.ExcelWriter(destination_file, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
        ws = writer.book.active
        # Define el ancho de las columnas
        for column in ws.columns:
            max_length = max(len(str(cell.value)) for cell in column)
            ws.column_dimensions[column[0].column_letter].width = max_length + 2
        # Crea la tabla en Excel
        tab = Table(displayName="Tabla_Pendientes", ref=ws.dimensions)
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=True)
        tab.tableStyleInfo = style
        ws.add_table(tab)
else:
    print(f"No se encontró el archivo {nombre_archivo} en el rango {hora_inicio} a {hora_fin} del día {fecha_actual}.")
    logger.critical(f"No se encontró el archivo {nombre_archivo} en el rango {hora_inicio} a {hora_fin} del día {fecha_actual}.")

logger.critical(" ############### end script execution ################")




