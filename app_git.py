import logic.funciones as fn
import logic.generate as generate

import argparse
import os
import json
import logging

from datetime import datetime

# Configuración del logger
# Crear el nombre del archivo con fecha y hora
nombre_archivo = f"proceso_clonacion_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

logging.basicConfig(
    filename=nombre_archivo,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# Crear el parser
parser = argparse.ArgumentParser(description="APP capstone desde la línea de comandos.")

# Agregar argumentos
parser.add_argument('--config', type=str, help='archivo de configuración')
parser.add_argument('--source', type=str, help='archivo con datos de repositorios')
parser.add_argument('--type', type=str, default="download", help='view | download | report')
parser.add_argument('--verbose', type=bool, default=False, help='True | False')
parser.add_argument('--action', type=str, default="si", help='si | no')

# Parsear los argumentos
args = parser.parse_args()

# Validar la existencia de parámetros
if args.source is None:
    parser.error("--source es un parámetro requerido")

# Acceder a los argumentos
print(f"Archivo fuente de los equipos: {args.source}")
print(f"Tipo ejecucion: {args.type}")
print(f"Acción clonación: {args.action}")

archivo = args.source
ruta_archivo = os.path.join("uploads", archivo)
archivo_config = args.config
ruta_archivo_config = os.path.join("config", archivo_config)
with_clone = True if args.action == "si" else False

if os.path.exists(ruta_archivo_config):
    # Abrir y leer el archivo JSON
    with open(ruta_archivo_config, 'r') as file:
        datos = json.load(file)

# Lee archivo y comprueba las columnas
teams, columns_problem = fn.validate_file_teams(archivo)
if len(columns_problem) == 0:
    print("log: Archivo de equipos válido")
    sede = archivo.split(".")[0].upper()
    mensajes = generate.generate_equipos(datos['upload_folder'], archivo, 
                                         datos["generate_folder"], sede)
    for mensaje in mensajes:
        print("log: ", mensaje)    

    print(f"log: Generando reportes de sede {sede}...")
    total, sin_informar, desertores, log_git = fn.revision_repositorio(teams, sede, logging, 
                                                                       clonar = with_clone,
                                                                       verbose = args.verbose)

    print(f"log: Se han descargado exitosamente {total} repositorios")
    print(f"log: Hay {sin_informar} equipos sin informar repositorio")
    
    # Genera reporte de desertores
    filename_report = f"reporte_desertores_{sede}.xlsx"
    total, salida = fn.reporte_desertores(desertores, filename_report)
    print(f"log: Reporte con {total} desertores generado existosamente en archivo {salida}")
    
    # Reporte con la descarga de los repositorios
    try:
        report_git = f"reporte_git_{sede}.xlsx"
        total, salida = fn.reporte_git(log_git, report_git)
        print(f"log: Reporte con {total} repositorios generado existosamente en archivo {salida}")
    except ValueError:
        print(f"log: Error al generar reporte git")
    
    print(f"log: Archivo de log en {nombre_archivo}")
else:
    print(f"Archivo de equipos inválido, {columns_problem} son las columnas con problemas")
