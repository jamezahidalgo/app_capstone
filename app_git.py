import logic.funciones as fn
import logic.generate as generate

import argparse
import os
import json

# Crear el parser
parser = argparse.ArgumentParser(description="APP capstone desde la línea de comandos.")

# Agregar argumentos
parser.add_argument('--config', type=str, help='archivo de configuración')
parser.add_argument('--source', type=str, help='archivo con datos de repositorios')
parser.add_argument('--type', type=str, default="download", help='view | download | report')
parser.add_argument('--verbose', type=bool, default=False, help='True | False')

# Parsear los argumentos
args = parser.parse_args()

# Validar la existencia de parámetros
if args.source is None:
    parser.error("--source es un parámetro requerido")

# Acceder a los argumentos
print(f"Archivo fuente de los equipos: {args.source}")
print(f"Tipo ejecucion: {args.type}")

archivo = args.source
ruta_archivo = os.path.join("uploads", archivo)
archivo_config = args.config
ruta_archivo_config = os.path.join("config", archivo_config)

if os.path.exists(ruta_archivo_config):
    # Abrir y leer el archivo JSON
    with open(ruta_archivo_config, 'r') as file:
        datos = json.load(file)

# Lee archivo y comprueba las columnas
teams, columns_problem = fn.validate_file_teams(archivo)
if len(columns_problem) == 0:
    print("log: Archivo de equipos válido")
    mensajes = generate.generate_equipos(datos['upload_folder'], archivo, 
                                         datos["generate_folder"])
    for mensaje in mensajes:
        print("log: ", mensaje)    

    total, sin_informar, desertores = fn.revision_repositorio(teams)

    print(f"log: Se han descargado exitosamente {total} repositorios")
    print(f"log: Hay {sin_informar} equipos sin informar repositorio")
    
    # Genera reporte de desertores
    total, salida = fn.reporte_desertores(desertores)
    print(f"log: Reporte con {total} desertores generado existosamente en archivo {salida}")
    
else:
    print(f"Archivo de equipos inválido, {columns_problem} son las columnas con problemas")
