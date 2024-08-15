import logic.funciones as fn
import logic.generate as generate

import argparse
import os
import json

# Crear el parser
parser = argparse.ArgumentParser(description="APP capstone desde la línea de comandos.")

# Agregar argumentos
parser.add_argument('--config', type=str, help='archivo de configuración')
parser.add_argument('--type', type=str, default="teams", help='init | teams | report')

# Parsear los argumentos
args = parser.parse_args()

# Validar la existencia de parámetros
if args.config is None:
    parser.error("--config es un parámetro requerido")

# Acceder a los argumentos
print(f"Carpeta carga: {args.config}")
print(f"Tipo ejecucion: {args.type}")
archivo = args.config
ruta_archivo = os.path.join("config", archivo)

# Verificar si el archivo existe
if os.path.exists(ruta_archivo):
    # Abrir y leer el archivo JSON
    with open(ruta_archivo, 'r') as file:
        datos = json.load(file)

    if args.type == "init":
        print("Generando planillas de equipos  ...")
        mensajes = generate.generate_files(datos["upload_folder"], datos["archivo_inscritos"],
                                           datos["generate_folder"],)
        for mensaje in mensajes:
            print(mensaje)
    elif args.type == "teams":
        print("Generando evidencias a partir de los equipos informados ...")    
        # Genera los equipos
        mensajes = generate.generate_equipos(datos['upload_folder'], datos["archivo_equipos"], 
                                         datos["generate_folder"])
        for mensaje in mensajes:
            print(mensaje)
    elif args.type == "report":    
        print("Generando reportes ...")
        for mensaje in generate.generate_summary(datos['upload_folder'], datos["generate_folder"],
                                     datos["archivo_equipos"]):
            print(mensaje)

