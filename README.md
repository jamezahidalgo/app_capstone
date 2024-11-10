# app_capstone
Automatización generación planillas

## Funcionalidad de app_git

Paso 1 - Tener la planilla de los datos de los números de equipos y enlaces de repositorios en una carpeta llamada uploads

Paso 2 - Tener el archivo config.json en una carpeta llamada config

Paso 3 - Activar el entorno

Paso 4 - Ejecutar desde la terminal 

$ python app_git.py --source archivo.xlsx --config config.json

En este caso archivo.xlsx es el nombre del archivo donde se encuentra la información de los equipos.

En caso de que se requiera mostrar el nombre de la sede, sección y equipo en la terminal puede ejecutar:

$ python app_git.py --source archivo.xlsx --config config.json --verbose=True

En caso de que se requiera sólo revisar sin descargar los repositorios puede ejecutar:

$ python app_git.py --source archivo.xlsx --config config.json --action=no

Paso 5 - Verá una serie de mensajes LOGS en la terminal 

El archivo generado se llama resumen_evidencias.xlsx y se encuentra en la carpeta generate

Al final se genera un archivo log con fecha y hora. Nombre del archivo: proceso_clonacion<fecha y hora>.log
