import zipfile
import pandas as pd
import numpy as np

import os
import openpyxl
import json
import shutil
import re

import subprocess
import unicodedata

from datetime import datetime

def descomprimir(lst_repositorios : np.array, verbose : bool = False):
    ok, not_ok = [], []
    for archivo_zip in lst_repositorios:
        # Directorio de destino para la extracción
        directorio_destino = archivo_zip.split("-Main.zip")[0]

        # Crear el directorio de destino si no existe
        if not os.path.exists(directorio_destino):
            os.makedirs(directorio_destino)

        # Descomprimir el archivo zip
        if not os.path.exists(archivo_zip):
            not_ok.append(archivo_zip.split("-Main.zip")[0])
            if verbose:
                print("Imposible descomprimir archivo {}".format(archivo_zip))
        else:
            with zipfile.ZipFile(archivo_zip, 'r') as zip_ref:
                zip_ref.extractall(directorio_destino)
                ok.append(archivo_zip.split("-Main.zip")[0])
            if verbose:
                print("El archivo ZIP se ha descomprimido exitosamente en:", directorio_destino)  
    return ok, not_ok

def calculate_percentage(data : pd.DataFrame, x_sede : str, x_seccion : str):
    s_query = f"sede == '{x_sede}' and seccion == '{x_seccion}' and estado_zip == 'OK'"
    count = data.query(s_query).shape[0]
    total = data.query(f"sede == '{x_sede}' and seccion == '{x_seccion}'").shape[0]
    try:
        percentage = (count / total)*100
    except ZeroDivisionError:
        percentage = 0
    return round(percentage)

def allowed_file(filename :str, extensions : set):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in extensions

def validate_file_teams(source : str):
    file_path = os.path.join("uploads", source)  
    lst_columns = ['sede', 'seccion', 'docente', 'rut_estudiante', 'equipo', 'link_github']
    columns_with_problems = []
    data_teams = pd.read_excel(file_path)
    data_teams.columns = data_teams.columns.str.lower().str.replace(" ", "_", regex=True)
    for column in lst_columns:
        if not column in data_teams.columns:
            columns_with_problems.append(column)
    return data_teams, columns_with_problems

def generar_nombre_unico(ruta: str) -> str:
    """Genera un nombre único si el archivo ya existe."""
    directorio, nombre_archivo = os.path.split(ruta)
    nombre, extension = os.path.splitext(nombre_archivo)
    contador = 1

    # Mientras exista un archivo con el mismo nombre, incrementa el contador.
    while os.path.exists(ruta):
        ruta = os.path.join(directorio, f"{nombre}_{contador}{extension}")
        contador += 1
    return ruta

def normalizar(texto: str) -> str:
    """Elimina tildes y reemplaza 'ñ' o 'Ñ' por 'n', convierte a minúsculas."""
    # Normaliza el texto para separar los diacríticos
    texto_normalizado = unicodedata.normalize('NFD', texto.lower())
    # Elimina los diacríticos (caracteres con categoría 'Mn')
    texto_sin_tildes = ''.join(c for c in texto_normalizado if unicodedata.category(c) != 'Mn')
    # Reemplaza 'ñ' y 'Ñ' por 'n'
    return texto_sin_tildes.replace('ñ', 'n').replace('Ñ', 'N').lower()

def renombrar_archivos_directorio(directorio):
    # Obtiene los archivos de la carpeta
    for root, _, files in os.walk(directorio):
        # Revisión de archivos
        for filename in files:
            nuevo_nombre = normalizar(filename.lower())
            ruta_vieja = os.path.join(root, filename.lower())
            ruta_nueva = os.path.join(root, nuevo_nombre)

            if ruta_vieja != ruta_nueva:  # Renombra sólo si hay cambios
                # Verifica si existe un archivo con el mismo nombre normalizado.
                if os.path.exists(ruta_nueva):
                    ruta_nueva = generar_nombre_unico(ruta_nueva)                
                # Renombra el archivo
                os.rename(ruta_vieja, ruta_nueva)
                #print(f'Renombrado: {ruta_vieja} -> {ruta_nueva}')

def clonar_repositorio(git_url : str, destino : str, logging):
    """
    Clona un repositorio Git dado el URL del repositorio y la ruta de destino.
    
    :param git_url: URL del repositorio Git
    :param destino: Ruta de destino donde se clonará el repositorio
    """
    if not os.path.exists(destino):
        os.makedirs(destino)
    else:
        # Borra el contenido del directorio
        for archivo in os.listdir(destino):
            ruta_archivo = os.path.join(destino, archivo)
            if os.path.isfile(ruta_archivo) or os.path.islink(ruta_archivo):
                os.unlink(ruta_archivo)  # Borra archivos o enlaces simbólicos
            elif os.path.isdir(ruta_archivo):
                shutil.rmtree(ruta_archivo)  # Borra subdirectorios
    
    try:
        # Ejecutar el comando git clone
        subprocess.run(["git", "clone", git_url, destino], check=True)
        # Elimina tildes y reemplaza las ñ por n de los nombres de archivos en el repositorio clonado
        renombrar_archivos_directorio(destino)

        logging.info(f"Repositorio clonado en {destino}")
        
        return True
    except subprocess.CalledProcessError as e:
        logging.error(f"Error al clonar el repositorio {destino}: {e}")
        return False

def obtener_ruta_directorio_real(ruta_directorio):
    """Obtiene la ruta real del directorio ignorando mayúsculas y minúsculas."""
    directorio_padre, nombre_directorio = os.path.split(ruta_directorio)
    
    # Si el directorio padre no existe, devuelve None
    if not os.path.isdir(directorio_padre):
        return None
    
    # Busca en el directorio padre ignorando mayúsculas
    for item in os.listdir(directorio_padre):
        if item.lower() == nombre_directorio.lower() and os.path.isdir(os.path.join(directorio_padre, item)):
            return os.path.join(directorio_padre, item)
    
    return None

def archivo_existe_independiente_mayusculas(ruta_directorio : str, nombre_archivo : str, logging):
    """Verifica si un archivo existe en el directorio ignorando mayúsculas/minúsculas."""
    # Obtiene la ruta correcta del directorio
    ruta_directorio_real = obtener_ruta_directorio_real(ruta_directorio)
    
    # Verifica que el directorio exista
    if ruta_directorio_real is None:        
        logging.error(f"No existe directorio: {ruta_directorio}")
        return False
    
    # Busca el archivo en el directorio con nombres en minúsculas
    nombre_archivo = nombre_archivo.lower()
    for archivo in os.listdir(ruta_directorio_real):
        if archivo.lower() == nombre_archivo:
            return True
    return False


def directorio_existe_ignorar_mayusculas(ruta_directorio : str):
    # Obtiene la ruta al directorio padre
    directorio_padre, nombre_directorio = os.path.split(ruta_directorio)
    
    # Si el directorio padre no existe, el directorio tampoco
    if not os.path.isdir(directorio_padre):
        return False
    
    # Busca en el directorio padre, ignorando mayúsculas
    for item in os.listdir(directorio_padre):
        if item.lower() == nombre_directorio.lower() and os.path.isdir(os.path.join(directorio_padre, item)):
            return True
    return False    

def revision_evidencias_individuales(source_file : str, source : pd.DataFrame, sede : str,
                                     seccion : str, equipo : str, fase : int, logging, verbose=False):
    file_path = os.path.join("generate", source_file)  
    evidencias = pd.read_excel(file_path, sheet_name="individuales")
    evidencias_grupales = pd.read_excel(file_path, sheet_name="grupales")

    # Filtra los estudiantes
    filtro = (
        (source['sede'] == sede) &
        (source['seccion'] == seccion) &
        (source['equipo'] == equipo)
    )
    # Definir las extensiones válidas
    extensiones_validas_doc = ['doc', 'pdf']
    extensiones_validas_planillas = ['xls']
    # Obtener el listado de estudiantes que cumplen con los criterios
    estudiantes_filtrados = source[filtro]['estudiante'].tolist()   
    logging.info(f"{sede}-{seccion}-{equipo} tiene {len(estudiantes_filtrados)} estudiantes")
    for numero, estudiante in enumerate(estudiantes_filtrados):
        estudiante = normalizar(estudiante)
        filtro_evidencias = (
            (evidencias['sede'] == sede) &
            (evidencias['seccion'] == seccion) &
            (evidencias['estudiante'] == estudiante.lower()) &
            (evidencias['fase'] == fase)
        )       
        evidencias_filtradas = evidencias[filtro_evidencias]['evidencia'].tolist()
        logging.info(f"Estudiante {numero+1}: {estudiante}")
        prefijo = f"descargas/{sede}/{seccion}/equipo-{equipo}"
        for x_evidencia in evidencias_filtradas:                        
            ruta_completa_archivo = os.path.join(f"{prefijo}/Fase {fase}/Evidencias individuales".lower(), x_evidencia.lower())
            # Verifica si existe la evidencia definida en el estándar            
            ruta_directorio = f"{prefijo}/Fase {fase}/Evidencias individuales"
            if not archivo_existe_independiente_mayusculas(ruta_directorio, x_evidencia, logging):
                evidencia_sin_extension = x_evidencia.lower().split(".")[0] + "." + x_evidencia.lower().split(".")[1]
                # Verifica si se trata de un documento                
                todo_correcto = False
                for ext in extensiones_validas_doc:
                    #ruta_directorio = f"{prefijo}/Fase {fase}/Evidencias individuales".lower()
                    other_name = f"{evidencia_sin_extension}.{ext}"
                    ruta_completa_archivo_n = os.path.join(f"{prefijo}/Fase {fase}/Evidencias individuales".lower(), other_name.lower())
                    
                    if archivo_existe_independiente_mayusculas(ruta_directorio, other_name, logging):                    
                        todo_correcto = True     
                        logging.info(f"Archivo {ruta_completa_archivo_n} encontrado con extensión {ext}")                            
                        evidencias.iloc[evidencias.index[evidencias['evidencia'] == x_evidencia.lower()].tolist()[0],7] = "OK"                       
                        break 
                if not todo_correcto:
                    logging.error(f"Archivo faltante {ruta_completa_archivo} (con ninguna extensión posible)")                                       
            else:
                todo_correcto = True
                logging.info(f"Archivo encontrado {ruta_completa_archivo}")

                evidencias.iloc[evidencias.index[evidencias['evidencia'] == x_evidencia.lower()].tolist()[0],7] = "OK"            
    
    with pd.ExcelWriter(file_path) as writer:
        evidencias.to_excel(writer, sheet_name="individuales", index=False)        
        evidencias_grupales.to_excel(writer, sheet_name="grupales", index=False)

def verificar_archivo_presentacion(directorio : str):
    # Expresión regular para archivos que comiencen con la frase y tengan las extensiones .pdf, .pptx, o .ppt
    patron = re.compile(r'^presentacion\s+idea\s+de\s+proyecto.*\.(pdf|pptx|ppt)$', re.IGNORECASE)
    try:
        for archivo in os.listdir(directorio):
            archivo_normalizado = normalizar(archivo.lower())  # Normaliza para quitar tildes y pasar a minúsculas
            if patron.match(archivo_normalizado):  # Usamos match para verificar al inicio
                return True
    except FileNotFoundError:
        return False

def revision_evidencias_grupales(source_file : str, source : pd.DataFrame, sede : str,
                                     seccion : str, equipo : str, fase : int, logging, verbose = False):
    
    file_path = os.path.join("generate", source_file)  
    evidencias = pd.read_excel(file_path, sheet_name="individuales")
    evidencias_grupales = pd.read_excel(file_path, sheet_name="grupales")

    extensiones_validas_planillas = ['.xlsx', '.xls']
    
    # Filtro de los equipos
    filtro = (
        (source['sede'] == sede) &
        (source['seccion'] == seccion) &
        (source['equipo'] == equipo)
    )

    filtro_evidencias = (
            (evidencias_grupales['sede'] == sede) &
            (evidencias_grupales['seccion'] == seccion) &
            (evidencias_grupales['equipo'] == equipo) &
            (evidencias_grupales['fase'] == fase)
        )  
    # Filtra las evidencias del equipo y la fase
    evidencias_equipo = evidencias_grupales[filtro_evidencias]['evidencia'].tolist()
    prefijo = f"descargas/{sede}/{seccion}/equipo-{equipo}"
    ruta_directorio = f"{prefijo}/Fase {fase}/Evidencias grupales"
    for x_evidencia in evidencias_equipo:
        # Verifica si se trata de la planilla de notas
        if x_evidencia.lower().startswith("planilla"):                        
            ruta_completa_archivo = os.path.join(f"{prefijo}/Fase {fase}/Evidencias grupales".lower(), x_evidencia.lower())
            if not archivo_existe_independiente_mayusculas(ruta_directorio, x_evidencia, logging):
                evidencia_sin_extension = x_evidencia.split(".")[0]
                if x_evidencia.lower().endswith("xlsx"):
                    todo_correcto = False
                    for ext in extensiones_validas_planillas:
                        other_name = f"{evidencia_sin_extension}.{ext}"
                        ruta_completa_archivo_n = os.path.join(f"{prefijo}/Fase {fase}/Evidencias grupales".lower(), other_name.lower())
                        if archivo_existe_independiente_mayusculas(ruta_directorio, other_name, logging):
                            todo_correcto = True                            
                            break                                                          
                if todo_correcto:
                    # Marca la evidencia como entregada
                    evidencias_grupales.iloc[evidencias_grupales.index[((evidencias_grupales['evidencia'] == x_evidencia.lower()) &
                                                                     (evidencias_grupales['equipo'] == equipo) &
                                                               (evidencias_grupales['sede'] == sede) &
                                                               (evidencias_grupales['seccion'] == seccion))].tolist()[0],6] = "OK"            

            else: 
                todo_correcto = True        
                # Marca la evidencia como entregada
                evidencias_grupales.iloc[evidencias_grupales.index[((evidencias_grupales['evidencia'] == x_evidencia.lower()) &
                                                                     (evidencias_grupales['equipo'] == equipo) &
                                                               (evidencias_grupales['sede'] == sede) &
                                                               (evidencias_grupales['seccion'] == seccion))].tolist()[0],6] = "OK"                    
        # Verifica si se trata del archivo de la presentación
        if x_evidencia.lower().startswith("presentación") or x_evidencia.lower().startswith("presentacion"):            
            if verificar_archivo_presentacion(ruta_directorio):
                logging.info(f"{ruta_directorio}/{x_evidencia} (pdf | pptx | ppt) encontrado")
                # Marca la evidencia como entregada
                evidencias_grupales.iloc[evidencias_grupales.index[((evidencias_grupales['evidencia'] == x_evidencia.lower()) &
                                                                     (evidencias_grupales['equipo'] == equipo) &
                                                               (evidencias_grupales['sede'] == sede) &
                                                               (evidencias_grupales['seccion'] == seccion))].tolist()[0],6] = "OK"                 
            else:
                logging.error(f"'{ruta_directorio}/{x_evidencia}' (pdf | pptx | ppt) faltante")
        # Verifica si se trata de los documentos
        if x_evidencia.lower().endswith("docx"): 
            todo_correcto = False
            if x_evidencia.lower().endswith("(espanol).docx"):
                # Busca con el patron
                if x_evidencia.startswith("1.5"):
                    patron = re.compile(r'^1.5_GuiaEstudiante_Fase\s+1_Definicion\s+proyecto\s+apt.*\.(pdf|docx|doc)$', re.IGNORECASE)
                if x_evidencia.startswith("2.4"):
                    # 2.4_GuiaEstudiante_Fase 2_DesarrolloProyecto APT
                    patron = re.compile(r'^2.4_GuiaEstudiante_Fase\s+2_DesarrolloProyecto\s+apt.*\.(pdf|docx|doc)$', re.IGNORECASE)                                    
                if x_evidencia.startswith("2.6"):
                    # 2.6_GuiaEstudiante_Fase 2_Informe Final Proyecto APT (Espanol).docx
                    patron = re.compile(r'^2.6_GuiaEstudiante_Fase\s+2_informe\s+final\s+proyecto\s+apt.*\.(pdf|docx|doc)$', re.IGNORECASE)                                    

                try:
                    ruta_completa_archivo = os.path.join(f"{prefijo}/Fase {fase}/Evidencias grupales".lower(), x_evidencia.lower())
                    for archivo in os.listdir(ruta_directorio):
                        archivo_normalizado = normalizar(archivo.lower())  # Normaliza para quitar tildes y pasar a minúsculas
                        if patron.match(archivo_normalizado):  # Usamos match para verificar al inicio
                            logging.info(f"Archivo {ruta_completa_archivo} encontrado con o sin el sufijo espanol")                            
                            # Marca la evidencia como entregada
                            evidencias_grupales.iloc[evidencias_grupales.index[((evidencias_grupales['evidencia'] == x_evidencia.lower()) &
                                                                     (evidencias_grupales['equipo'] == equipo) &
                                                               (evidencias_grupales['sede'] == sede) &
                                                               (evidencias_grupales['seccion'] == seccion))].tolist()[0],6] = "OK"
                            todo_correcto = True
                            break
                except FileNotFoundError:
                    pass  
                if not todo_correcto:
                    logging.error(f"Archivo {ruta_completa_archivo} faltante")        
            else:
                
                ruta_completa_archivo = os.path.join(f"{prefijo}/Fase {fase}/Evidencias grupales".lower(), x_evidencia.lower())
                extensiones_validas_doc = ['doc', 'pdf']
                if archivo_existe_independiente_mayusculas(ruta_directorio, x_evidencia, logging):
                    logging.info(f"Archivo {ruta_completa_archivo} encontrado")                            
                    # Marca la evidencia como entregada
                    evidencias_grupales.iloc[evidencias_grupales.index[((evidencias_grupales['evidencia'] == x_evidencia.lower()) &
                                                                     (evidencias_grupales['equipo'] == equipo) &
                                                               (evidencias_grupales['sede'] == sede) &
                                                               (evidencias_grupales['seccion'] == seccion))].tolist()[0],6] = "OK"
                else:
                    todo_correcto = False
                    evidencia_sin_extension = x_evidencia.split(".")[0]
                    for ext in extensiones_validas_doc:
                        other_name = f"{evidencia_sin_extension}.{ext}"
                        ruta_completa_archivo_n = os.path.join(f"{prefijo}/Fase {fase}/Evidencias grupales".lower(), other_name.lower())
                    
                        if archivo_existe_independiente_mayusculas(ruta_directorio, other_name, logging):                    
                            todo_correcto = True     
                            logging.info(f"Archivo {ruta_completa_archivo_n} encontrado con extensión {ext}")                            
                            evidencias.iloc[evidencias.index[evidencias['evidencia'] == x_evidencia.lower()].tolist()[0],7] = "OK"                       
                            break 
                    if not todo_correcto:
                        logging.error(f"Archivo {ruta_completa_archivo} faltante")     
                
    with pd.ExcelWriter(file_path) as writer:
        evidencias.to_excel(writer, sheet_name="individuales", index=False)        
        evidencias_grupales.to_excel(writer, sheet_name="grupales", index=False)
    
    # Calcula los avances
    # Evidencias grupales
    calcula_avances(evidencias_grupales, sede)
    # Evidencias individuales
    calcula_avances_por_estudiante(evidencias, sede)
    
    return True

def revision_repositorio(source : pd.DataFrame, sede : str, logging, clonar : bool, verbose = False):
    """
    Revisa el estado de los repositorios contenidos en el dataframe indicado

    :param source: dataframe que contiene datos de los equipos incluyendo los nombres de los repositorios
    """
    # Log de resultados
    log_repositorios = []
    # Agrupar por 'sede', 'docente', 'sección' y 'equipo'
    result = source.groupby(['sede', 'docente', 'seccion', 'equipo'], as_index=False).agg(
        cantidad_estudiantes=('estudiante', 'count'),
        enlace=('link_github', 'first')
    )

    # Mostrar el resultado
    #print(result)    
        
    total, total_sin_informar, total_estructura_correcta = 0, 0, 0
    
    for index, row in result.iterrows():
        sede = row['sede']
        seccion = row['seccion']
        equipo = row['equipo']
        repositorio = row['enlace']
                    
        # Comprobar si la columna de repositorio es NaN y que se tenga número de equipo        
        if not pd.isna(repositorio) and equipo > 0:
            if verbose:
                print(f"\tSede: {sede}, Sección {seccion}, Equipo: {equipo}, Repositorio: {repositorio}")
            path_destino = f"descargas/{sede}/{seccion}/equipo-{equipo}"
            clonado = True
            if clonar and clonar_repositorio(repositorio, path_destino, logging):
                log_repositorios.append([repositorio, path_destino])
                # Log de éxito
                logging.info(f"Repositorio clonado y revisado con éxito en {path_destino}")
            elif clonar:
                log_repositorios.append([repositorio, f'Error al clonar: {seccion}/equipo-{equipo}'])
                git_error = f"{seccion}/equipo-{equipo}"
                logging.error(f"Error al clonar el repositorio {git_error} en {path_destino}")
                clonado = False
            if not clonado:
                continue
            # Revisión de evidencias por fase
            for fase in range(1,4):
                if revision_evidencias_individuales(f"resumen_evidencias_{sede}.xlsx", source, 
                                                        sede, seccion, equipo, fase, logging, verbose):
                    total_estructura_correcta+=1
                total+=1
                # Revisión de evidencias grupales
                if revision_evidencias_grupales(f"resumen_evidencias_{sede}.xlsx", source, sede, 
                                                seccion, equipo, fase, logging, verbose):
                    total_estructura_correcta+=1
        else:
            total_sin_informar+=1

    return total, total_sin_informar, source.query("equipo == 0")[['sede', 'seccion','docente', 'rut_estudiante', 'estudiante']], log_repositorios    

def calcula_avances(data_evidencias : pd.DataFrame, sede : str):
    # Agrupar por las columnas deseadas y contar la cantidad de "OK" y "NO"
    resultado = data_evidencias.groupby(['sede', 'seccion', 'docente', 'equipo', 'fase', 'estado']).size().unstack(fill_value=0).reset_index()

    # Calcular el total de registros para cada grupo
    # Verifica que exista la columna OK para el caso de que exista ninguna evidencia
    if "OK" not in resultado.columns: 
        resultado['OK'] = 0
    
    resultado['Total'] = resultado['OK'] + resultado['NO']
        
    # Calcular el porcentaje de OK y NO
    resultado['% OK'] = round(resultado['OK'] / resultado['Total'], 2) #* 100
    resultado['% NO'] = round(resultado['NO'] / resultado['Total'],2) #* 100

    # Agrega metadata del reporte
    # Obtener la fecha y hora actual
    fecha_emision = datetime.now().strftime("%Y-%m-%d")
    hora_emision = datetime.now().strftime("%H:%M:%S")
    # Agregar una hoja con los metadatos
    metadatos_data = {
        "Campo": ["Fecha de emisión", "Hora de emisión"],
        "Valor": [fecha_emision, hora_emision]
    }
    metadatos = pd.DataFrame(metadatos_data)  
    # Mostrar el resultado
    #print(resultado[['sede', 'seccion', 'docente', 'equipo', 'fase', '% OK', '% NO']])
    file_path = os.path.join("generate", f"reporte_evidencias_equipos_{sede}.xlsx")
    with pd.ExcelWriter(file_path) as reporte:
        metadatos.to_excel(reporte, sheet_name="Metadata", index=False)
        resultado.to_excel(reporte, sheet_name="reporte", index=False)

def calcula_avances_por_estudiante(data_evidencias : pd.DataFrame, sede : str):
    # Obtiene los desertores
    desertores = []
    # Agrupar por las columnas deseadas y contar la cantidad de "OK" y "NO"
    resultado = data_evidencias.groupby(['sede', 'seccion', 'docente', 'estudiante', 
                                         'fase', 'estado']).size().unstack(fill_value=0).reset_index()

    if "OK" not in resultado.columns: 
        resultado['OK'] = 0

    # Calcular el total de registros para cada grupo
    resultado['Total'] = resultado['OK'] + resultado['NO']

    # Calcular el porcentaje de OK y NO
    resultado['% OK'] = round((resultado['OK'] / resultado['Total']),2) #* 100
    resultado['% NO'] = round((resultado['NO'] / resultado['Total']),2) #* 100

    # Mostrar el resultado
    #print(resultado[['sede', 'seccion', 'docente', 'estudiante', 'fase', '% OK', '% NO']])
    # Obtener la fecha y hora actual
    fecha_emision = datetime.now().strftime("%Y-%m-%d")
    hora_emision = datetime.now().strftime("%H:%M:%S")
    # Agregar una hoja con los metadatos
    metadatos_data = {
        "Campo": ["Fecha de emisión", "Hora de emisión"],
        "Valor": [fecha_emision, hora_emision]
    }
    metadatos = pd.DataFrame(metadatos_data)     
    file_path = os.path.join("generate", f"reporte_evidencias_individuales_{sede}.xlsx")
    with pd.ExcelWriter(file_path) as reporte:
        metadatos.to_excel(reporte, sheet_name="Metadata", index=False)
        resultado.to_excel(reporte, sheet_name="reporte", index=False)
        
        #desertores[['sede', 'seccion', 'docente', 'estudiante']].to_excel(reporte, sheet_name="desertores", index=False)

def reporte_desertores(data_desertores : pd.DataFrame, filename :str):
    file_path = os.path.join("generate", filename)
    with pd.ExcelWriter(file_path) as reporte:
        data_desertores.to_excel(reporte, sheet_name="desertores", index=False)
    return data_desertores.shape[0], file_path
    
def reporte_git(data_git : np.array, filename :str):
    data_repositorios = pd.DataFrame(data_git)
    data_repositorios.columns = ['enlace', 'estado']
    file_path = os.path.join("generate", filename)
    with pd.ExcelWriter(file_path) as reporte:
        data_repositorios.to_excel(reporte, sheet_name="estado_descarga", index=False)
    return data_repositorios.shape[0], file_path