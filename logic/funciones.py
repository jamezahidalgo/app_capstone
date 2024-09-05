import zipfile
import pandas as pd
import numpy as np

import os
import openpyxl
import json
import shutil

import subprocess

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

def clonar_repositorio(git_url : str, destino : str):
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
        print(f"Repositorio clonado en {destino}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error al clonar el repositorio: {e}")
        return False

def revision_evidencias_individuales(source_file : str, source : pd.DataFrame, sede : str,
                                     seccion : str, equipo : str, fase : int, verbose=False):
    file_path = os.path.join("generate", source_file)  
    evidencias = pd.read_excel(file_path, sheet_name="individuales")
    evidencias_grupales = pd.read_excel(file_path, sheet_name="grupales")

    # Filtra los estudiantes
    filtro = (
        (source['sede'] == sede) &
        (source['seccion'] == seccion) &
        (source['equipo'] == equipo)
    )
    # Obtener el listado de estudiantes que cumplen con los criterios
    estudiantes_filtrados = source[filtro]['estudiante'].tolist()   
    for estudiante in estudiantes_filtrados:
        filtro_evidencias = (
            (evidencias['sede'] == sede) &
            (evidencias['seccion'] == seccion) &
            (evidencias['estudiante'] == estudiante) &
            (evidencias['fase'] == fase)
        )       
        evidencias_filtradas = evidencias[filtro_evidencias]['evidencia'].tolist()
        prefijo = f"descargas/{sede}/{seccion}/equipo-{equipo}"
        for x_evidencia in evidencias_filtradas:
            ruta_completa_archivo = os.path.join(f"{prefijo}/Fase {fase}/Evidencias individuales".lower(), x_evidencia.lower())
            if not os.path.exists(ruta_completa_archivo):
                if verbose:
                    print(f"Archivo faltante: {ruta_completa_archivo}")
                todo_correcto = False            
            else:
                if verbose:
                    print(f"Archivo OK: {ruta_completa_archivo}")
                # Marca la evidencia como entregada
                evidencias.iloc[evidencias.index[evidencias['evidencia'] == x_evidencia].tolist()[0],7] = "OK"
    
    with pd.ExcelWriter(file_path) as writer:
        evidencias.to_excel(writer, sheet_name="individuales", index=False)        
        evidencias_grupales.to_excel(writer, sheet_name="grupales", index=False)

def revision_evidencias_grupales(source_file : str, source : pd.DataFrame, sede : str,
                                     seccion : str, equipo : str, fase : int, verbose = False):
    
    file_path = os.path.join("generate", source_file)  
    evidencias = pd.read_excel(file_path, sheet_name="individuales")
    evidencias_grupales = pd.read_excel(file_path, sheet_name="grupales")

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
    for x_evidencia in evidencias_equipo:
        # Verifica si se trata del archivo de la presentación
        if x_evidencia.startswith("Presentación"):
            archivo_sin_extension = x_evidencia
            for extension in ['.pdf', '.pptx']:
                x_evidencia = archivo_sin_extension + extension   
                ruta_completa_archivo = os.path.join(f"{prefijo}/Fase {fase}/Evidencias grupales".lower(), x_evidencia.lower())
                if not os.path.exists(ruta_completa_archivo):
                    if verbose: 
                        print(f"Archivo faltante: {ruta_completa_archivo}")
                    todo_correcto = False            
                else:
                    if verbose: 
                        print(f"Archivo OK: {ruta_completa_archivo}")
                    # Marca la evidencia como entregada
                    evidencias_grupales.iloc[evidencias_grupales.index[((evidencias_grupales['evidencia'] == archivo_sin_extension) &
                                                                     (evidencias_grupales['equipo'] == equipo) &
                                                               (evidencias_grupales['sede'] == sede) &
                                                               (evidencias_grupales['seccion'] == seccion))].tolist()[0],6] = "OK"                 
                    
                    break
        else:    
            ruta_completa_archivo = os.path.join(f"{prefijo}/Fase {fase}/Evidencias grupales".lower(), x_evidencia.lower())
            if not os.path.exists(ruta_completa_archivo):
                if verbose:
                    print(f"Archivo faltante: {ruta_completa_archivo}")
                todo_correcto = False            
            else:
                if verbose:
                    print(f"Archivo OK: {ruta_completa_archivo}")
                # Marca la evidencia como entregada
                evidencias_grupales.iloc[evidencias_grupales.index[((evidencias_grupales['evidencia'] == x_evidencia) &
                                                                     (evidencias_grupales['equipo'] == equipo) &
                                                               (evidencias_grupales['sede'] == sede) &
                                                               (evidencias_grupales['seccion'] == seccion))].tolist()[0],6] = "OK"            
    
    with pd.ExcelWriter(file_path) as writer:
        evidencias.to_excel(writer, sheet_name="individuales", index=False)        
        evidencias_grupales.to_excel(writer, sheet_name="grupales", index=False)
    
    # Calcula los avances
    # Evidencias grupales
    calcula_avances(evidencias_grupales)
    # Evidencias individuales
    calcula_avances_por_estudiante(evidencias)
    
    return True

def revision_repositorio(source : pd.DataFrame, verbose = False):
    """
    Revisa el estado de los repositorios contenidos en el dataframe indicado

    :param source: dataframe que contiene datos de los equipos incluyendo los nombres de los repositorios
    """
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
                    
        # Comprobar si la columna de repositorio es NaN
        
        if not pd.isna(repositorio):
            if verbose:
                print(f"Sede: {sede}, Sección {seccion}, Equipo: {equipo}, Repositorio: {repositorio}")
            path_destino = f"descargas/{sede}/{seccion}/equipo-{equipo}"
            if clonar_repositorio(repositorio, path_destino):
                if revision_evidencias_individuales("resumen_evidencias.xlsx", source, sede, seccion, equipo,1, verbose):
                    total_estructura_correcta+=1
                total+=1
                # Revisión de evidencias grupales
                if revision_evidencias_grupales("resumen_evidencias.xlsx", source, sede, seccion, equipo,1, verbose):
                    total_estructura_correcta+=1
        else:
            total_sin_informar+=1
    

    return total, total_sin_informar, source.query("equipo == 0")[['sede', 'seccion','docente', 'rut_estudiante', 'estudiante']]    

def calcula_avances(data_evidencias : pd.DataFrame):
    # Agrupar por las columnas deseadas y contar la cantidad de "OK" y "NO"
    resultado = data_evidencias.groupby(['sede', 'seccion', 'docente', 'equipo', 'fase', 'estado']).size().unstack(fill_value=0).reset_index()

    # Calcular el total de registros para cada grupo
    # Verifica que exista la columna OK para el caso de que exista ninguna evidencia
    if "OK" not in resultado.columns: 
        resultado['OK'] = 0
    
    resultado['Total'] = resultado['OK'] + resultado['NO']
    
    #print(resultado.query("seccion == 'PTY4614-006'"))
    
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
    file_path = os.path.join("generate", "reporte_evidencias_equipos.xlsx")
    with pd.ExcelWriter(file_path) as reporte:
        metadatos.to_excel(reporte, sheet_name="Metadata", index=False)
        resultado.to_excel(reporte, sheet_name="reporte", index=False)

def calcula_avances_por_estudiante(data_evidencias : pd.DataFrame):
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
    file_path = os.path.join("generate", "reporte_evidencias_individuales.xlsx")
    with pd.ExcelWriter(file_path) as reporte:
        metadatos.to_excel(reporte, sheet_name="Metadata", index=False)
        resultado.to_excel(reporte, sheet_name="reporte", index=False)
        
        #desertores[['sede', 'seccion', 'docente', 'estudiante']].to_excel(reporte, sheet_name="desertores", index=False)

def reporte_desertores(data_desertores : pd.DataFrame, filename = "reporte_desertores.xlsx"):
    file_path = os.path.join("generate", filename)
    with pd.ExcelWriter(file_path) as reporte:
        data_desertores.to_excel(reporte, sheet_name="desertores", index=False)
    return data_desertores.shape[0], file_path
    
