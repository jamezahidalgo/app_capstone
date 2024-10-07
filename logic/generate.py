import pandas as pd
import numpy as np

import os
import openpyxl
import json

from datetime import datetime

from logic.funciones import descomprimir, calculate_percentage

def generate_equipos(folder : str, archivo_equipos : str, generate_folder : str, sede : str):
    lst_messages = []
    # Genera de planilla de los equipos
    file_path = os.path.join(folder, archivo_equipos)
    lst_messages.append(f"Se han cargado los datos desde {file_path}")
    equipos = pd.read_excel(file_path)
    n_desertores = equipos.query("equipo == 0").shape[0]
    n_sedes = pd.DataFrame(equipos.groupby("sede")).shape[0]
    lst_messages.append(f"Hay {n_desertores} desertores registrados")
    lst_messages.append(f"Hay {n_sedes} sedes cargadas")
    lst_messages.append(f"Hay {equipos.shape[0]} estudiantes cargados en total")
    equipos = equipos.query("equipo > 0")
    resumen_equipos = pd.DataFrame(equipos.groupby(["sede","seccion","equipo"]).size()).reset_index()
    resumen_equipos.columns.values[-1] = 'total_integrantes'
    resumen_equipos['git'] = "agregar aca enlace"
  
    #resumen_equipos['codigo'] = resumen_equipos.apply(lambda x : "{}-{}-{}".format(x["sede"], 
    #                                                        x["seccion"], 
    #                                                        f"EQUIPO-{x['equipo']}"), axis=1)
    resumen_equipos['estado_zip'] = "NO"
    file_path = os.path.join(generate_folder, "resumen_equipos.xlsx")
    resumen_equipos.to_excel(file_path)
    lst_messages.append(f"Actualizado resumen equipos en /{file_path}")
    
    # Resumen de evidencias individuales
    evidencias_F1 = ['AutoevaluacionCompetenciasFase1', 'DiarioReflexionFase1', 
                  'AutoevaluacionFase1']
    
    evidencias_F2 = ['DiarioReflexionFase2']
    evidencias_F3 = ['DiarioReflexionFase3']

    lst_evidencias = []
    for sede in equipos['sede'].unique():
        for seccion in equipos.query("sede == '{}'".format(sede))['seccion'].unique():        
            s_query = f"sede == '{sede}' and seccion == '{seccion}'"
            docente = equipos.query(f"sede == '{sede}' and seccion == '{seccion}'")['docente'].unique()
            carpeta = f"{sede}-{seccion}-{docente[0]}"
            
            for estudiante in equipos.query(s_query)[['rut_estudiante', 'estudiante']].values.tolist(): 
                rut = estudiante[0]           
                nombre_completo = estudiante[1]
                if len(nombre_completo.split(" ")) == 2:
                    nombre_para_archivo = nombre_completo.split(" ")[0] + "_" + nombre_completo.split(" ")[1]
                else:
                    nombre_para_archivo = nombre_completo.split(" ")[0] + "_" + nombre_completo.split(" ")[2]
                n_indice = 1
                for evidencia in evidencias_F1:
                    #nombre_evidencia = f"{rut}_1.1_APT122_{evidencia}.docx"
                    nombre_evidencia = f"{nombre_para_archivo}_1.{n_indice}_APT122_{evidencia}.docx"
                    registro = [sede, seccion, docente[0], rut, nombre_completo, 1, nombre_evidencia, "NO"]
                    lst_evidencias.append(registro)
                    n_indice += 1
                n_indice = 1
                for evidencia in evidencias_F2:
                    #nombre_evidencia = f"{rut}_2.1_APT122_{evidencia}.docx"
                    nombre_evidencia = f"{nombre_para_archivo}_2.{n_indice}_APT122_{evidencia}.docx"
                    registro = [sede, seccion, docente[0], rut, nombre_completo, 2, nombre_evidencia, "NO"]
                    lst_evidencias.append(registro)
                    n_indice += 1
                n_indice = 1
                for evidencia in evidencias_F3:
                    #nombre_evidencia = f"{rut}_3.1_APT122_{evidencia}.docx"
                    nombre_evidencia = f"{nombre_para_archivo}_3.{n_indice}_APT122_{evidencia}.docx"
                    registro = [sede, seccion, docente[0], rut, nombre_completo, 3, nombre_evidencia, "NO"]
                    lst_evidencias.append(registro)
                    n_indice += 1

    resumen_evidencias = pd.DataFrame(lst_evidencias, columns=["sede","seccion","docente", "rut_estudiante", "estudiante", 
                                                               "fase","evidencia","estado"])
    
    # Resumen de evidencias grupales
    evidencias_grupales_F1 = ['Presentación idea de proyecto',
                           '1.4_APT122_FormativaFase1.docx',
                           '1.5_GuiaEstudiante_Fase 1_Definicion Proyecto APT (Español).docx',
                           'PLANILLA DE EVALUACIÓN FASE 1.xlsx']
    
    evidencias_grupales_F2 = ['2.4_GuiaEstudiante_Fase2_DesarrolloProyecto APT.docx',
                              '2.6_GuiaEstudiante_Fase2_Informe Final Proyecto APT.docx',
                              'PLANILLA DE EVALUACIÓN AVANCE FASE 2.xlsx',
                              'PLANILLA DE EVALUACIÓN FINAL FASE 2.xlsx'
                              ]
    
    evidencias_grupales_F3 = ['Presentación Final del proyecto (Español)',
                              'PLANILLA DE EVALUACIÓN FASE 3.xlsx']

    lst_evidencias_grupales = []
    
    for sede in resumen_equipos['sede'].unique():
        for seccion in resumen_equipos.query("sede == '{}'".format(sede))['seccion'].unique():
            s_query = f"sede == '{sede}' and seccion == '{seccion}'"
            docente = equipos.query(f"sede == '{sede}' and seccion == '{seccion}'")['docente'].unique()
            for equipo in resumen_equipos.query(s_query)[['equipo']].values.tolist(): 
                # Evidencias grupales
                for fase in list(enumerate([evidencias_grupales_F1, evidencias_grupales_F2, evidencias_grupales_F3], start=1)):                
                    for evidencia in fase[1]:                        
                        registro = [sede, seccion, docente[0], equipo[0], fase[0], evidencia, "NO"]
                        lst_evidencias_grupales.append(registro)                        

    resumen_evidencias_grupales = pd.DataFrame(lst_evidencias_grupales, columns=["sede","seccion","docente", "equipo", "fase",
                                                                                  "evidencia","estado"])
    # Cargar el libro de trabajo existente
    file_path = os.path.join(generate_folder, f"resumen_evidencias_{sede}.xlsx")

    with pd.ExcelWriter(file_path) as writer:
        resumen_evidencias.to_excel(writer, sheet_name="individuales", index=False)
        resumen_evidencias_grupales.to_excel(writer, sheet_name="grupales", index=False)

    lst_messages.append(f"Actualizado resumen evidencias en /{file_path}")
    return lst_messages

# Genera planilla para poder registrar los equipos
def generate_files(folder : str, archivo_inscritos: str, generate_folder : str): 
    logs = []
    #file_path = os.path.join(folder, archivo_inscritos)
    file_path = archivo_inscritos
    data_asignatura = pd.read_excel(file_path, header=1)
    #flash(f'Hay {data_asignatura.shape[0]} registros cargados')
    data_asignatura.columns = data_asignatura.columns.str.lower().str.replace(" ", "_", regex=True)
    #yield f"data:{json.dumps({'progress': 5})}\n\n" 
    progress, message = 5, 'Datos cargados'
    data = {'progress': progress, 'message': message}
    yield f"data:{json.dumps(data)}\n\n"
    logs.append("Datos cargados")
    # Generación de planillas
    # Resumen de equipos
    lst_estudiantes = []
    for sede in data_asignatura['sede_alumno'].unique():
        for seccion in data_asignatura.query("sede_alumno == '{}'".format(sede))['sección'].unique():        
            s_query = f"sede_alumno == '{sede}' and sección == '{seccion}'"
            docente = data_asignatura.query(f"sede_alumno == '{sede}' and sección == '{seccion}'")['docente'].unique()                
        
            for estudiante in data_asignatura.query(s_query)[['lastname', 'firstname', 'password']].values.tolist():
                registro = [sede, seccion, docente[0], estudiante[2], estudiante[0] + " " + estudiante[1], 0]
                lst_estudiantes.append(registro)

    data_equipos = pd.DataFrame(lst_estudiantes, columns=["sede","seccion","docente", "rut_estudiante", "estudiante","equipo"])
    data_equipos["rut_estudiante"] = data_equipos["rut_estudiante"].astype("str")
    # Carpeta con los archivos generados
    directory_path = generate_folder #app.config['GENERATE_FOLDER'] 
    # Verifica la existencia de la carpeta
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)
    
    #path_generate = os.path.join(directory_path, "data_equipos.xlsx")
    archivo_equipos = "data_equipos.xlsx"
    path_generate = os.path.join(directory_path, archivo_equipos)
    
    #data_equipos.to_excel("data_equipos.xlsx")
    data_equipos.to_excel(path_generate)
    #yield f"data:{json.dumps({'progress': 40})}\n\n"
    #equipos.to_excel("data_equipos_final.xlsx")
    progress, message = 75, 'Planilla de equipos OK'
    data = {'progress': progress, 'message': message}
    yield f"data:{json.dumps(data)}\n\n"
    logs.append("Planilla de equipos OK")
    progress, message = 100, 'Proceso finalizado!!!'
    data = {'progress': progress, 'message': message}
    yield f"data:{json.dumps(data)}\n\n"        
    logs.append("Proceso finalizado!!!")
    return logs

def generate_summary(upload_folder : str, generate_folder : str, listado_equipos : str):
    lst_logs = []
    # Carga los equipos    
    archivo_equipos = listado_equipos # app.config['LISTADO_EQUIPOS']
    #file_path = os.path.join(app.config['UPLOAD_FOLDER'], archivo_equipos)
    file_path = os.path.join(upload_folder, archivo_equipos)
    # Verificar si el archivo existe
    if not os.path.isfile(file_path):
        raise ValueError(f"El archivo {file_path} no es un archivo de excel o no existe.")
    
    equipos = pd.read_excel(file_path)
    equipos = equipos.drop(["Unnamed: 0"], axis = 1)
    lst_logs.append("Archivo de equipos cargado")
    # Separa a los desertores
    desertores = equipos.query("equipo == 0")
    equipos = equipos.query("equipo > 0")

    archivo_equipos = "resumen_equipos.xlsx" #app.config['ARCHIVO_EQUIPOS']
    #file_path = os.path.join(app.config['GENERATE_FOLDER'], archivo_equipos)
    file_path = os.path.join(generate_folder, archivo_equipos)
    resumen_equipos = pd.read_excel(file_path)    
    resumen_equipos = resumen_equipos.drop(["Unnamed: 0"], axis = 1)
    # Obtiene el total de equipos por sede y sección
    estado_secciones = pd.DataFrame(equipos.groupby(['sede', 'seccion'])['equipo'].count()).reset_index()
    estado_secciones.columns = ['sede', 'seccion', 'total_equipos']
    
    # Crea la lista de repositorios que deberían estar descargados
    lst_repositorios = resumen_equipos["codigo"].apply(lambda registro : registro + "-Main.zip")

    # Para efectos de prueba sólo se usan los 2 primeros
    equipo_ok, equipos_no_ok = descomprimir(lst_repositorios[:2], True)

    resumen_equipos["estado_zip"] = resumen_equipos.apply(lambda equipo : "SI" if equipo["codigo"] in equipo_ok else "NO", axis = 1)

    # Calcula porcentaje de pendientes
    for sede in np.unique(estado_secciones.sede.values):    
        for seccion in estado_secciones.query("sede == '{}'".format(sede))['seccion'].unique():        
            cumple = calculate_percentage(resumen_equipos, sede, seccion)        
            #print(f"{sede}, {seccion}: {cumple}% OK y {no_cumple}% PENDIENTE")
            estado_secciones.loc[(estado_secciones['seccion'] == seccion) &(estado_secciones['sede'] == sede), 'ok'] = cumple
        
    estado_secciones['pendiente'] = estado_secciones['ok'].apply(lambda x : 100 - x)
    
    equipos_final = equipos

    archivo_evidencias = "resumen_evidencias.xlsx" #app.config['RESUMEN_EVIDENCIAS']
    file_path = os.path.join(generate_folder, archivo_evidencias)   
    resumen_evidencias = pd.read_excel(file_path)
    #resumen_evidencias = resumen_evidencias.drop(["Unnamed: 0"], axis = 1)

    # Se deben descartar los desertores
    # Realizar un merge de los DataFrames
    desertores = desertores.drop(['equipo'], axis=1)
    sin_desertores = resumen_evidencias.merge(desertores, on=['sede', 'seccion', 'docente', 
                                                              'rut_estudiante', 'estudiante'], how='left', indicator=True)

    # Filtrar las filas que están sólo en el grupo sin desertores
    resumen_evidencias = sin_desertores[sin_desertores['_merge'] == 'left_only'].drop(columns=['_merge'])
    #print(result_df.columns)
    #random_seccion = resumen_evidencias.sample(n=1500, random_state=29)
    #resumen_evidencias.loc[random_seccion.index, 'estado'] = "OK"
    #print(equipos_final.columns)
    #print(resumen_evidencias.columns)
    #directory_path = app.config['GENERATE_FOLDER']
    #path_generate = os.path.join(generate_folder, "resumen_evidencias.xlsx")
    #resumen_evidencias.to_excel(path_generate)
    #resumen_evidencias.to_excel("resumen_evidencias.xlsx")
        
    # Carga las evidencias
    #path_generate = os.path.join(generate_folder, "resumen_evidencias.xlsx")
    #resumen_evidencias_final = pd.read_excel(path_generate)
    #resumen_evidencias_final = resumen_evidencias_final.drop(["Unnamed: 0"], axis = 1)

    progress, message = 75, 'Generando resumen evidencias...' 
    data = {'progress': progress, 'message': message}
    yield f"data:{json.dumps(data)}\n\n"
    
    data_final = pd.merge(
        resumen_evidencias,             
        equipos_final, 
        on=['sede', 'seccion', 'docente', 'rut_estudiante', 'estudiante'], 
        how='left')
    print(data_final.columns)
    # Genera reporte de evidencias por sede
    # Cuenta los OK por cada sección y sede
    reporte = data_final.groupby(["sede", "seccion", "docente"])["estado"].value_counts().unstack(fill_value=0)
    # Calcula los porcentajes
    porcentajes = reporte.div(reporte.sum(axis=1), axis=0)
    print(porcentajes.head(5))
    porcentajes = pd.DataFrame(porcentajes).reset_index().round(2)
    porcentajes['pendiente'] = porcentajes['NO'].apply(lambda x : 100 - x)
    print("**", porcentajes.columns)
    porcentajes.columns = ["Sede", "Sección", "Docente", "%_evidencias_pendientes", "%_evidencias_ok"]
    # porcentajes.to_excel("reporte_evidencias.xlsx", index=False)
    # Obtiene el reporte por estudiante
    reporte_por_estudiante = data_final[["sede", "seccion", "docente", "rut_estudiante", "estudiante", "evidencia", "estado"]]

    # Agrega metadata del reporte
    # Obtener la fecha y hora actual
    fecha_emision = datetime.now().strftime("%Y-%m-%d")
    hora_emision = datetime.now().strftime("%H:%M:%S")

    # Detalles de los archivos de fuentes de datos
    fuentes_datos = [
        {"nombre": listado_equipos, "descripcion": "Conformación de equipos"},
        {"nombre": "resumen_evidencias.xlsx", "descripcion": "Resumen de evidencias"},
    ]

    # Agregar una hoja con los metadatos
    metadatos_data = {
        "Campo": ["Fecha de emisión", "Hora de emisión"] + ["Nombre del archivo", "Descripción"] * len(fuentes_datos),
        "Valor": [fecha_emision, hora_emision] + [item for sublist in fuentes_datos for item in sublist.values()]
    }
    metadatos = pd.DataFrame(metadatos_data)    

    # Genera el informe    
    path_generate = os.path.join(generate_folder, "reporte_evidencias.xlsx")    

    with pd.ExcelWriter(path_generate) as reporte:
        metadatos.to_excel(reporte, sheet_name="Metadata", index=False)
        porcentajes.to_excel(reporte, sheet_name="por_sede", index=False)
        reporte_por_estudiante.to_excel(reporte, sheet_name="por estudiante", index=False)
        desertores.to_excel(reporte, sheet_name="desertores", index=False)
        

    porcentajes.columns = porcentajes.columns.str.lower().str.replace("%", "porc", regex=True)
    lst_logs.append("Proceso finalizado!!!")
        
    #flash("Planillas generadas exitosamente")
    progress, message = 100, 'Proceso finalizado. Fuente equipos: ' + archivo_equipos
    data = {'progress': progress, 'message': message}
    yield f"data:{json.dumps(data)}\n\n"    

    return lst_logs

