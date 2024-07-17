import os
import openpyxl
import zipfile
import time
import json
import pandas as pd
import numpy as np

from flask import Flask, request, redirect, url_for, render_template, flash, session, Response
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = 'uploads'
GENERATE_FOLDER = 'generate'
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['GENERATE_FOLDER'] = GENERATE_FOLDER
app.secret_key = 'supersecretkey'  # Necesario para mostrar mensajes flash

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
    percentage = (count / total)*100
    return round(percentage)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('Se debe indicar un archivo')
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        flash('No ha seleccionado un archivo')
        return redirect(url_for('index'))
        #return redirect(request.url)
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        session['archivo'] = file_path

        flash('Archivo con inscritos cargado satisfactoriamente')              
        return redirect(url_for('index'))
    else:
        flash('Sólo se permiten archivos xls, xlsx')
        return redirect(request.url)

@app.route('/start')
def start():
    return redirect(url_for('progress_page'))

@app.route('/progress_page')
def progress_page():
    return render_template('progress.html')

#@app.route('/process', methods=['POST'])
@app.route('/process')
def processExcel():
    # Obtiene el nombre del archivo
    archivo = session['archivo']
    
    return Response(generate_files(archivo), mimetype='text/event-stream')    

# Genera el proceso
def generate_files(archivo : str):
    data_asignatura = pd.read_excel(archivo, header=1)
    #flash(f'Hay {data_asignatura.shape[0]} registros cargados')
    data_asignatura.columns = data_asignatura.columns.str.lower().str.replace(" ", "_", regex=True)
    #yield f"data:{json.dumps({'progress': 5})}\n\n" 
    progress, message = 5, 'Datos cargados'
    data = {'progress': progress, 'message': message}
    yield f"data:{json.dumps(data)}\n\n"
    
    # Generación de planillas
    # Resumen de evidencias
    evidencias = ['AutoevaluacionCompetenciasFase1', 'DiarioReflexionFase1', 'AutoevaluacionReflexionFase1']
    lst_evidencias = []
    for sede in data_asignatura['sede_alumno'].unique():
        for seccion in data_asignatura.query("sede_alumno == '{}'".format(sede))['sección'].unique():        
            s_query = f"sede_alumno == '{sede}' and sección == '{seccion}'"
            docente = data_asignatura.query(f"sede_alumno == '{sede}' and sección == '{seccion}'")['docente'].unique()
            carpeta = f"{sede}-{seccion}-{docente[0]}"
            s_query = f"sede_alumno == '{sede}' and sección == '{seccion}'"
            for estudiante in data_asignatura.query(s_query)[['lastname', 'firstname']].values.tolist():            
                # Verifica apellido compuesto
                if len(estudiante[0].split(" ")) > 2:
                    if len(estudiante[0].split(" ")) > 3:
                        apellido = estudiante[0].split(" ")[0] + " " + estudiante[0].split(" ")[1]  + " " + estudiante[0].split(" ")[2]
                    else:
                        apellido = estudiante[0].split(" ")[0] + " " + estudiante[0].split(" ")[1]
                else:    
                    apellido = estudiante[0].split(" ")[0]
                    nombre = estudiante[1].split(" ")[0]
                    nombre_completo = estudiante[0] + " " + estudiante[1]
                    for evidencia in evidencias:
                        nombre_evidencia = f"{apellido}_{nombre}_1.1_APT122_{evidencia}.docx"
                        registro = [sede, seccion, docente[0], nombre_completo, nombre_evidencia, "NO"]
                        lst_evidencias.append(registro)
                        #print("\t",nombre_evidencia)
    resumen_evidencias = pd.DataFrame(lst_evidencias, columns=["sede","seccion","docente", "estudiante", 
                "evidencia","estado"])
    progress, message = 25, 'Nombres de Evidencias generadas'
    data = {'progress': progress, 'message': message}
    yield f"data:{json.dumps(data)}\n\n"
    
    #yield f"data:{json.dumps({'progress': 25})}\n\n" 
    # Resumen de equipos
    lst_estudiantes = []
    for sede in data_asignatura['sede_alumno'].unique():
        for seccion in data_asignatura.query("sede_alumno == '{}'".format(sede))['sección'].unique():        
            s_query = f"sede_alumno == '{sede}' and sección == '{seccion}'"
            docente = data_asignatura.query(f"sede_alumno == '{sede}' and sección == '{seccion}'")['docente'].unique()                
        
            for estudiante in data_asignatura.query(s_query)[['lastname', 'firstname']].values.tolist():
                registro = [sede, seccion, docente[0], estudiante[0] + " " + estudiante[1], 0]
                lst_estudiantes.append(registro)

    data_equipos = pd.DataFrame(lst_estudiantes, columns=["sede","seccion","docente", "estudiante","equipo"])
    # Carpeta con los archivos generados
    directory_path = app.config['GENERATE_FOLDER'] 
    # Verifica la existencia de la carpeta
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)
    
    path_generate = os.path.join(directory_path, "data_equipos.xlsx")
    #data_equipos.to_excel("data_equipos.xlsx")
    data_equipos.to_excel(path_generate)
    #yield f"data:{json.dumps({'progress': 40})}\n\n"
    # Simulación de asignación de equipos
    equipos = pd.read_excel(path_generate)
    equipos = equipos.drop(["Unnamed: 0"], axis = 1)

    # Asigna el número del equipo
    for sede in np.unique(equipos.sede.values):    
        for seccion in equipos.query("sede == '{}'".format(sede))['seccion'].unique():
            s_query = f"sede == '{sede}' and seccion == '{seccion}'"
            total_seccion = equipos.query(s_query).shape[0]
            restantes = total_seccion % 3
            total_equipos = total_seccion // 3 + (restantes > 0)
        
            for numero, estudiante in enumerate(equipos.query(s_query)[['estudiante']].values.tolist()):
                sx_query = s_query + f" and estudiante == '{estudiante[0]}'"
                indice = equipos.query(sx_query).index[0]
                numero_equipo = ((numero+1) // 3) + (1 if (numero+1) % 3 > 0 else 0)
                equipos.iloc[indice,4] = numero_equipo

    # Guarda la asignación de equipos
    path_generate = os.path.join(directory_path, "data_equipos_final.xlsx")
    equipos.to_excel(path_generate)
    #equipos.to_excel("data_equipos_final.xlsx")
    progress, message = 50, 'Planilla de equipos OK'
    data = {'progress': progress, 'message': message}
    yield f"data:{json.dumps(data)}\n\n"
    #yield f"data:{json.dumps({'progress': 50})}\n\n"    
    # Revisión de carga de evidencias
    resumen_equipos = pd.DataFrame(equipos.groupby(["sede","seccion","equipo"]).size()).reset_index()
    resumen_equipos.columns.values[-1] = 'total_integrantes'
    resumen_equipos['codigo'] = resumen_equipos.apply(lambda x : "{}-{}-{}".format(x["sede"], 
                                                            x["seccion"], 
                                                            f"EQUIPO-{x['equipo']}"), axis=1)

    # Obtiene el total de equipos por sede y sección
    estado_secciones = pd.DataFrame(resumen_equipos.groupby(['sede', 'seccion'])['equipo'].count()).reset_index()
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
    
    # Simulación de carga de ZIP
    # Selecciona equipos al azar y simula la entrega del ZIP
    random_seccion = resumen_equipos.sample(n=150, random_state=29)
    resumen_equipos.loc[random_seccion.index, 'estado_zip'] = "OK"

    # Resumen proceso
    path_generate = os.path.join(directory_path, "data_equipos_final.xlsx")
    equipos_final = pd.read_excel(path_generate)
    equipos_final = equipos_final.drop(["Unnamed: 0"], axis = 1)

    random_seccion = resumen_evidencias.sample(n=1500, random_state=29)
    resumen_evidencias.loc[random_seccion.index, 'estado'] = "OK"
    
    path_generate = os.path.join(directory_path, "resumen_evidencias.xlsx")
    resumen_evidencias.to_excel(path_generate)
    #resumen_evidencias.to_excel("resumen_evidencias.xlsx")
        
    # Carga las evidencias
    path_generate = os.path.join(directory_path, "resumen_evidencias.xlsx")
    resumen_evidencias_final = pd.read_excel(path_generate)
    resumen_evidencias_final = resumen_evidencias_final.drop(["Unnamed: 0"], axis = 1)

    progress, message = 75, 'Generando resumen evidencias...'
    data = {'progress': progress, 'message': message}
    yield f"data:{json.dumps(data)}\n\n"

    data_final = pd.merge(
        resumen_evidencias_final,             
        equipos_final, 
        on=['sede', 'seccion', 'docente', 'estudiante'], 
        how='left')

    # Genera reporte de evidencias por sede
    # Cuenta los OK por cada sección y sede
    reporte = data_final.groupby(["sede", "seccion", "docente"])["estado"].value_counts().unstack(fill_value=0)
    # Calcula los porcentajes
    porcentajes = reporte.div(reporte.sum(axis=1), axis=0)
    porcentajes = pd.DataFrame(porcentajes).reset_index().round(2)
    
    porcentajes.columns = ["Sede", "Sección", "Docente", "%_evidencias_pendientes", "%_evidencias_ok"]
    # porcentajes.to_excel("reporte_evidencias.xlsx", index=False)
    # Obtiene el reporte por estudiante
    reporte_por_estudiante = data_final[["sede", "seccion", "docente", "estudiante", "evidencia", "estado"]]

    path_generate = os.path.join(directory_path, "reporte_evidencias.xlsx")    

    with pd.ExcelWriter(path_generate) as reporte:
        porcentajes.to_excel(reporte, sheet_name="por_sede", index=False)
        reporte_por_estudiante.to_excel(reporte, sheet_name="por estudiante", index=False)

    porcentajes.columns = porcentajes.columns.str.lower().str.replace("%", "porc", regex=True)
    print("Proceso finalizado!!!")
        
    #flash("Planillas generadas exitosamente")
    progress, message = 100, 'Proceso finalizado!!!'
    data = {'progress': progress, 'message': message}
    yield f"data:{json.dumps(data)}\n\n"    
    #yield f"data:{json.dumps({'progress': 100})}\n\n"
    #return redirect(url_for('index'))
    #return render_template('process.html')
    #return Response(generate(), mimetype='text/event-stream')

if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(debug=True)
