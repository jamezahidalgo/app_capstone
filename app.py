import os
import openpyxl
import zipfile
import time
import json
import pandas as pd
import numpy as np

from flask import Flask, request, redirect, url_for, render_template, flash, session, Response
from werkzeug.utils import secure_filename

from datetime import datetime

from logic.funciones import allowed_file, descomprimir
from logic.generate import generate_equipos, generate_files, generate_summary

UPLOAD_FOLDER = 'uploads'
GENERATE_FOLDER = 'generate'
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['GENERATE_FOLDER'] = GENERATE_FOLDER

app.config['ARCHIVO_EQUIPOS'] = 'resumen_equipos.xlsx'
app.config['LISTADO_EQUIPOS'] = 'data_equipos_final.xlsx'
app.config['PLANILLA_EQUIPOS'] = 'data_equipos.xlsx'
app.config['RESUMEN_EVIDENCIAS'] = 'resumen_evidencias.xlsx'

app.secret_key = 'supersecretkey'  # Necesario para mostrar mensajes flash

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/carga')
def carga():
    return render_template('carga.html')

@app.route('/cargaEquipos')
def cargaEquipos():
    return render_template('cargarEquipos.html')

@app.route('/reporteEvidencias')
def reporteEvidencias():
    return render_template('reportes.html')


@app.route('/equipos', methods=['POST'])
def seteo_equipos():
    if 'file' not in request.files:
        flash('Se debe indicar un archivo')
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        flash('No ha seleccionado un archivo')
        return redirect(url_for('index'))
        #return redirect(request.url)
    if file and allowed_file(file.filename, ALLOWED_EXTENSIONS):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        session['archivo_equipos'] = file_path
        app.config['LISTADO_EQUIPOS'] = filename
        file.save(file_path)
        flash('Actualización de equipos cargada satisfactoriamente')              
        flash('Ubicación : ' + session['archivo_equipos'])
        mensajes = generate_equipos(app.config['UPLOAD_FOLDER'], filename, app.config['GENERATE_FOLDER'])
        flash('Planilla equipos generada satisfactoriamente')
        for mensaje in mensajes:
            flash(mensaje)
        #return redirect(url_for('index'))
        return redirect(url_for('cargaEquipos'))
    else:
        flash('Sólo se permiten archivos xls, xlsx')
        return redirect(request.url)    

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
    if file and allowed_file(file.filename, ALLOWED_EXTENSIONS):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        session['archivo'] = file_path
        file.save(file_path)
        #flash('Archivo con inscritos cargado satisfactoriamente')              
        #return redirect(url_for('index'))
        return redirect(url_for('progress_page'))
    else:
        flash('Sólo se permiten archivos xls, xlsx')
        return redirect(request.url)

@app.route('/start')
def start():
    return redirect(url_for('progress_page'))

@app.route('/progress_page')
def progress_page():
    #return render_template('page3.html')
    return render_template('progress.html')

@app.route('/process')
def processExcel():
    # Obtiene el nombre del archivo
    archivo = session['archivo']
    #archivo_equipos = app.config['PLANILLA_EQUIPOS'] 
    
    return Response(generate_files(app.config['UPLOAD_FOLDER'], archivo,
                                   app.config['GENERATE_FOLDER']), mimetype='text/event-stream')    

@app.route("/report")
def processReport():    
    return Response(generate_summary(app.config['UPLOAD_FOLDER'], app.config['GENERATE_FOLDER'],
                                     app.config['LISTADO_EQUIPOS']), mimetype='text/event-stream')    

@app.route('/reportes')
def reportes():
    return redirect(url_for('report_page'))

@app.route('/report_page')
def report_page():
    return render_template('report.html')

# Simula la asignación de equipos
@app.route('/simulacion', methods=['POST'])
def simulacion(path_generate : str):
    directory_path = app.config['GENERATE_FOLDER'] 
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
    path_generate = os.path.join(directory_path, "data_equipos_simulado.xlsx")
    equipos.to_excel(path_generate)

# Ejecucion
if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    ruta_archivo = os.path.join("config", "config.json")

    # Verificar si el archivo existe
    if os.path.exists(ruta_archivo):
        # Abrir y leer el archivo JSON
        with open(ruta_archivo, 'r') as file:
            datos = json.load(file)   

        app.config['LISTADO_EQUIPOS'] = datos["archivo_equipos"]
        app.config['UPLOAD_FOLDER'] = datos["upload_folder"]
        app.config['GENERATE_FOLDER'] = datos["generate_folder"]        
    
    app.run(debug=True)
