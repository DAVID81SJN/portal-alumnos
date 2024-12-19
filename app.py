import os
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for

app = Flask(__name__)

# Carpeta donde se almacenarán las imágenes subidas
UPLOAD_FOLDER = 'static/uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['ALLOWED_EXTENSIONS'] = {'png', 'jpg', 'jpeg', 'gif'}

# Verifica si el archivo es una imagen
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

# Verificar si la carpeta de subida existe, si no, la crea
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Ruta principal del formulario
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Recibir los datos del formulario
        name = request.form['name']
        process = request.form['process']
        total_score = request.form['total_score']
        achieved_score = request.form['achieved_score']

        # Verificar si se subió una imagen
        if 'file' not in request.files:
            return "No file part", 400
        file = request.files['file']
        if file and allowed_file(file.filename):
            filename = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filename)

            # Guardar los datos del formulario en un archivo Excel
            # Verificar si el archivo de Excel ya existe
            excel_file = 'evaluaciones.xlsx'
            if os.path.exists(excel_file):
                # Si el archivo existe, cargarlo y agregar los nuevos datos
                df = pd.read_excel(excel_file)
            else:
                # Si no existe, crear un nuevo DataFrame con las columnas necesarias
                columns = ['Nombre', 'Proceso', 'Puntaje Total', 'Puntaje Logrado', 'Imagen']
                df = pd.DataFrame(columns=columns)

            # Agregar una nueva fila con los datos del formulario
            new_data = {'Nombre': name,
                        'Proceso': process,
                        'Puntaje Total': total_score,
                        'Puntaje Logrado': achieved_score,
                        'Imagen': filename}

            # Usar pd.concat en lugar de append
            new_row = pd.DataFrame([new_data])
            df = pd.concat([df, new_row], ignore_index=True)

            # Guardar los datos en el archivo Excel
            df.to_excel(excel_file, index=False)

            return f'Formulario enviado con éxito. Imagen guardada en {filename} y datos almacenados en Excel.'

        return 'Archivo no permitido', 400
    return render_template('form.html')

if __name__ == '__main__':
    app.run(debug=True)
