import os
from flask import Flask, render_template, request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.service_account import Credentials
import pandas as pd

# Configuración de Flask
app = Flask(__name__)
UPLOAD_FOLDER = 'static/uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Configuración de Google Drive API
SCOPES = ['https://www.googleapis.com/auth/drive']
creds = Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=creds)

# ID de la carpeta en Google Drive (debes crearla manualmente y colocar su ID aquí)
FOLDER_ID = 'TU_CARPETA_ID_EN_DRIVE'

# Ruta del archivo Excel en Google Drive
EXCEL_FILE_NAME = 'datos.xlsx'
EXCEL_FILE_ID = None

@app.route('/', methods=['GET', 'POST'])
def index():
    global EXCEL_FILE_ID

    if request.method == 'POST':
        name = request.form['name']
        process = request.form['process']
        total_score = request.form['total_score']
        achieved_score = request.form['achieved_score']
        file = request.files['file']

        # Guardar imagen en carpeta local
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)

        # Subir imagen a Google Drive
        image_metadata = {'name': file.filename, 'parents': [FOLDER_ID]}
        media = MediaFileUpload(file_path, mimetype='image/png')
        drive_service.files().create(body=image_metadata, media_body=media, fields='id').execute()

        # Descargar Excel desde Google Drive si existe
        if not EXCEL_FILE_ID:
            response = drive_service.files().list(q=f"name='{EXCEL_FILE_NAME}' and parents in '{FOLDER_ID}'",
                                                  spaces='drive',
                                                  fields='files(id, name)').execute()
            files = response.get('files', [])
            if files:
                EXCEL_FILE_ID = files[0]['id']
                request = drive_service.files().get_media(fileId=EXCEL_FILE_ID)
                with open(EXCEL_FILE_NAME, 'wb') as f:
                    f.write(request.execute())

        # Actualizar o crear Excel
        try:
            df = pd.read_excel(EXCEL_FILE_NAME)
        except FileNotFoundError:
            df = pd.DataFrame(columns=['Nombre', 'Proceso', 'Puntaje Total', 'Puntaje Logrado'])

        new_data = {
            'Nombre': name,
            'Proceso': process,
            'Puntaje Total': total_score,
            'Puntaje Logrado': achieved_score
        }
        df = df.append(new_data, ignore_index=True)
        df.to_excel(EXCEL_FILE_NAME, index=False)

        # Subir el Excel actualizado a Google Drive
        excel_metadata = {'name': EXCEL_FILE_NAME, 'parents': [FOLDER_ID]}
        media = MediaFileUpload(EXCEL_FILE_NAME, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        if EXCEL_FILE_ID:
            drive_service.files().update(fileId=EXCEL_FILE_ID, media_body=media).execute()
        else:
            drive_service.files().create(body=excel_metadata, media_body=media, fields='id').execute()

        return 'Formulario enviado con éxito. Imagen y datos almacenados en Google Drive.'

    return render_template('form.html')

if __name__ == '__main__':
    app.run(debug=True)
