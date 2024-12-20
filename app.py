from flask import Flask, render_template, request
import os
import pandas as pd

app = Flask(__name__)

# Configuración de directorios
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'static/uploads')
EXCEL_FILE = os.path.join(os.getcwd(), 'evaluaciones.xlsx')

# Crear carpetas y archivo Excel si no existen
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
if not os.path.exists(EXCEL_FILE):
    pd.DataFrame(columns=["Nombre", "Proceso", "Puntaje Total", "Puntaje Logrado", "Imagen"]).to_excel(EXCEL_FILE, index=False)

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/', methods=['POST'])
def submit_form():
    try:
        # Capturar datos del formulario
        name = request.form['name']
        process = request.form['process']
        total_score = request.form['total_score']
        achieved_score = request.form['achieved_score']
        file = request.files['file']

        # Guardar la imagen
        filename = f"{name.replace(' ', '_')}_{file.filename}"
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)

        # Guardar datos en el Excel
        df = pd.read_excel(EXCEL_FILE)
        new_row = {
            "Nombre": name,
            "Proceso": process,
            "Puntaje Total": total_score,
            "Puntaje Logrado": achieved_score,
            "Imagen": filepath
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)

        return "Formulario enviado con éxito. Imagen y datos almacenados correctamente."

    except Exception as e:
        return f"Error procesando el formulario: {e}"

if __name__ == '__main__':
    app.run(debug=True)
