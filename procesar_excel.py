from flask import Flask, render_template, request, send_file
from openpyxl import load_workbook
import logging

app = Flask(__name__)

# Configurar el registro de errores
logging.basicConfig(filename='error.log', level=logging.ERROR)

# Ruta para la página de inicio donde se sube el archivo
@app.route('/')
def index():
    return render_template('index.html')

# Ruta para manejar la subida y procesamiento del archivo Excel
@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            return "No se ha enviado ningún archivo", 400
        
        file = request.files['file']
        if file.filename == '':
            return "Archivo no seleccionado", 400
        
        # Cargar la plantilla original
        plantilla_path = 'Plantilla.xlsx'  # Asegúrate de que esté en el mismo directorio
        wb_plantilla = load_workbook(plantilla_path)
        ws_plantilla = wb_plantilla.active
        
        # Cargar el archivo subido
        wb_subido = load_workbook(file)
        ws_subido = wb_subido.active
        
        # Iterar por las filas y transferir datos a partir de la fila 7
        for i, row in enumerate(ws_subido.iter_rows(min_row=2, values_only=True), start=7):
            if len(row) < 19:  # Asegúrate de que haya al menos 19 columnas
                logging.error(f'Fila {i} tiene menos de 19 columnas: {row}')
                continue  # Saltar a la siguiente fila si hay menos de 19 columnas
            
            ws_plantilla[f'C{i}'] = row[0]  # Columna A ("Name")
            ws_plantilla[f'D{i}'] = row[1]  # Columna B ("Surname")
            ws_plantilla[f'E{i}'] = row[14] if len(row) > 14 else ""  # Columna O ("E-mail")
            ws_plantilla[f'F{i}'] = row[18] if len(row) > 18 else ""  # Columna S ("Phone number")
            ws_plantilla[f'G{i}'] = 'D_PCC'  # Workgroup de ejemplo
            ws_plantilla[f'H{i}'] = 'Team_D_CCH_PCC_1'  # Team de ejemplo
            ws_plantilla[f'L{i}'] = row[4] if len(row) > 4 else ""  # "Is PCC"
            ws_plantilla[f'Q{i}'] = row[15] if len(row) > 15 else ""  # "CTI User"
            ws_plantilla[f'R{i}'] = row[15] if len(row) > 15 else ""  # "TTG UserID 1"
            ws_plantilla[f'V{i}'] = 'Agent'  # "Campaign Level"

        # Guardar el archivo generado
        output_file = 'Plantilla_generada.xlsx'
        wb_plantilla.save(output_file)

        # Enviar el archivo generado al cliente
        return send_file(output_file, as_attachment=True)

    except Exception as e:
        logging.error(f'Error: {e}')
        return "Se produjo un error interno", 500

if __name__ == '__main__':
    app.run(debug=True)  # Asegúrate de que el debug esté en True para desarrollo
