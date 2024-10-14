import logging

# Configurar el registro de errores
logging.basicConfig(filename='error.log', level=logging.ERROR)

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            return "No se ha enviado ningún archivo", 400
        
        file = request.files['file']
        if file.filename == '':
            return "Archivo no seleccionado", 400
        
        # Cargar la plantilla original
        plantilla_path = 'PlantillaSTEP4.xlsx'
        wb_plantilla = load_workbook(plantilla_path)
        ws_plantilla = wb_plantilla.active
        
        # Cargar el archivo subido
        wb_subido = load_workbook(file)
        ws_subido = wb_subido.active
        
        # Iterar por las filas y transferir datos a partir de la fila 7
        for i, row in enumerate(ws_subido.iter_rows(min_row=2, values_only=True), start=7):
            ws_plantilla[f'C{i}'] = row[0]  # Columna A ("Name")
            ws_plantilla[f'D{i}'] = row[1]  # Columna B ("Surname")
            ws_plantilla[f'E{i}'] = row[14]  # Columna O ("E-mail")
            ws_plantilla[f'F{i}'] = row[18]  # Columna S ("Phone number")
            ws_plantilla[f'G{i}'] = 'D_PCC'  # Workgroup de ejemplo
            ws_plantilla[f'H{i}'] = 'Team_D_CCH_PCC_1'  # Team de ejemplo
            ws_plantilla[f'L{i}'] = row[4]  # "Is PCC"
            ws_plantilla[f'Q{i}'] = row[15]  # "CTI User"
            ws_plantilla[f'R{i}'] = row[15]  # "TTG UserID 1"
            ws_plantilla[f'V{i}'] = 'Agent'  # "Campaign Level"

        # Guardar el archivo generado
        output_file = 'Plantilla_generada.xlsx'
        wb_plantilla.save(output_file)

        # Enviar el archivo generado al cliente
        return send_file(output_file, as_attachment=True)

    except Exception as e:
        logging.error(f'Error: {e}')
        return "Se produjo un error interno", 500
