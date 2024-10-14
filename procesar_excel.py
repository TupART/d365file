import openpyxl
import os

def procesar_datos(file, selected):
    workbook = openpyxl.load_workbook(file)
    sheet = workbook.active

    # Cargar la plantilla
    template_path = 'PlantillaSTEP4.xlsx'
    template_wb = openpyxl.load_workbook(template_path)
    template_sheet = template_wb.active

    # Empieza a copiar datos desde la fila 7
    row_start = 7

    for row_num, row in enumerate(sheet.iter_rows(min_row=3, max_row=sheet.max_row, values_only=True), start=row_start):
        name = row[0]  # Columna A: Name
        surname = row[1]  # Columna B: Surname
        
        # Si el nombre y apellido está en la lista de seleccionados, lo agregamos a la plantilla
        if f'{name} {surname}' in selected:
            template_sheet[f'C{row_num}'] = name
            template_sheet[f'D{row_num}'] = surname
            # Aquí deberías agregar las demás columnas como se especificó anteriormente

    # Guardar la plantilla procesada
    output_path = 'Plantilla_generada.xlsx'
    template_wb.save(output_path)
    
    return output_path
