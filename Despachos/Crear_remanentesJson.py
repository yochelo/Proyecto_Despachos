import openpyxl
import json
import os

def generar_remanentes_desde_excel(ruta_excel, salida_json):
    # Abrir la planilla
    wb = openpyxl.load_workbook(ruta_excel)
    ws = wb.active

    data = {}

    for fila in ws.iter_rows(min_row=2):  # asumimos encabezado en fila 1
        nne = str(fila[0].value).strip() if fila[0].value else None
        item = str(fila[1].value).strip() if fila[1].value else None

        if nne and item:
            data[nne] = {"item": item, "remanente": 0}

    # Guardar el JSON
    os.makedirs(os.path.dirname(salida_json), exist_ok=True)
    with open(salida_json, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

    print(f"Remanentes guardados en: {salida_json}")


# Ruta original y salida
ruta_planilla = r"C:\Users\el_xe\OneDrive\Escritorio\Biomedicos\Despachos\planilla.xlsx"
ruta_json = r"C:\Users\el_xe\OneDrive\Escritorio\Biomedicos\Despachos\remanentes\remanentes.json"

# Ejecutar
generar_remanentes_desde_excel(ruta_planilla, ruta_json)
