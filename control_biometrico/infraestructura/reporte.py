from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableStyleInfo


_COLUMNAS = {
    "Marcaciones": [
        ("Cédula", "cedula"),
        ("Nombre", "nombre"),
        ("Total días con novedades", "dias_con_novedades"),
        ("Observaciones", "observaciones"),
    ],
    "Atrasos": [
        ("Cédula", "cedula"),
        ("Nombre", "nombre"),
        ("Atraso acumulado", "total_minutos"),
        ("Observaciones", "observaciones"),
    ],
    "Salidas anticipadas": [
        ("Cédula", "cedula"),
        ("Nombre", "nombre"),
        ("Tiempo acumulado", "total_minutos"),
        ("Observaciones", "observaciones"),
    ],
    "Horas extra": [
        ("Cédula", "cedula"),
        ("Nombre", "nombre"),
        ("Tiempo posterior al horario pendiente de validación", "total_minutos"),
        ("Observaciones", "observaciones"),
    ],
}

_ANCHOS = {
    "Marcaciones": (18, 36, 24, 90),
    "Atrasos": (18, 36, 24, 90),
    "Salidas anticipadas": (18, 36, 24, 90),
    "Horas extra": (18, 36, 48, 90),
}


def exportar(path, periodo, filas_por_hoja):
    libro = Workbook()
    libro.remove(libro.active)
    for numero, nombre_hoja in enumerate(_COLUMNAS):
        hoja = libro.create_sheet(nombre_hoja)
        columnas = _COLUMNAS[nombre_hoja]
        ultima_columna = len(columnas)

        hoja.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ultima_columna)
        titulo = hoja.cell(row=1, column=1, value=f"Período: {periodo}")
        titulo.font = Font(bold=True, size=13, color="26445C")
        titulo.alignment = Alignment(vertical="center")
        hoja.row_dimensions[1].height = 24

        for indice, (encabezado, _) in enumerate(columnas, 1):
            celda = hoja.cell(row=2, column=indice, value=encabezado)
            celda.font = Font(bold=True, color="FFFFFF")
            celda.fill = PatternFill("solid", fgColor="26445C")
            celda.alignment = Alignment(wrap_text=True, vertical="center")
            hoja.column_dimensions[celda.column_letter].width = _ANCHOS[nombre_hoja][indice - 1]
        hoja.row_dimensions[2].height = 34

        for fila in filas_por_hoja.get(nombre_hoja, []):
            valores = []
            for encabezado, clave in columnas:
                valor = fila[clave]
                if clave == "observaciones":
                    valor = "\n".join(valor)
                elif clave == "total_minutos":
                    valor = valor / 1440
                valores.append(valor)
            hoja.append(valores)

        for fila in hoja.iter_rows(min_row=3):
            for celda in fila:
                celda.alignment = Alignment(wrap_text=True, vertical="top")
            if nombre_hoja != "Marcaciones":
                fila[2].number_format = '[h] "h" mm "min"'

        fin = hoja.max_row
        ref = f"A2:{hoja.cell(row=2, column=ultima_columna).column_letter}{fin}"
        tabla = Table(displayName=f"Tabla{numero + 1}", ref=ref)
        tabla.autoFilter = AutoFilter(ref=ref)
        tabla.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        hoja.add_table(tabla)
        hoja.freeze_panes = "A3"
        hoja.sheet_view.showGridLines = False

    libro.properties.title = f"Control biométrico {periodo}"
    libro.properties.subject = "Marcaciones, novedades y tiempos por validar"
    libro.save(path)
