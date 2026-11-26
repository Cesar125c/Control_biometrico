from pathlib import Path
import tkinter as tk
from tkinter import filedialog

from control_biometrico.infraestructura.lector_excel import cargar
from control_biometrico.infraestructura.lector_excel import PeriodoSinRegistrosError
from control_biometrico.aplicacion.servicio import generar
from control_biometrico.infraestructura.reporte import exportar


def ejecutar():
    ventana = tk.Tk()
    ventana.withdraw()
    ventana.update_idletasks()
    ventana.attributes("-topmost", True)

    try:
        ruta_seleccionada = filedialog.askopenfilename(
            parent=ventana,
            title="Selecciona el Excel de marcaciones",
            filetypes=[("Archivos de Excel", "*.xlsx")],
        )
    finally:
        ventana.destroy()

    if not ruta_seleccionada:
        print("No seleccionaste ningún archivo.")
        return

    try:
        ruta_excel = Path(ruta_seleccionada)
        periodo = input("Período del reporte (MM/AAAA): ").strip()
        ruta_salida = ruta_excel.with_name(
            f"reporte_{periodo.replace('/', '_')}.xlsx"
        )

        jornadas = cargar(ruta_excel, periodo)
        filas = generar(jornadas)
        exportar(ruta_salida, periodo, filas)

        print(f"Reporte generado: {ruta_salida.resolve()}")
        for nombre_hoja, registros in filas.items():
            print(f"{nombre_hoja}: {len(registros)} personas")

    except PeriodoSinRegistrosError as error:
        print(error)
    except (ValueError, OSError, KeyError) as error:
        print(f"Error: {error}")
