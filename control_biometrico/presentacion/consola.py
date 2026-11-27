from datetime import date, time
from pathlib import Path
import re
import tkinter as tk
from tkinter import filedialog

from control_biometrico.infraestructura.lector_excel import (
    PeriodoOrigenRequeridoError,
    PeriodoSinRegistrosError,
    cargar,
)
from control_biometrico.aplicacion.servicio import generar
from control_biometrico.infraestructura.reporte import exportar


def _fecha_excepcion(texto, periodo):
    if not re.fullmatch(r"\d{2}/\d{2}/\d{4}", texto):
        raise ValueError(f"Usa DD/MM/AAAA para la fecha: {texto}")
    dia, mes, anio = map(int, texto.split("/"))
    try:
        fecha = date(anio, mes, dia)
    except ValueError as error:
        raise ValueError(f"La fecha {texto} no es válida.") from error
    if fecha.strftime("%m/%Y") != periodo:
        raise ValueError(f"La fecha {texto} debe pertenecer al período {periodo}.")
    return fecha


def _leer_fechas_excluidas(periodo):
    while True:
        texto = input(
            "Fechas excluidas (DD/MM/AAAA, separadas por comas; Enter omite): "
        ).strip()
        try:
            return {
                _fecha_excepcion(valor.strip(), periodo)
                for valor in texto.split(",")
            } if texto else set()
        except ValueError as error:
            print(f"Error: {error}")


def _leer_salidas_excepcionales(periodo):
    while True:
        texto = input(
            "Salidas excepcionales (DD/MM/AAAA HH:MM, separadas por comas; Enter omite): "
        ).strip()
        salidas = {}
        try:
            for valor in texto.split(",") if texto else ():
                partes = valor.split()
                if len(partes) != 2 or not re.fullmatch(
                    r"(?:[01]\d|2[0-3]):[0-5]\d", partes[1]
                ):
                    raise ValueError(
                        f"Usa DD/MM/AAAA HH:MM con una hora válida: {valor.strip()}"
                    )
                fecha = _fecha_excepcion(partes[0], periodo)
                hora = time(*map(int, partes[1].split(":")))
                if fecha in salidas and salidas[fecha] != hora:
                    raise ValueError(f"Ingresa una sola hora para {partes[0]}.")
                salidas[fecha] = hora
            return salidas
        except ValueError as error:
            print(f"Error: {error}")


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

        try:
            jornadas = cargar(ruta_excel, periodo)
        except PeriodoOrigenRequeridoError as error:
            print(error)
            anio_origen = input(
                "Año de origen que confirmas para reasignar (Enter cancela): "
            ).strip()
            if not anio_origen:
                print("No se reasignaron fechas; operación cancelada.")
                return
            jornadas = cargar(ruta_excel, periodo, int(anio_origen))
        fechas_excluidas = _leer_fechas_excluidas(periodo)
        salidas_excepcionales = _leer_salidas_excepcionales(periodo)
        filas = generar(jornadas, fechas_excluidas, salidas_excepcionales)
        exportar(ruta_salida, periodo, filas)

        print(f"Reporte generado: {ruta_salida.resolve()}")
        for nombre_hoja, registros in filas.items():
            print(f"{nombre_hoja}: {len(registros)} personas")
        personas_con_avisos = sum(bool(fila["casos_revision"]) for fila in filas["Marcaciones"])
        if personas_con_avisos:
            print(
                f"{personas_con_avisos} personas tienen observaciones o avisos. "
                "Consulta «Casos para revisión» en la pestaña Marcaciones."
            )

    except PeriodoSinRegistrosError as error:
        print(error)
    except (ValueError, OSError, KeyError) as error:
        print(f"Error: {error}")
