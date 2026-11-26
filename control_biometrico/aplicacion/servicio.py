from control_biometrico.analisis import (
    atrasos,
    horas_extra,
    marcaciones,
    salidas_anticipadas,
)
from control_biometrico.dominio.nombres import clave


def generar(jornadas):
    agrupado = {}
    for jornada in jornadas:
        identidad = (clave(jornada.nombre), jornada.cedula)
        if identidad not in agrupado:
            agrupado[identidad] = {
                "cedula": jornada.cedula,
                "nombre": jornada.nombre,
                "dias_novedad": set(),
                "marcaciones": [],
                "observaciones": [],
                "atrasos": [],
                "atrasos_minutos": 0,
                "salidas": [],
                "salidas_minutos": 0,
                "horas_extra": [],
                "horas_extra_minutos": 0,
            }
        fila = agrupado[identidad]
        dia = jornada.fecha.strftime("%d/%m/%Y")

        faltantes = marcaciones.analizar(jornada)
        if faltantes:
            fila["marcaciones"].extend(f"{dia}: {texto}" for texto in faltantes)
            fila["dias_novedad"].add(jornada.fecha)
        if jornada.observacion:
            fila["observaciones"].append(f"{dia}: {jornada.observacion}")
            fila["dias_novedad"].add(jornada.fecha)

        minutos = atrasos.analizar(jornada)
        if minutos:
            fila["atrasos"].append(f"{dia}: {minutos} min de atraso")
            fila["atrasos_minutos"] += minutos

        minutos = salidas_anticipadas.analizar(jornada)
        if minutos:
            fila["salidas"].append(f"{dia}: {minutos} min de salida anticipada")
            fila["salidas_minutos"] += minutos

        minutos = horas_extra.analizar(jornada)
        if minutos:
            fila["horas_extra"].append(
                f"{dia}: {minutos} min posteriores al horario; pendiente de validación"
            )
            fila["horas_extra_minutos"] += minutos

    filas = list(agrupado.values())
    filas.sort(key=lambda fila: (clave(fila["nombre"]), fila["cedula"]))
    return {
        "Marcaciones": [
            {
                "cedula": fila["cedula"],
                "nombre": fila["nombre"],
                "dias_con_novedades": len(fila["dias_novedad"]),
                "observaciones": fila["marcaciones"] + fila["observaciones"],
            }
            for fila in filas
            if fila["marcaciones"] or fila["observaciones"]
        ],
        "Atrasos": [
            {
                "cedula": fila["cedula"],
                "nombre": fila["nombre"],
                "total_minutos": fila["atrasos_minutos"],
                "observaciones": fila["atrasos"],
            }
            for fila in filas
            if fila["atrasos_minutos"]
        ],
        "Salidas anticipadas": [
            {
                "cedula": fila["cedula"],
                "nombre": fila["nombre"],
                "total_minutos": fila["salidas_minutos"],
                "observaciones": fila["salidas"],
            }
            for fila in filas
            if fila["salidas_minutos"]
        ],
        "Horas extra": [
            {
                "cedula": fila["cedula"],
                "nombre": fila["nombre"],
                "total_minutos": fila["horas_extra_minutos"],
                "observaciones": fila["horas_extra"],
            }
            for fila in filas
            if fila["horas_extra_minutos"]
        ],
    }
