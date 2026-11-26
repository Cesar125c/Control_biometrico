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
        identidad = (
            ("cedula", jornada.cedula)
            if jornada.cedula
            else ("nombre", clave(jornada.nombre))
        )
        if identidad not in agrupado:
            agrupado[identidad] = {
                "cedula": jornada.cedula,
                "nombre": jornada.nombre,
                "faltantes_marcacion": [],
                "casos_revision": [],
                "dias_con_timbres_faltantes": set(),
                "revision_vista": set(),
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
            fila["faltantes_marcacion"].extend(
                f"{dia}: {texto}" for texto in faltantes
            )
            fila["dias_con_timbres_faltantes"].add(jornada.fecha)
        if jornada.observacion:
            texto = (
                jornada.observacion
                if jornada.revision_sin_fecha
                else f"{dia}: {jornada.observacion}"
            )
            if texto not in fila["revision_vista"]:
                fila["casos_revision"].append(texto)
                fila["revision_vista"].add(texto)
        for etiqueta, texto_fuente in (
            ("Observaciones", jornada.observacion_biometrica),
            ("Justificación", jornada.justificacion_biometrica),
        ):
            texto_fuente = str(texto_fuente or "").strip()
            if texto_fuente:
                detalle = f"{dia}: {etiqueta} BIOMETRICO: {texto_fuente}"
                if detalle not in fila["revision_vista"]:
                    fila["casos_revision"].append(detalle)
                    fila["revision_vista"].add(detalle)

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
                "dias_con_timbres_faltantes": len(
                    fila["dias_con_timbres_faltantes"]
                ),
                "detalle_timbres_faltantes": fila["faltantes_marcacion"],
                "casos_revision": fila["casos_revision"],
            }
            for fila in filas
            if fila["faltantes_marcacion"] or fila["casos_revision"]
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
