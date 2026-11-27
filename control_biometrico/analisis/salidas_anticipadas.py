from datetime import datetime


def analizar(jornada, salida_excepcional=None):
    if (
        jornada.duplicada
        or jornada.salida_invalida
        or not jornada.entrada_programada
        or not jornada.salida_programada
        or not jornada.salida_real
    ):
        return 0
    salida_esperada = salida_excepcional or jornada.salida_programada
    diferencia = datetime.combine(jornada.fecha, salida_esperada) - datetime.combine(
        jornada.fecha, jornada.salida_real
    )
    return max(0, int(diferencia.total_seconds() // 60))
