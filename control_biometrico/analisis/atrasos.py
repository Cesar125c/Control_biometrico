from datetime import datetime


def analizar(jornada):
    if (
        jornada.duplicada
        or jornada.entrada_invalida
        or jornada.entrada_programada is None
        or jornada.entrada_real is None
    ):
        return 0
    diferencia = datetime.combine(jornada.fecha, jornada.entrada_real) - datetime.combine(
        jornada.fecha, jornada.entrada_programada
    )
    return max(0, int(diferencia.total_seconds() // 60))
