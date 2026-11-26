from datetime import datetime
def analizar(j):
    if j.duplicada or not j.entrada_programada or not j.salida_programada or not j.entrada_real:return 0
    m=int((datetime.combine(j.fecha,j.entrada_real)-datetime.combine(j.fecha,j.entrada_programada)).total_seconds()//60)
    return max(0, m)
