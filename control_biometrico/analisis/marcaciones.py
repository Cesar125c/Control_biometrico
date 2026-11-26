def analizar(j):
    if j.duplicada:
        return []
    if not j.entrada_programada or not j.salida_programada: return []
    entrada_falta = not j.entrada_real and not j.entrada_invalida
    salida_falta = not j.salida_real and not j.salida_invalida
    if entrada_falta and salida_falta:
        return ["Sin marcación de entrada ni salida"]
    if entrada_falta: return ["Sin marcación de entrada"]
    if salida_falta: return ["Sin marcación de salida"]
    return []
