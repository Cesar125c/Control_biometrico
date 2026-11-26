def analizar(j):
    if (
        not j.registro_biometrico
        or j.duplicada
        or j.entrada_invalida
        or j.salida_invalida
    ):
        return []
    if not j.entrada_programada or not j.salida_programada:
        return []
    entrada_falta = j.entrada_real is None
    salida_falta = j.salida_real is None
    if entrada_falta and salida_falta:
        return ["Sin entrada ni salida"]
    if entrada_falta:
        return ["Sin entrada"]
    if salida_falta:
        return ["Sin salida"]
    return []
