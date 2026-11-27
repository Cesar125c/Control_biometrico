import calendar
import math
import re
from collections import defaultdict
from datetime import date, datetime, time, timedelta

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel

from control_biometrico.dominio.modelos import Jornada
from control_biometrico.dominio.nombres import clave


EQUIVALENCIAS_ANIO_BIOMETRICO = {2030: 2026}


class PeriodoSinRegistrosError(ValueError):
    def __init__(self, periodo, periodos_disponibles):
        self.periodo = periodo
        self.periodos_disponibles = periodos_disponibles
        disponibles = ", ".join(periodos_disponibles)
        detalle = (
            f" Períodos detectados en el archivo: {disponibles}."
            if disponibles
            else " No se detectaron fechas válidas en BIOMETRICO."
        )
        super().__init__(f"No se encontraron marcaciones para {periodo}.{detalle}")


class PeriodoOrigenRequeridoError(ValueError):
    def __init__(self, periodo, periodos_disponibles, anios_origen):
        self.periodo = periodo
        self.periodos_disponibles = periodos_disponibles
        self.anios_origen = tuple(anios_origen)
        periodos = ", ".join(periodos_disponibles)
        anios = ", ".join(str(anio) for anio in anios_origen)
        super().__init__(
            f"No hay fechas para {periodo}. Períodos detectados: {periodos}. "
            f"Confirma cuál año de origen del mismo mes se debe reasignar: {anios}."
        )


def hora(valor):
    if isinstance(valor, datetime):
        return valor.time().replace(tzinfo=None)
    if isinstance(valor, time):
        return valor.replace(tzinfo=None)
    if isinstance(valor, timedelta):
        if valor.total_seconds() < 0:
            return None
        segundos = int(valor.total_seconds()) % 86400
        return (datetime.min + timedelta(seconds=segundos)).time()
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        if not math.isfinite(valor) or valor < 0:
            return None
        segundos = int((valor % 1) * 86400)
        return (datetime.min + timedelta(seconds=segundos)).time()
    if isinstance(valor, str):
        texto = valor.strip()
        if not texto:
            return None
        try:
            return datetime.fromisoformat(texto).time()
        except ValueError:
            pass
        for formato in ("%H:%M", "%H:%M:%S", "%H:%M:%S.%f"):
            try:
                return datetime.strptime(texto, formato).time()
            except ValueError:
                continue
    return None


def fecha(valor):
    """Interpreta la fecha almacenada sin reasignar su año."""
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        try:
            convertida = from_excel(valor)
            if isinstance(convertida, datetime):
                return convertida.date()
            if isinstance(convertida, date):
                return convertida
        except (OverflowError, TypeError, ValueError):
            return None
    if isinstance(valor, str):
        texto = valor.strip()
        for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(texto, formato).date()
            except ValueError:
                continue
        try:
            return datetime.fromisoformat(texto).date()
        except ValueError:
            return None
    return None


def _periodo(periodo):
    if not re.fullmatch(r"(0[1-9]|1[0-2])/\d{4}", periodo):
        raise ValueError("Período inválido, usa MM/AAAA")
    mes, anio = map(int, periodo.split("/"))
    return mes, anio


def _localizar_encabezados(hoja, encabezados):
    for numero, fila in enumerate(hoja.iter_rows(values_only=True), 1):
        indices = {
            clave(valor): indice
            for indice, valor in enumerate(fila)
            if valor is not None
        }
        if all(clave(nombre) in indices for nombre in encabezados):
            return numero, indices
    raise ValueError(f"Faltan encabezados en la hoja {hoja.title}")


def _dia_semana(valor):
    dias = {
        "MONDAY": "LUNES",
        "TUESDAY": "MARTES",
        "WEDNESDAY": "MIERCOLES",
        "THURSDAY": "JUEVES",
        "FRIDAY": "VIERNES",
        "SATURDAY": "SABADO",
        "SUNDAY": "DOMINGO",
    }
    return dias.get(clave(valor), clave(valor))


def _texto_cedula(valor):
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    if isinstance(valor, int):
        return str(valor)
    return str(valor).strip()


def _vacio(valor):
    return valor is None or (isinstance(valor, str) and not valor.strip())


def _fecha_objetivo(fecha_origen, anio_objetivo):
    try:
        return fecha_origen.replace(year=anio_objetivo)
    except ValueError as error:
        raise ValueError(
            f"La fecha {fecha_origen:%d/%m/%Y} no existe en {anio_objetivo}; "
            "no se modificó el día ni el mes."
        ) from error


def _leer_feriados(hoja, fila_encabezado, indice_feriados, mes, anio_objetivo):
    feriados = set()
    if indice_feriados is None:
        return feriados
    for (valor,) in hoja.iter_rows(
        min_row=fila_encabezado + 1,
        min_col=indice_feriados + 1,
        max_col=indice_feriados + 1,
        values_only=True,
    ):
        candidatos = []
        if isinstance(valor, (datetime, date)):
            candidatos.append(valor.date() if isinstance(valor, datetime) else valor)
        elif isinstance(valor, (int, float)) and not isinstance(valor, bool):
            if float(valor).is_integer() and 1 <= int(valor) <= 31:
                try:
                    feriados.add(date(anio_objetivo, mes, int(valor)))
                except ValueError:
                    pass
                continue
            try:
                convertido = from_excel(valor)
                if isinstance(convertido, datetime):
                    candidatos.append(convertido.date())
                elif isinstance(convertido, date):
                    candidatos.append(convertido)
            except (OverflowError, TypeError, ValueError):
                pass
        elif isinstance(valor, str) and valor.strip():
            texto = valor.strip()
            for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
                try:
                    candidatos.append(datetime.strptime(texto, formato).date())
                    break
                except ValueError:
                    continue
            if not candidatos:
                sin_anio = re.fullmatch(r"(\d{1,2})[/-](\d{1,2})", texto)
                if sin_anio:
                    primero, segundo = map(int, sin_anio.groups())
                    dia_feriado = (
                        segundo if primero == mes else primero if segundo == mes else None
                    )
                    if dia_feriado is not None:
                        try:
                            feriados.add(date(anio_objetivo, mes, dia_feriado))
                        except ValueError:
                            pass
                        continue
                trozos = re.split(r"[,;\s]+", texto)
                if all(trozo.isdigit() for trozo in trozos if trozo):
                    for trozo in trozos:
                        if trozo and 1 <= int(trozo) <= 31:
                            try:
                                feriados.add(date(anio_objetivo, mes, int(trozo)))
                            except ValueError:
                                pass
        for feriado in candidatos:
            if feriado.month == mes and feriado.year == anio_objetivo:
                feriados.add(feriado)
            elif feriado.month == mes and feriado.year not in (anio_objetivo,):
                # Solo se corrige el año biométrico; el feriado conserva su año original.
                continue
    return feriados


def _seleccionar_anio_origen(mes, anio, periodos, anio_origen_confirmado):
    if (anio, mes) in periodos:
        return anio
    anios_mismo_mes = sorted(y for y, m in periodos if m == mes)
    disponibles = [f"{m:02d}/{y}" for y, m in sorted(periodos)]
    if not anios_mismo_mes:
        raise PeriodoSinRegistrosError(f"{mes:02d}/{anio}", disponibles)

    if anio_origen_confirmado in anios_mismo_mes:
        return anio_origen_confirmado
    if anio_origen_confirmado is not None:
        raise PeriodoOrigenRequeridoError(
            f"{mes:02d}/{anio}", disponibles, anios_mismo_mes
        )
    # Equivalencia conocida para el libro de referencia: 2030 -> 2026.
    anio_equivalente = next(
        (
            origen
            for origen, destino in EQUIVALENCIAS_ANIO_BIOMETRICO.items()
            if destino == anio and origen in anios_mismo_mes
        ),
        None,
    )
    if anio_equivalente is not None:
        return anio_equivalente
    raise PeriodoOrigenRequeridoError(f"{mes:02d}/{anio}", disponibles, anios_mismo_mes)


def cargar(ruta, periodo, anio_origen_confirmado=None):
    mes, anio = _periodo(periodo)
    libro = load_workbook(ruta, read_only=True, data_only=True)
    try:
        hojas = {clave(hoja.title): hoja for hoja in libro.worksheets}
        if "BIOMETRICO" not in hojas or "HORARIOS AE" not in hojas:
            raise ValueError(
                "El archivo debe contener las hojas BIOMETRICO y HORARIOS AE"
            )
        biometrico = hojas["BIOMETRICO"]
        horarios = hojas["HORARIOS AE"]
        fila_b, indices_b = _localizar_encabezados(
            biometrico,
            [
                "3. Apellidos y nombres",
                "4. Cedula",
                "7. Fecha",
                "8. Marc/Ent",
                "9. Marc/Sal",
                "9. Observaciones",
                "10. Justificacion",
            ],
        )
        fila_h, indices_h = _localizar_encabezados(
            horarios,
            ["Docente", "Dia", "Hora de entrada", "Hora de salida", "Feriados"],
        )
        indice = lambda indices, nombre: indices[clave(nombre)]

        indice_cedula_h = next(
            (
                indices_h[clave(alias)]
                for alias in (
                    "Cedula",
                    "Cédula",
                    "Numero de cedula",
                    "Número de cédula",
                    "Nro de cedula",
                )
                if clave(alias) in indices_h
            ),
            None,
        )
        indice_feriados = indices_h[clave("Feriados")]
        feriados = _leer_feriados(horarios, fila_h, indice_feriados, mes, anio)

        horarios_por_persona = defaultdict(lambda: defaultdict(dict))
        nombres_horario = {}
        columna_salida = indice(indices_h, "Hora de salida")
        for numero, fila in enumerate(
            horarios.iter_rows(min_row=fila_h + 1, values_only=True), fila_h + 1
        ):
            nombre = fila[indice(indices_h, "Docente")]
            dia = fila[indice(indices_h, "Dia")]
            cedula = (
                _texto_cedula(fila[indice_cedula_h])
                if indice_cedula_h is not None
                else ""
            )
            persona = clave(nombre)
            dia_normalizado = _dia_semana(dia)
            if not persona and not cedula or not dia_normalizado:
                continue
            identidad = (persona, cedula)
            entrada = hora(fila[indice(indices_h, "Hora de entrada")])
            salida = hora(fila[columna_salida])
            horarios_por_persona[identidad][dia_normalizado].setdefault(
                (entrada, salida), numero
            )
            nombres_horario.setdefault(
                identidad, str(nombre).strip() if nombre else "(Nombre no identificado)"
            )

        filas_biometricas = []
        periodos_detectados = set()
        bio_identidades_por_nombre = defaultdict(set)
        identidades_biometricas = set()
        for numero, fila in enumerate(
            biometrico.iter_rows(min_row=fila_b + 1, values_only=True), fila_b + 1
        ):
            valor_fecha = fila[indice(indices_b, "7. Fecha")]
            valor_nombre = fila[indice(indices_b, "3. Apellidos y nombres")]
            valor_cedula = fila[indice(indices_b, "4. Cedula")]
            valor_entrada = fila[indice(indices_b, "8. Marc/Ent")]
            valor_salida = fila[indice(indices_b, "9. Marc/Sal")]
            valor_observacion = fila[indice(indices_b, "9. Observaciones")]
            valor_justificacion = fila[indice(indices_b, "10. Justificacion")]
            if all(
                _vacio(valor)
                for valor in (
                    valor_fecha,
                    valor_nombre,
                    valor_cedula,
                    valor_entrada,
                    valor_salida,
                    valor_observacion,
                    valor_justificacion,
                )
            ):
                continue
            fecha_origen = fecha(valor_fecha)
            if fecha_origen is None:
                raise ValueError(
                    f"No se pudo interpretar la fecha en BIOMETRICO, fila {numero}"
                )
            periodos_detectados.add((fecha_origen.year, fecha_origen.month))
            nombre_texto = str(valor_nombre).strip() if valor_nombre else ""
            nombre_clave = clave(nombre_texto)
            cedula = _texto_cedula(valor_cedula)
            identidad_bio = ("cedula", cedula) if cedula else ("nombre", nombre_clave)
            if identidad_bio[1]:
                identidades_biometricas.add(identidad_bio)
                if nombre_clave:
                    bio_identidades_por_nombre[nombre_clave].add(identidad_bio)
            filas_biometricas.append(
                {
                    "fila": numero,
                    "fecha_origen": fecha_origen,
                    "nombre": nombre_texto,
                    "nombre_clave": nombre_clave,
                    "cedula": cedula,
                    "identidad": identidad_bio,
                    "entrada": valor_entrada,
                    "salida": valor_salida,
                    "observacion": str(valor_observacion).strip()
                    if valor_observacion
                    else "",
                    "justificacion": str(valor_justificacion).strip()
                    if valor_justificacion
                    else "",
                }
            )

        anio_origen = _seleccionar_anio_origen(
            mes, anio, periodos_detectados, anio_origen_confirmado
        )
        biomarcaciones = defaultdict(list)
        identidades_periodo = set()
        for fila in filas_biometricas:
            fecha_origen = fila["fecha_origen"]
            if (fecha_origen.year, fecha_origen.month) != (anio_origen, mes):
                continue
            fecha_objetivo = _fecha_objetivo(fecha_origen, anio)
            fila["fecha"] = fecha_objetivo
            identidades_periodo.add(fila["identidad"])
            biomarcaciones[(fila["identidad"], fecha_objetivo)].append(fila)

        jornadas = []
        ancla_periodo = date(anio, mes, 1)
        for identidad_horario, dias in horarios_por_persona.items():
            nombre_clave, cedula_horario = identidad_horario
            nombre = nombres_horario[identidad_horario]
            if cedula_horario:
                candidatos = {("cedula", cedula_horario)} & identidades_biometricas
            else:
                candidatos = bio_identidades_por_nombre.get(nombre_clave, set())

            if len(candidatos) != 1:
                if not candidatos:
                    texto_revision = (
                        "No se encontró una coincidencia para este docente en BIOMETRICO"
                    )
                else:
                    texto_revision = (
                        "La identidad no es inequívoca en BIOMETRICO; confirmar docente"
                    )
                jornadas.append(
                    Jornada(
                        cedula_horario,
                        nombre,
                        ancla_periodo,
                        None,
                        None,
                        None,
                        None,
                        texto_revision,
                        registro_biometrico=False,
                        revision_sin_fecha=True,
                    )
                )
                continue

            identidad_bio = next(iter(candidatos))
            cedula_salida = (
                cedula_horario
                or (identidad_bio[1] if identidad_bio[0] == "cedula" else "")
            )
            for dia_numero in range(1, calendar.monthrange(anio, mes)[1] + 1):
                fecha_dia = date(anio, mes, dia_numero)
                if fecha_dia in feriados:
                    continue
                dia_nombre = _dia_semana(fecha_dia.strftime("%A"))
                if dia_nombre not in dias:
                    continue
                horarios_dia = dias[dia_nombre]
                if len(horarios_dia) != 1:
                    jornadas.append(
                        Jornada(
                            cedula_salida,
                            nombre,
                            fecha_dia,
                            None,
                            None,
                            None,
                            None,
                            "Horario ambiguo para este día; revisar HORARIOS AE",
                            registro_biometrico=False,
                        )
                    )
                    continue
                (entrada_programada, salida_programada), fila_horario = next(
                    iter(horarios_dia.items())
                )
                if entrada_programada is None and salida_programada is None:
                    continue
                if entrada_programada is None:
                    jornadas.append(
                        Jornada(
                            cedula_salida,
                            nombre,
                            fecha_dia,
                            entrada_programada,
                            salida_programada,
                            None,
                            None,
                            "Horario incompleto; revisar HORARIOS AE",
                            registro_biometrico=False,
                        )
                    )
                    continue

                obs = []
                if salida_programada is None:
                    celda = f"{get_column_letter(columna_salida + 1)}{fila_horario}"
                    obs.append(
                        f"Sin hora de salida programada válida en {horarios.title}!{celda}. "
                        "Revisa el horario. No se calculó salida anticipada ni horas extra."
                    )
                registros = biomarcaciones.get((identidad_bio, fecha_dia), [])
                if not registros:
                    jornadas.append(
                        Jornada(
                            cedula_salida,
                            nombre,
                            fecha_dia,
                            entrada_programada,
                            salida_programada,
                            None,
                            None,
                            "; ".join([*obs, "Sin registro en BIOMETRICO"]),
                            registro_biometrico=False,
                        )
                    )
                    continue

                notas = sorted(
                    {
                        fila["observacion"]
                        for fila in registros
                        if fila["observacion"]
                    }
                )
                justificaciones = sorted(
                    {
                        fila["justificacion"]
                        for fila in registros
                        if fila["justificacion"]
                    }
                )
                duplicada = len(registros) > 1
                entrada_real = salida_real = None
                entrada_invalida = salida_invalida = False
                if duplicada:
                    obs.append(
                        f"Hay {len(registros)} filas biométricas para esta fecha; revisar"
                    )
                else:
                    registro = registros[0]
                    entrada_real = hora(registro["entrada"])
                    salida_real = hora(registro["salida"])
                    entrada_invalida = (
                        not _vacio(registro["entrada"]) and entrada_real is None
                    )
                    salida_invalida = (
                        not _vacio(registro["salida"]) and salida_real is None
                    )
                    if entrada_invalida:
                        obs.append("Timbre de entrada inválido; revisar")
                    if salida_invalida:
                        obs.append("Timbre de salida inválido; revisar")
                    if (
                        entrada_real is not None
                        and salida_real is not None
                        and salida_real < entrada_real
                    ):
                        salida_invalida = True
                        obs.append(
                            f"Entrada {entrada_real:%H:%M:%S} y salida {salida_real:%H:%M:%S} "
                            f"en {biometrico.title}, fila {registro['fila']}. "
                            "La salida figura antes de la entrada. Confirma su fecha real. "
                            "No se calculó salida anticipada ni horas extra."
                        )
                jornadas.append(
                    Jornada(
                        cedula_salida,
                        nombre,
                        fecha_dia,
                        entrada_programada,
                        salida_programada,
                        None if duplicada else entrada_real,
                        None if duplicada else salida_real,
                        "; ".join(obs),
                        entrada_invalida,
                        salida_invalida,
                        duplicada,
                        True,
                        " | ".join(notas),
                        " | ".join(justificaciones),
                    )
                )

        return sorted(
            jornadas,
            key=lambda jornada: (
                clave(jornada.nombre),
                jornada.fecha,
                jornada.cedula,
            ),
        )
    finally:
        libro.close()
