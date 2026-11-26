import calendar
import math
import re
from collections import defaultdict
from datetime import date, datetime, time, timedelta

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

from control_biometrico.dominio.modelos import Jornada
from control_biometrico.dominio.nombres import clave


class PeriodoSinRegistrosError(ValueError):
    def __init__(self, periodo, periodos_disponibles):
        self.periodo = periodo
        self.periodos_disponibles = periodos_disponibles
        if periodos_disponibles:
            disponibles = ", ".join(periodos_disponibles)
            detalle = f" Períodos detectados en el archivo: {disponibles}."
        else:
            detalle = " No se detectaron fechas válidas en BIOMETRICO."
        super().__init__(
            f"No se encontraron marcaciones para el período {periodo}.{detalle}"
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
        try:
            return datetime.fromisoformat(texto).date()
        except ValueError:
            pass
        for formato in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%m/%d/%Y"):
            try:
                return datetime.strptime(texto, formato).date()
            except ValueError:
                continue
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
    return str(valor).strip()


def cargar(ruta, periodo):
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
            ],
        )
        fila_h, indices_h = _localizar_encabezados(
            horarios, ["Docente", "Dia", "Hora de entrada", "Hora de salida"]
        )
        indice = lambda indices, nombre: indices[clave(nombre)]

        horarios_por_persona = defaultdict(lambda: defaultdict(set))
        nombres_horario = {}
        cedulas_horario = {}
        indice_cedula_h = indices_h.get(clave("Cedula"), indices_h.get(clave("Cédula")))
        for fila in horarios.iter_rows(min_row=fila_h + 1, values_only=True):
            nombre = fila[indice(indices_h, "Docente")]
            dia = fila[indice(indices_h, "Dia")]
            if not nombre or not dia:
                continue
            persona = clave(nombre)
            cedula = _texto_cedula(fila[indice_cedula_h]) if indice_cedula_h is not None else ""
            identidad = (persona, cedula)
            entrada = hora(fila[indice(indices_h, "Hora de entrada")])
            salida = hora(fila[indice(indices_h, "Hora de salida")])
            horarios_por_persona[identidad][_dia_semana(dia)].add((entrada, salida))
            nombres_horario.setdefault(identidad, str(nombre).strip())
            cedulas_horario[identidad] = cedula

        personas_por_nombre = defaultdict(set)
        personas_periodo_por_nombre = defaultdict(set)
        biomarcaciones = defaultdict(list)
        periodos_detectados = set()
        biometrico_sin_horario = []
        for numero, fila in enumerate(
            biometrico.iter_rows(min_row=fila_b + 1, values_only=True),
            fila_b + 1,
        ):
            valor_fecha = fila[indice(indices_b, "7. Fecha")]
            nombre = fila[indice(indices_b, "3. Apellidos y nombres")]
            valor_cedula = fila[indice(indices_b, "4. Cedula")]
            valor_entrada = fila[indice(indices_b, "8. Marc/Ent")]
            valor_salida = fila[indice(indices_b, "9. Marc/Sal")]
            valores_relevantes = (
                valor_fecha,
                nombre,
                valor_cedula,
                valor_entrada,
                valor_salida,
            )
            if not any(
                valor is not None
                and (not isinstance(valor, str) or valor.strip())
                for valor in valores_relevantes
            ):
                continue
            cedula = _texto_cedula(valor_cedula)
            fecha_marcacion = fecha(valor_fecha)
            if fecha_marcacion is None:
                raise ValueError(
                    f"No se pudo interpretar la fecha en BIOMETRICO, fila {numero}"
                )
            periodos_detectados.add((fecha_marcacion.year, fecha_marcacion.month))
            nombre_texto = str(nombre).strip() if nombre else ""
            persona = clave(nombre_texto)
            if persona:
                personas_por_nombre[persona].add(cedula)
            if (fecha_marcacion.year, fecha_marcacion.month) != (anio, mes):
                continue
            if not persona:
                biometrico_sin_horario.append(
                    (cedula, fecha_marcacion, numero, valor_entrada, valor_salida)
                )
                continue
            personas_periodo_por_nombre[persona].add(cedula)
            biomarcaciones[(persona, cedula, fecha_marcacion)].append(
                (numero, nombre_texto, valor_entrada, valor_salida)
            )

        periodo_disponible = (anio, mes)
        if periodo_disponible not in periodos_detectados:
            disponibles = [
                f"{mes_detectado:02d}/{anio_detectado}"
                for anio_detectado, mes_detectado in sorted(periodos_detectados)
            ]
            raise PeriodoSinRegistrosError(periodo, disponibles)

        jornadas = []
        for identidad_horario, dias in horarios_por_persona.items():
            persona, cedula_horario = identidad_horario
            identidades = set()
            if cedula_horario:
                identidades.add((persona, cedula_horario))
            else:
                cedulas = personas_periodo_por_nombre.get(persona) or personas_por_nombre.get(persona) or {""}
                # Un nombre ambiguo sin cédula en HORARIOS no permite asignar jornadas con seguridad.
                if len(cedulas) > 1:
                    identidades.update((persona, c) for c in cedulas)
                else:
                    identidades.add((persona, next(iter(cedulas))))
            nombre = nombres_horario[identidad_horario]
            for dia_numero in range(1, calendar.monthrange(anio, mes)[1] + 1):
                fecha_dia = date(anio, mes, dia_numero)
                dia_nombre = _dia_semana(fecha_dia.strftime("%A"))
                if dia_nombre not in dias:
                    continue
                horarios_dia = dias[dia_nombre]
                for _, cedula in identidades:
                    clave_persona = (persona, cedula, fecha_dia)
                    registros = biomarcaciones.pop(clave_persona, [])
                    if not registros:
                        candidatos = [
                            llave for llave in biomarcaciones
                            if llave[2] == fecha_dia
                            and (
                                (cedula and llave[1] and cedula == llave[1])
                                or (persona == llave[0] and (not cedula or not llave[1]))
                            )
                        ]
                        # Solo se usa el nombre como respaldo si identifica una única fila.
                        if len(candidatos) == 1:
                            registros = biomarcaciones.pop(candidatos[0])
                    observaciones = []
                    duplicada = len(registros) > 1
                    if duplicada:
                        observaciones.append(
                            f"Hay {len(registros)} filas biométricas para la misma persona y fecha; revisar sin sumar tiempos"
                        )
                    if len(horarios_dia) != 1:
                        entrada_programada = salida_programada = None
                        observaciones.append(
                            "Horario ambiguo para este día; revisar HORARIOS AE"
                        )
                    else:
                        entrada_programada, salida_programada = next(iter(horarios_dia))
                        if not entrada_programada or not salida_programada:
                            observaciones.append(
                                "No existe un horario válido para este día; revisar HORARIOS AE"
                            )

                    entrada_real = salida_real = None
                    entrada_invalida = salida_invalida = False
                    if len(registros) == 1:
                        _, _, valor_entrada, valor_salida = registros[0]
                        entrada_real = hora(valor_entrada)
                        salida_real = hora(valor_salida)
                        entrada_invalida = valor_entrada not in (None, "") and entrada_real is None
                        salida_invalida = valor_salida not in (None, "") and salida_real is None
                        if entrada_invalida:
                            observaciones.append("Marcación de entrada inválida; revisar")
                        if salida_invalida:
                            observaciones.append("Marcación de salida inválida; revisar")
                    if len(identidades) > 1:
                        observaciones.append(
                            "El nombre normalizado aparece asociado a varias cédulas; verificar identidad"
                        )
                    jornadas.append(
                        Jornada(
                            cedula,
                            nombre,
                            fecha_dia,
                            entrada_programada,
                            salida_programada,
                            None if duplicada else entrada_real,
                            None if duplicada else salida_real,
                            "; ".join(observaciones),
                            entrada_invalida,
                            salida_invalida,
                            duplicada,
                        )
                    )

        for (persona, cedula, fecha_dia), registros in biomarcaciones.items():
            nombre = registros[0][1]
            observaciones = [
                "No existe un horario válido para este día; revisar HORARIOS AE"
            ]
            duplicada = len(registros) > 1
            entrada_real = salida_real = None
            entrada_invalida = salida_invalida = False
            if duplicada:
                observaciones.append(
                    f"Hay {len(registros)} filas biométricas para la misma persona y fecha; revisar sin sumar tiempos"
                )
            else:
                _, _, valor_entrada, valor_salida = registros[0]
                entrada_real = hora(valor_entrada)
                salida_real = hora(valor_salida)
                entrada_invalida = valor_entrada not in (None, "") and entrada_real is None
                salida_invalida = valor_salida not in (None, "") and salida_real is None
                if entrada_invalida:
                    observaciones.append("Marcación de entrada inválida; revisar")
                if salida_invalida:
                    observaciones.append("Marcación de salida inválida; revisar")
            jornadas.append(
                Jornada(
                    cedula,
                    nombre,
                    fecha_dia,
                    None,
                    None,
                    None if duplicada else entrada_real,
                    None if duplicada else salida_real,
                    "; ".join(observaciones),
                    entrada_invalida,
                    salida_invalida,
                    duplicada,
                )
            )

        sin_nombre = defaultdict(list)
        for registro in biometrico_sin_horario:
            cedula, fecha_dia, _, _, _ = registro
            sin_nombre[(cedula, fecha_dia)].append(registro)
        for (cedula, fecha_dia), registros in sin_nombre.items():
            duplicada = len(registros) > 1
            observaciones = [
                f"Fila biométrica {registro[2]} sin nombre; revisar"
                for registro in registros
            ]
            entrada_real = salida_real = None
            entrada_invalida = salida_invalida = False
            if duplicada:
                observaciones.append(
                    f"Hay {len(registros)} filas biométricas con la misma cédula y fecha; revisar sin sumar tiempos"
                )
            else:
                _, _, _, valor_entrada, valor_salida = registros[0]
                entrada_real = hora(valor_entrada)
                salida_real = hora(valor_salida)
                entrada_invalida = valor_entrada not in (None, "") and entrada_real is None
                salida_invalida = valor_salida not in (None, "") and salida_real is None
                if entrada_invalida:
                    observaciones.append("Marcación de entrada inválida; revisar")
                if salida_invalida:
                    observaciones.append("Marcación de salida inválida; revisar")
            jornadas.append(
                Jornada(
                    cedula,
                    "(Nombre no identificado)",
                    fecha_dia,
                    None,
                    None,
                    None if duplicada else entrada_real,
                    None if duplicada else salida_real,
                    "; ".join(observaciones),
                    entrada_invalida,
                    salida_invalida,
                    duplicada,
                )
            )

        return sorted(jornadas, key=lambda jornada: (jornada.nombre, jornada.fecha, jornada.cedula))
    finally:
        libro.close()
