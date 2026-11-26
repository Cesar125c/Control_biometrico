from dataclasses import dataclass
from datetime import date, time


@dataclass
class Jornada:
    cedula: str
    nombre: str
    fecha: date
    entrada_programada: time | None
    salida_programada: time | None
    entrada_real: time | None
    salida_real: time | None
    observacion: str = ""
    entrada_invalida: bool = False
    salida_invalida: bool = False
    duplicada: bool = False
