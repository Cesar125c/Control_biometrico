# Control biométrico por consola

La aplicación conserva el selector de archivos de Windows para elegir un libro
Excel `.xlsx`, y solicita por consola el período del informe en formato
`MM/AAAA`. En el archivo fuente actual, las fechas que tienen año 2030 se
interpretan como 2026 solo al solicitar un período de 2026 sin registros
directos. Si el año solicitado coincide con las fechas del libro, se procesan
sin ajuste. Otras equivalencias de año requieren confirmar el año de origen.
Los ajustes se aplican en memoria; el Excel original no se modifica.

El libro de entrada debe conservar las hojas `BIOMETRICO` y `HORARIOS AE`.
El lector busca los encabezados originales:

- `BIOMETRICO`: `3. Apellidos y nombres`, `4. Cedula`, `7. Fecha`,
  `8. Marc/Ent`, `9. Marc/Sal`, `9. Observaciones` y `10. Justificación`.
- `HORARIOS AE`: `Docente`, `Dia`, `Hora de entrada`, `Hora de salida` y
  `Feriados`. La hoja `DOCENTES` es un catálogo y no crea turnos.

Los días esperados se obtienen de los días declarados para cada docente en
`HORARIOS AE`. Las marcaciones repetidas de una misma persona y fecha se
envían a revisión y no se usan para sumar tiempos. En días con horario
habitual, el reporte marca el tiempo posterior al horario como pendiente de
validación.
En `Marcaciones`, los días con timbres faltantes se separan de los casos para
revisión; no se cuentan como faltantes los docentes sin identidad confirmada.
La pestaña incluye `Cédula`, `Nombre`, `Días con timbres faltantes`, `Detalle de
timbres faltantes` y `Casos para revisión`. Las observaciones y justificaciones
del archivo biométrico se conservan en la columna de revisión.

El Excel de salida contiene pestañas separadas `Marcaciones`, `Atrasos`,
`Salidas anticipadas` y `Horas extra`, con filtros y acumulados de tiempo que
pueden superar 24 horas.

## Excepciones del período

Antes de generar el reporte, la consola pide dos listas opcionales. Separa
varios valores con comas o pulsa Enter para omitir el campo.

- Fechas excluidas, por ejemplo `25/03/2026`. Los feriados o suspensiones de
  día completo quedan fuera de los cuatro cálculos, incluso si hubo timbres.
- Salidas excepcionales, por ejemplo `13/03/2026 13:30`. La hora indicada
  reemplaza la salida programada solo si es anterior. Si el turno termina
  a la hora de corte o antes, se usa su salida programada. Cuando aplica
  el corte, salir antes suma salida anticipada, salir después suma horas
  extra directamente y salir justo a esa hora no genera novedad de salida.
  La hora de entrada no interviene en esta comparación.

Si una fecha está excluida y también tiene una salida excepcional, prevalece
la exclusión de todo el día.

## Ejecución en Windows

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py
```
