# Control biométrico por consola

La aplicación conserva el selector de archivos de Windows para elegir un libro
Excel `.xlsx`, y solicita por consola el período del informe en formato
`MM/AAAA`. Las fechas se usan tal como están guardadas en el archivo; no se
ajustan para hacerlas coincidir con el período solicitado.

El libro de entrada debe conservar las hojas `BIOMETRICO` y `HORARIOS AE`.
El lector busca los encabezados originales:

- `BIOMETRICO`: `3. Apellidos y nombres`, `4. Cedula`, `7. Fecha`,
  `8. Marc/Ent` y `9. Marc/Sal`.
- `HORARIOS AE`: `Docente`, `Dia`, `Hora de entrada` y `Hora de salida`.

Los días esperados se obtienen de los días declarados para cada docente en
`HORARIOS AE`. Las marcaciones repetidas de una misma persona y fecha se
envían a revisión y no se usan para sumar tiempos. Las horas posteriores al
horario son solo tiempo pendiente de validación, no horas extra autorizadas.

El Excel de salida contiene pestañas separadas `Marcaciones`, `Atrasos`,
`Salidas anticipadas` y `Horas extra`, con filtros y acumulados de tiempo que
pueden superar 24 horas.

## Ejecución en Windows

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py
```
