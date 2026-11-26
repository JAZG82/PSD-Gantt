# PSD Gantt

Editor y visor de cronogramas para proyectos operacionales de Planta Santo Domingo (PySide6, tema oscuro, sin licencias).

## Ejecutar desde el código
```
cd ~/Downloads/psd_gantt
python3 -m venv .venv
./.venv/bin/pip install -r requirements.txt
./.venv/bin/python gantt_app.py
```
También se generan ejecutables (Mac Apple Silicon y Windows) con GitHub Actions: pestaña **Releases** del repositorio.

## Al abrir
Pregunta si quieres **comenzar desde cero**, **leer un archivo Excel** o ver un ejemplo (datos ficticios). Desde cero y al leer un Excel sin hoja «Calendario» se pide el calendario de trabajo.

## Importar desde Microsoft Project
«Abrir Excel…» reconoce la exportación de Project (Archivo → Exportar → Excel, hoja «Tabla_Tareas»): lee tareas, duración en días, comienzo, predecesoras, nivel de esquema y notas; toma los recursos y el **% de avance** (ponderado por horas de trabajo) de «Tabla_asignación». Genera IDs por esquema (1, 1-1, 1-1-1…; el Id original queda en Notas) e **infiere el calendario** (días de la semana y horas por día) comprobando que reproduce las fechas de fin de Project. Limitaciones: predecesoras con tipo distinto de Fin→Inicio o con desfase se importan como Fin→Inicio sin desfase; feriados y excepciones de calendario de Project no vienen en la exportación (agrégalos en «Calendario…»); un `.mpp` directo no se puede leer.

## Calendario de trabajo (global)
Botón **Calendario…**: horas por día, días de la semana que se trabaja y fechas sin trabajo (feriados, detenciones; se pueden agregar rangos). Atajos: Lun–Vie 8 h, Lun–Sáb 10 h, todos los días 12 h (turno 12x12). Se guarda dentro del Excel (hoja «Calendario») y los días sin trabajo se sombrean en la Gantt.

## Tareas
- **Duración en días hábiles**: el fin se calcula saltando los días sin trabajo. Si editas el **Fin**, se recalcula la duración. **Horas** = días hábiles × horas por día.
- **Jerarquía por ID**: `CP-01` resume a `CP-01-1`, `CP-01-2`… (inicio, fin, avance ponderado y horas se calculan solos).
- **Predecesoras** (coma): Fin→Inicio. «Replanificar por predecesoras» mueve tareas solo hacia adelante, al primer día hábil siguiente.
- Barra azul = tarea · amarilla = opcional · gris = resumen · borde rojo = atrasada · línea naranja = hoy.
- Fechas en `dd-mm-yy`.

## Salidas
- **Excel**: hojas «Tareas» y «Calendario».
- **PNG / PDF**: tabla + Gantt.
- **Microsoft Project**: «Copiar para Project» (al portapapeles) o «Exportar para Project…» (xlsx), con el mismo formato de tus planillas «Pegar_en_Project»: Nivel de esquema, Modo de tarea, Nombre de tarea, Duración («5 d»), Comienzo (solo tareas sin predecesoras), Fin (vacío), Predecesoras (N.º de fila), Recursos (= responsable). En Project, usa el **mismo calendario** de trabajo para que las fechas coincidan.

## Pendiente
Arrastrar barras con el mouse, hitos, calendario por tarea, carpeta compartida (SharePoint).
