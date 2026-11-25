"""Núcleo de datos de PSD Gantt: tareas, jerarquía por ID, cálculo de resúmenes,
replanificación por predecesoras y lectura/escritura de Excel. Sin dependencias de interfaz."""
import datetime as dt
import re
from dataclasses import dataclass, field
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

FMT_FECHA = "%d-%m-%y"          # formato del proyecto PSD: dd-mm-yy
COLUMNAS = ["ID", "Tarea", "Responsable", "Inicio", "Dur (días háb.)", "Fin", "Horas",
            "Avance %", "Predecesora", "Opcional", "Notas"]
DIAS = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]


class Calendario:
    """Calendario global: horas por día, días de la semana que se trabaja y fechas sin trabajo.
    Tratar como inmutable (el caché de fechas libres se arma al crear)."""

    def __init__(self, horas=8.0, dias=(0, 1, 2, 3, 4), libres=()):
        self.horas = float(horas)
        self.dias = tuple(sorted(set(dias)))
        self.libres = sorted(libres, key=lambda x: x[0])      # [(fecha, motivo)]
        self._s = {f for f, _ in self.libres}
        if not self.dias:
            raise ValueError("El calendario necesita al menos un día laboral por semana.")

    def es_habil(self, d):
        return d.weekday() in self.dias and d not in self._s

    def desde(self, d):
        """Primer día hábil igual o posterior a d."""
        for _ in range(3700):
            if self.es_habil(d):
                return d
            d += dt.timedelta(days=1)
        raise ValueError("No se encontró un día hábil: revisa el calendario.")

    def sumar(self, ini, n):
        """Fecha de término de una tarea que parte en `ini` y dura n días hábiles (ini cuenta como día 1)."""
        d, k = self.desde(ini), 1
        while k < max(1, n):
            d += dt.timedelta(days=1)
            if self.es_habil(d):
                k += 1
        return d

    def contar(self, ini, fin):
        """Días hábiles entre ini y fin, ambos incluidos (mínimo 1)."""
        n, d = 0, ini
        while d <= fin:
            n += d.weekday() in self.dias and d not in self._s
            d += dt.timedelta(days=1)
        return max(1, n)

    def resumen(self):
        return (f"{self.horas:g} h/día · " + ",".join(DIAS[i] for i in self.dias)
                + (f" · {len(self.libres)} día(s) sin trabajo" if self.libres else ""))


@dataclass
class Tarea:
    id: str = ""
    nombre: str = ""
    resp: str = ""
    inicio: dt.date = None
    dur: int = 1                  # días hábiles
    avance: float = 0.0
    pred: str = ""
    opcional: bool = False
    notas: str = ""


@dataclass
class Fila:
    t: Tarea
    nivel: int
    resumen: bool
    ini: dt.date
    fin: dt.date
    av: float
    horas: float = 0.0


def fmt(d):
    return d.strftime(FMT_FECHA) if d else ""


def parse_fecha(v):
    """Acepta date/datetime o texto dd-mm-yy, dd-mm-yyyy, dd/mm/yy, yyyy-mm-dd. None si vacío."""
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip().replace("/", "-").replace(".", "-")
    for f in ("%d-%m-%y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    raise ValueError(f"Fecha no válida: «{v}» (usa dd-mm-yy)")


def clave_natural(s):
    return [int(x) if x.isdigit() else x.lower() for x in re.split(r"(\d+)", s)]


def padre_de(id_, ids):
    """Ancestro más cercano existente, quitando segmentos '-' desde la derecha (CP-01-1 → CP-01)."""
    p = id_
    while "-" in p:
        p = p.rsplit("-", 1)[0]
        if p in ids:
            return p
    return None


def calcular(tareas, cal):
    """Devuelve lista de Fila (mismo orden). Las hojas toman inicio (ajustado al primer día hábil)
    y calculan el fin con su duración en días hábiles. Las tareas con hijos son resumen:
    inicio=mínimo, fin=máximo, avance ponderado por duración, horas = suma."""
    ids = {t.id for t in tareas}
    por_id = {t.id: t for t in tareas}
    hijos = {t.id: [] for t in tareas}
    for t in tareas:
        p = padre_de(t.id, ids)
        if p:
            hijos[p].append(t.id)
    memo = {}

    def datos(id_):
        if id_ in memo:
            return memo[id_]
        t = por_id[id_]
        if hijos[id_]:
            hs = [datos(h) for h in hijos[id_]]
            con = [h for h in hs if h[0] and h[1]]
            horas = sum(h[3] for h in hs)
            if con:
                ini = min(h[0] for h in con)
                fin = max(h[1] for h in con)
                pesos = [((h[1] - h[0]).days + 1, h[2]) for h in con]
                tot = sum(p for p, _ in pesos)
                r = (ini, fin, sum(p * a for p, a in pesos) / tot if tot else 0.0, horas)
            else:
                r = (None, None, 0.0, horas)
        elif t.inicio:
            ini = cal.desde(t.inicio)
            r = (ini, cal.sumar(ini, t.dur), t.avance, max(1, t.dur) * cal.horas)
        else:
            r = (None, None, t.avance, 0.0)
        memo[id_] = r
        return r

    def nivel(id_):
        n, p = 0, padre_de(id_, ids)
        while p:
            n += 1
            p = padre_de(p, ids)
        return n

    return [Fila(t, nivel(t.id), bool(hijos[t.id]), *datos(t.id)) for t in tareas]


def validar(tareas):
    """Lista de textos con problemas (IDs repetidos o vacíos, falta inicio, duración < 1, predecesoras inexistentes)."""
    pr, vistos, ids = [], set(), {t.id for t in tareas}
    padres = {padre_de(t.id, ids) for t in tareas}
    for t in tareas:
        if not t.id:
            pr.append("Hay una tarea sin ID.")
        elif t.id in vistos:
            pr.append(f"ID repetido: {t.id}")
        vistos.add(t.id)
        if t.id not in padres:
            if not t.inicio:
                pr.append(f"{t.id}: falta la fecha de inicio.")
            if t.dur < 1:
                pr.append(f"{t.id}: la duración debe ser al menos 1 día hábil.")
        for p in _preds(t):
            if p not in ids:
                pr.append(f"{t.id}: predecesora «{p}» no existe.")
    return pr


def _preds(t):
    return [x.strip() for x in re.split(r"[;,]", t.pred or "") if x.strip()]


def replanificar(tareas, cal):
    """Fin→Inicio: cada tarea (no resumen) parte el primer día hábil posterior al fin de su predecesora
    más tardía, conservando su duración. Solo mueve hacia adelante, nunca hacia atrás.
    Devuelve cantidad de tareas movidas."""
    movidas = set()
    for _ in range(len(tareas) + 1):
        filas = {f.t.id: f for f in calcular(tareas, cal)}
        cambio = False
        for t in tareas:
            if filas[t.id].resumen or not t.inicio:
                continue
            fines = [filas[p].fin for p in _preds(t) if p in filas and filas[p].fin]
            if not fines:
                continue
            nuevo = cal.desde(max(fines) + dt.timedelta(days=1))
            if nuevo > cal.desde(t.inicio):
                t.inicio = nuevo
                movidas.add(t.id)
                cambio = True
        if not cambio:
            return len(movidas)
    raise ValueError("Dependencias circulares entre predecesoras.")


# ---------------------------------------------------------------- Excel
def guardar_excel(tareas, cal, ruta):
    wb = Workbook()
    ws = wb.active
    ws.title = "Tareas"
    ws.append(COLUMNAS)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2F2F2D")
    for f in calcular(tareas, cal):
        t = f.t
        ws.append([t.id, t.nombre, t.resp, f.ini, (cal.contar(f.ini, f.fin) if f.ini else None) if f.resumen else t.dur,
                   f.fin, round(f.horas, 1), round(f.av, 1), t.pred, "Sí" if t.opcional else "", t.notas])
    for row in ws.iter_rows(min_row=2):
        row[3].number_format = row[5].number_format = "dd-mm-yy"
        row[6].number_format = "0.0"
        row[7].number_format = "0.0"
    for letra, ancho in zip("ABCDEFGHIJK", (12, 46, 22, 11, 14, 11, 9, 10, 14, 10, 40)):
        ws.column_dimensions[letra].width = ancho
    ws.freeze_panes = "C2"
    ws.cell(row=len(tareas) + 3, column=1,
            value="Fin y Horas se calculan con la hoja «Calendario». Filas resumen: valores calculados.")
    wc = wb.create_sheet("Calendario")
    wc["A1"], wc["B1"] = "Horas por día", cal.horas
    wc["A2"], wc["B2"] = "Días laborales", ",".join(DIAS[i] for i in cal.dias)
    wc["A4"], wc["B4"] = "Días sin trabajo", "Motivo"
    for c in (wc["A4"], wc["B4"]):
        c.font = Font(bold=True)
    for f, m in cal.libres:
        wc.append([f, m])
        wc.cell(row=wc.max_row, column=1).number_format = "dd-mm-yy"
    wc.column_dimensions["A"].width = 18
    wc.column_dimensions["B"].width = 40
    wb.save(ruta)


def _norm(s):
    s = str(s or "").lower()
    for a, b in (("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u")):
        s = s.replace(a, b)
    return re.sub(r"[^a-z]", "", s)


def _leer_calendario(wb):
    if "Calendario" not in wb.sheetnames:
        return None
    ws = wb["Calendario"]
    filas = list(ws.iter_rows(values_only=True))
    try:
        horas = float(filas[0][1])
        nombres = {_norm(n): i for i, n in enumerate(DIAS)}
        dias = [nombres[_norm(x)] for x in str(filas[1][1]).split(",") if _norm(x) in nombres]
        libres = [(parse_fecha(r[0]), str(r[1] or "")) for r in filas[4:] if r and r[0]]
        return Calendario(horas, dias, libres)
    except (IndexError, TypeError, ValueError):
        return None


def leer_excel(ruta):
    """Devuelve (tareas, calendario o None si el archivo no trae hoja «Calendario»)."""
    wb = load_workbook(ruta, data_only=True)
    cal = _leer_calendario(wb)
    ws = wb["Tareas"] if "Tareas" in wb.sheetnames else wb.worksheets[0]
    filas = list(ws.iter_rows(values_only=True))
    if not filas:
        return [], cal
    cab = [_norm(c) for c in filas[0]]
    idx = {}
    for i, c in enumerate(cab):
        if c.startswith("dur"):
            idx.setdefault("dur", i)
        else:
            k = {"id": "id", "tarea": "nombre", "responsable": "resp", "inicio": "inicio", "fin": "fin",
                 "avance": "avance", "predecesora": "pred", "opcional": "opcional", "notas": "notas"}.get(c)
            if k:
                idx.setdefault(k, i)
    for obligatoria, nombre in (("id", "ID"), ("nombre", "Tarea"), ("inicio", "Inicio")):
        if obligatoria not in idx:
            raise ValueError(f"Falta la columna «{nombre}» en la hoja de tareas.")
    if "dur" not in idx and "fin" not in idx:
        raise ValueError("Falta la columna «Dur (días háb.)» o «Fin».")
    g = lambda r, k: r[idx[k]] if k in idx and idx[k] < len(r) else None
    base = cal or Calendario()
    tareas = []
    for r in filas[1:]:
        id_ = str(g(r, "id") or "").strip()
        if not id_ or id_.lower().startswith(("fin y horas", "las filas")):
            continue
        try:
            av = float(g(r, "avance") or 0)
        except (TypeError, ValueError):
            av = 0.0
        ini, fin = parse_fecha(g(r, "inicio")), parse_fecha(g(r, "fin"))
        try:
            dur = int(float(g(r, "dur")))
        except (TypeError, ValueError):
            dur = base.contar(ini, fin) if ini and fin else 1
        tareas.append(Tarea(id_, str(g(r, "nombre") or ""), str(g(r, "resp") or ""), ini, max(1, dur),
                            max(0.0, min(100.0, av)), str(g(r, "pred") or ""),
                            str(g(r, "opcional") or "").strip().lower() in ("sí", "si", "x", "1", "true", "opcional"),
                            str(g(r, "notas") or "")))
    return tareas, cal


# ---------------------------------------------------------------- Microsoft Project
PROJECT_COLS = ["Nivel de esquema", "Modo de tarea", "Nombre de tarea", "Duración", "Comienzo", "Fin",
                "Predecesoras", "Nombres de los recursos"]


def filas_project(tareas, cal):
    """Mismo formato que tus planillas «Pegar_en_Project»: nivel de esquema, nombre, duración en días
    hábiles («5 d»), comienzo solo si la tarea no tiene predecesoras, predecesoras como N.º de fila
    y recursos = responsable. Fin queda vacío (lo calcula Project con su calendario)."""
    filas = calcular(tareas, cal)
    nro = {f.t.id: i + 1 for i, f in enumerate(filas)}
    out = []
    for f in filas:
        t = f.t
        preds = ",".join(str(nro[p]) for p in _preds(t) if p in nro)
        out.append([f.nivel + 1, None, t.nombre,
                    None if f.resumen else f"{t.dur} d",
                    None if (f.resumen or preds or not f.ini) else f.ini,
                    None, preds or None, t.resp or None])
    return out


def texto_project(tareas, cal):
    """Texto separado por tabuladores, sin encabezados, listo para pegar en la hoja de tareas de Project
    (con las columnas en el orden de PROJECT_COLS)."""
    def c(v):
        return "" if v is None else (fmt(v) if isinstance(v, dt.date) else str(v))
    return "\n".join("\t".join(c(v) for v in fila) for fila in filas_project(tareas, cal))


def guardar_project(tareas, cal, ruta):
    wb = Workbook()
    ws = wb.active
    ws.title = "Pegar_en_Project"
    ws.append(PROJECT_COLS)
    for c in ws[1]:
        c.font = Font(bold=True)
    for fila in filas_project(tareas, cal):
        ws.append(fila)
        ws.cell(row=ws.max_row, column=5).number_format = "dd-mm-yy"
    for letra, ancho in zip("ABCDEFGH", (16, 14, 60, 10, 12, 12, 14, 22)):
        ws.column_dimensions[letra].width = ancho
    wb.save(ruta)


def ejemplo(hoy=None):
    """Datos FICTICIOS de muestra (fechas y responsables inventados, salvo lo indicado)."""
    h = hoy or dt.date.today()
    h = h - dt.timedelta(days=h.weekday() + 7)       # lunes de la semana anterior
    d = lambda n: h + dt.timedelta(days=n)
    T = Tarea
    return [
        T("CP-01", "Habilitación Molino 01 para mineral fresco", "Gerencia Planta", notas="EJEMPLO ficticio"),
        T("CP-01-1", "Compras y preparaciones previas", "Abastecimiento", d(0), 10, 60),
        T("CP-01-2", "Montaje correa CT-10 (feeder)", "Mantención", d(14), 8, 0, "CP-01-1"),
        T("CP-01-3", "Montaje correa CT-11 al buzón de alimentación", "Mantención", d(28), 8, 0, "CP-01-2"),
        T("CP-01-4", "Pruebas y puesta en marcha", "Operaciones", d(42), 4, 0, "CP-01-3"),
        T("CP-01-5", "Chute de descarga adicional", "Mantención", d(28), 5, 0, "CP-01-2", True),
    ]
