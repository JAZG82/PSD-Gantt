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
    """Jornada de un recurso (o de la planta): horas por día, días de la semana que se trabaja
    (o un ciclo rotativo N días trabaja / M libres) y fechas sin trabajo propias.
    Tratar como inmutable (el caché de fechas libres se arma al crear)."""

    def __init__(self, horas=8.0, dias=(0, 1, 2, 3, 4), libres=(), ciclo=None):
        self.horas = float(horas)
        self.dias = tuple(sorted(set(dias)))
        self.libres = sorted(libres, key=lambda x: x[0])      # [(fecha, motivo)]
        self._s = {f for f, _ in self.libres}
        self.ciclo = tuple(ciclo) if ciclo else None           # (días trabaja, días libres, fecha de inicio del ciclo)
        if self.ciclo:
            n, m, f = self.ciclo
            if n < 1 or m < 0 or not f:
                raise ValueError("Ciclo no válido: días de trabajo ≥ 1, días libres ≥ 0 y fecha de inicio.")
        elif not self.dias:
            raise ValueError("El calendario necesita al menos un día laboral por semana.")

    def trabaja(self, d):
        """¿Este calendario trabaja el día d? (patrón de semana o ciclo, menos sus fechas libres)."""
        if d in self._s:
            return False
        if self.ciclo:
            n, m, f = self.ciclo
            return (d - f).days % (n + m) < n
        return d.weekday() in self.dias

    es_habil = trabaja

    def de(self, t=None):
        return self

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
            n += self.es_habil(d)
            d += dt.timedelta(days=1)
        return max(1, n)

    def horas_dia(self, d):
        return self.horas if self.es_habil(d) else 0.0

    def horas_rango(self, ini, fin):
        h, d = 0.0, ini
        while d <= fin:
            h += self.horas_dia(d)
            d += dt.timedelta(days=1)
        return h

    def patron(self):
        if self.ciclo:
            n, m, f = self.ciclo
            return f"ciclo {n}x{m} desde {fmt(f)}"
        return ",".join(DIAS[i] for i in self.dias)

    def resumen(self):
        return (f"{self.horas:g} h/día · {self.patron()}"
                + (f" · {len(self.libres)} día(s) sin trabajo" if self.libres else ""))


class _CalTarea:
    """Calendario efectivo de una tarea: los días no laborales globales (feriados, detenciones) aplican
    a todos; con varios recursos la tarea avanza los días en que trabaja AL MENOS uno."""

    def __init__(self, base, cals):
        self.base, self.cals = base, cals

    def es_habil(self, d):
        return d not in self.base._s and any(c.trabaja(d) for c in self.cals)

    desde = Calendario.desde
    sumar = Calendario.sumar
    contar = Calendario.contar
    horas_rango = Calendario.horas_rango

    def horas_dia(self, d):
        if d in self.base._s:
            return 0.0
        return sum(c.horas for c in self.cals if c.trabaja(d))


def nombres_recurso(texto):
    """'Mecánica/Eléctrica' → ['Mecánica', 'Eléctrica'] (separadores: / ; ,)."""
    return [x.strip() for x in re.split(r"[/;,]", texto or "") if x.strip()]


class Calendarios:
    """Calendario global (base: jornada y feriados de toda la planta) + calendarios propios por recurso.
    Un recurso sin calendario propio trabaja según el patrón de la base. Se comporta como el calendario
    base en todo lo que no dependa de la tarea."""

    def __init__(self, base=None, recursos=None):
        self.base = base or Calendario()
        self.recursos = dict(recursos or {})

    def __getattr__(self, k):                 # horas, dias, libres, _s, es_habil, desde, sumar, contar…
        if k in ("base", "recursos"):
            raise AttributeError(k)
        return getattr(self.base, k)

    def de(self, t=None):
        nombres = nombres_recurso(t.resp) if t is not None else []
        cals = [self.recursos.get(n, self.base) for n in nombres] or [self.base]
        if len(cals) == 1 and cals[0] is self.base:
            return self.base
        return _CalTarea(self.base, cals)

    def resumen(self):
        n = len(self.recursos)
        return self.base.resumen() + (f" · {n} recurso(s) con calendario propio" if n else "")


def como_calendarios(cal):
    return cal if isinstance(cal, Calendarios) else Calendarios(cal)


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
            c = cal.de(t)
            ini = c.desde(t.inicio)
            fin = c.sumar(ini, t.dur)
            r = (ini, fin, t.avance, c.horas_rango(ini, fin))
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


def sin_recurso(tareas):
    """Tareas con trabajo propio (no resumen) sin recurso asignado: usan el calendario global."""
    ids = {t.id for t in tareas}
    padres = {padre_de(t.id, ids) for t in tareas}
    return [t for t in tareas if t.id not in padres and not nombres_recurso(t.resp)]


def recursos_en_uso(tareas):
    out = []
    for t in tareas:
        for n in nombres_recurso(t.resp):
            if n not in out:
                out.append(n)
    return out


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
            c = cal.de(t)
            nuevo = c.desde(max(fines) + dt.timedelta(days=1))
            if nuevo > c.desde(t.inicio):
                t.inicio = nuevo
                movidas.add(t.id)
                cambio = True
        if not cambio:
            return len(movidas)
    raise ValueError("Dependencias circulares entre predecesoras.")


# ---------------------------------------------------------------- Excel
def guardar_excel(tareas, cal, ruta):
    cal = como_calendarios(cal)
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
    b = cal.base
    wc["A1"], wc["B1"] = "Horas por día", b.horas
    wc["A2"], wc["B2"] = "Días laborales", ",".join(DIAS[i] for i in b.dias)
    wc["A3"], wc["B3"] = "Ciclo (trabaja;libres;inicio)", (
        f"{b.ciclo[0]};{b.ciclo[1]};{fmt(b.ciclo[2])}" if b.ciclo else "")
    wc["A5"], wc["B5"] = "Días sin trabajo (toda la planta)", "Motivo"
    for c in (wc["A5"], wc["B5"]):
        c.font = Font(bold=True)
    for f, m in b.libres:
        wc.append([f, m])
        wc.cell(row=wc.max_row, column=1).number_format = "dd-mm-yy"
    wc.column_dimensions["A"].width = 34
    wc.column_dimensions["B"].width = 40
    if cal.recursos:
        wr = wb.create_sheet("Calendarios_Recursos")
        wr.append(["Recurso", "Horas por día", "Días laborales", "Ciclo trabaja", "Ciclo libres", "Inicio ciclo"])
        for c in wr[1]:
            c.font = Font(bold=True)
        for n, r in cal.recursos.items():
            wr.append([n, r.horas, ",".join(DIAS[i] for i in r.dias),
                       r.ciclo[0] if r.ciclo else None, r.ciclo[1] if r.ciclo else None,
                       r.ciclo[2] if r.ciclo else None])
            wr.cell(row=wr.max_row, column=6).number_format = "dd-mm-yy"
        wl = wb.create_sheet("Libres_Recursos")
        wl.append(["Recurso", "Fecha sin trabajo", "Motivo"])
        for c in wl[1]:
            c.font = Font(bold=True)
        for n, r in cal.recursos.items():
            for f, m in r.libres:
                wl.append([n, f, m])
                wl.cell(row=wl.max_row, column=2).number_format = "dd-mm-yy"
        for ws_ in (wr, wl):
            for letra, ancho in zip("ABCDEF", (26, 16, 24, 14, 12, 14)):
                ws_.column_dimensions[letra].width = ancho
    wb.save(ruta)


def _norm(s):
    s = str(s or "").lower()
    for a, b in (("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u")):
        s = s.replace(a, b)
    return re.sub(r"[^a-z]", "", s)


def _dias_de(texto):
    nombres = {_norm(n): i for i, n in enumerate(DIAS)}
    return [nombres[_norm(x)] for x in str(texto or "").split(",") if _norm(x) in nombres]


def _leer_calendario(wb):
    """Devuelve Calendarios (base + recursos) o None si el libro no trae hoja «Calendario»."""
    if "Calendario" not in wb.sheetnames:
        return None
    filas = list(wb["Calendario"].iter_rows(values_only=True))
    try:
        horas = float(filas[0][1])
        dias = _dias_de(filas[1][1])
        ciclo = None
        if len(filas) > 2 and filas[2][1] and str(filas[2][0]).startswith("Ciclo"):
            n, m, f = str(filas[2][1]).split(";")
            ciclo = (int(n), int(m), parse_fecha(f))
        libres = [(parse_fecha(r[0]), str(r[1] or "")) for r in filas[3:] if r and r[0] and _es_fecha_txt(r[0])]
        base = Calendario(horas, dias, libres, ciclo)
    except (IndexError, TypeError, ValueError):
        return None
    recursos = {}
    if "Calendarios_Recursos" in wb.sheetnames:
        own = {}
        if "Libres_Recursos" in wb.sheetnames:
            for r in list(wb["Libres_Recursos"].iter_rows(values_only=True))[1:]:
                if r and r[0] and r[1]:
                    own.setdefault(str(r[0]).strip(), []).append((parse_fecha(r[1]), str(r[2] or "")))
        for r in list(wb["Calendarios_Recursos"].iter_rows(values_only=True))[1:]:
            if not (r and r[0]):
                continue
            try:
                ciclo = (int(r[3]), int(r[4] or 0), parse_fecha(r[5])) if r[3] else None
                recursos[str(r[0]).strip()] = Calendario(float(r[1] or base.horas), _dias_de(r[2]) or base.dias,
                                                         own.get(str(r[0]).strip(), []), ciclo)
            except (TypeError, ValueError):
                continue
    return Calendarios(base, recursos)


def _es_fecha_txt(v):
    try:
        return parse_fecha(v) is not None
    except ValueError:
        return False


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
    base = cal or Calendarios()
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


# ---------------------------------------------------------------- Importar desde Microsoft Project (Excel)
MESES = {m: i + 1 for i, m in enumerate(["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
                                           "agosto", "septiembre", "octubre", "noviembre", "diciembre"])}


def es_project(ruta):
    """True si el Excel es una exportación de Project (hoja «Tabla_Tareas»)."""
    try:
        return "Tabla_Tareas" in load_workbook(ruta, read_only=True).sheetnames
    except Exception:
        return False


def _fecha_es(v):
    """'23 noviembre 2026 09:00 AM' → date. También acepta fechas reales de Excel."""
    if isinstance(v, (dt.date, dt.datetime)):
        return v.date() if isinstance(v, dt.datetime) else v
    m = re.match(r"\s*(\d{1,2})\s+([a-záéíóú]+)\.?\s+(\d{4})", str(v or "").lower())
    if m and m.group(2) in MESES:
        return dt.date(int(m.group(3)), MESES[m.group(2)], int(m.group(1)))
    return parse_fecha(v)


def _dur_dias(v):
    """'20 días' / '1 día' / '2 sem' → días hábiles. Unidades distintas se rechazan (no se adivinan)."""
    m = re.match(r"\s*([\d.,]+)\s*([a-záéíóú]+)", str(v or "").lower())
    if not m:
        raise ValueError(f"Duración no reconocida: «{v}»")
    n, u = float(m.group(1).replace(",", ".")), m.group(2)
    if u.startswith("d"):
        return max(1, round(n))
    if u.startswith("sem") or u == "s":
        return max(1, round(n * 5))
    raise ValueError(f"Duración en unidad no soportada: «{v}» (usa días en Project).")


def leer_project(ruta):
    """Lee la exportación a Excel de Microsoft Project (Tabla_Tareas, Tabla_asignación).
    Devuelve (tareas, calendario inferido, texto de aviso para mostrar al usuario).
    Los IDs se generan por esquema (1, 1-1, 1-1-1…); el Id original de Project queda en Notas.
    Avance = promedio de «% trabajo completado» de las asignaciones, ponderado por horas de trabajo."""
    wb = load_workbook(ruta, data_only=True)
    ws = wb["Tabla_Tareas"]
    filas = list(ws.iter_rows(values_only=True))
    cab = {_norm(c): i for i, c in enumerate(filas[0])}
    for k in ("id", "nombre", "duracion", "comienzo", "fin", "niveldeesquema"):
        if k not in cab:
            raise ValueError(f"Falta la columna «{k}» en la hoja Tabla_Tareas.")
    g = lambda r, k: r[cab[k]] if k in cab and cab[k] < len(r) else None
    datos = [r for r in filas[1:] if g(r, "id") not in (None, "")]
    # asignaciones
    recursos, avance = {}, {}
    if "Tabla_asignación" in wb.sheetnames or "Tabla_asignacion" in wb.sheetnames:
        wa = wb["Tabla_asignación" if "Tabla_asignación" in wb.sheetnames else "Tabla_asignacion"]
        fa = list(wa.iter_rows(values_only=True))
        ca = {_norm(c): i for i, c in enumerate(fa[0])}
        for r in fa[1:]:
            nom = str(r[ca.get("nombredetarea", 0)] or "").strip()
            rec = str(r[ca.get("nombredelrecurso", 1)] or "").strip()
            pct = r[ca["trabajocompletado"]] if "trabajocompletado" in ca else None
            hrs = r[ca["trabajo"]] if "trabajo" in ca else None
            try:
                pct = float(str(pct).replace(",", ".").replace("%", "")); pct = pct * 100 if pct <= 1 else pct
            except (TypeError, ValueError):
                pct = 0.0
            try:
                hrs = float(re.sub(r"[^\d.,]", "", str(hrs)).replace(",", ".") or 0)
            except ValueError:
                hrs = 0.0
            if rec:
                recursos.setdefault(nom, [])
                if rec not in recursos[nom]:
                    recursos[nom].append(rec)
            avance.setdefault(nom, []).append((hrs, pct))
    # IDs por esquema
    niveles = [int(float(g(r, "niveldeesquema"))) for r in datos]
    cont, ids, pila = [0] * 12, [], []
    for n in niveles:
        n = max(1, min(n, 11))
        cont[n] += 1
        for k in range(n + 1, 12):
            cont[k] = 0
        ids.append("-".join(str(cont[k]) for k in range(1, n + 1)))
    por_pid = {str(int(float(g(r, "id")))): ids[i] for i, r in enumerate(datos)}
    # horas por día inferidas (trabajo / duración, 100 % de unidades)
    razones, tareas, lags = [], [], 0
    for i, r in enumerate(datos):
        resumen = i + 1 < len(datos) and niveles[i + 1] > niveles[i]
        nombre = str(g(r, "nombre") or "").strip()
        dur = _dur_dias(g(r, "duracion"))
        preds = []
        for tok in re.split(r"[;,]", str(g(r, "predecesoras") or "")):
            m = re.match(r"\s*(\d+)\s*(FC|CC|FF|CF|FS|SS)?\s*(.*)$", tok.upper())
            if m and m.group(1) in por_pid:
                preds.append(por_pid[m.group(1)])
                lags += bool(m.group(2) not in (None, "FC", "FS") or m.group(3))
        a = avance.get(nombre, [])
        th = sum(h for h, _ in a)
        av = (sum(h * p for h, p in a) / th) if th else (sum(p for _, p in a) / len(a) if a else 0.0)
        if not resumen and a and th and dur:
            razones.append(round(th / dur / max(1, len(a)), 1))
        tareas.append(Tarea(ids[i], nombre, "/".join(recursos.get(nombre, [])), _fecha_es(g(r, "comienzo")), dur,
                            max(0.0, min(100.0, av)), ",".join(preds), False,
                            (f"Id Project: {int(float(g(r, 'id')))}" + (f" · {g(r, 'notas')}" if g(r, "notas") else ""))))
    # calendario: se prueba qué días de la semana reproducen las fechas de fin de Project
    hojas = [(i, resumen) for i, resumen in enumerate(
        [i + 1 < len(datos) and niveles[i + 1] > niveles[i] for i in range(len(datos))]) if not resumen]
    mejor = None
    for dias in ((0, 1, 2, 3, 4), (0, 1, 2, 3, 4, 5), tuple(range(7))):
        c = Calendario(8, dias)
        ok = sum(1 for i, _ in hojas if _fecha_es(datos[i][cab["fin"]]) == c.sumar(tareas[i].inicio, tareas[i].dur))
        if mejor is None or ok > mejor[0]:
            mejor = (ok, dias)
    h = sorted(razones)[len(razones) // 2] if razones else 8.0
    cal = Calendarios(Calendario(h, mejor[1]))
    avisos = [f"Importadas {len(tareas)} tareas desde Project ({len(hojas)} con trabajo propio, "
              f"{len(tareas) - len(hojas)} resúmenes).",
              f"Calendario inferido: {cal.resumen()} — coinciden {mejor[0]} de {len(hojas)} fechas de fin; "
              + ("sin feriados (Project no los informa en la exportación)." if mejor[0] == len(hojas)
                 else "las demás difieren, probablemente por feriados o excepciones: agrégalos en «Calendario…»."),
              "Horas por día = trabajo ÷ duración de las asignaciones." if razones
              else "Horas por día no deducibles: se usan 8 h, ajústalo en «Calendario…»."]
    if lags:
        avisos.append(f"{lags} predecesora(s) con tipo o desfase distinto de Fin→Inicio: se importaron como Fin→Inicio sin desfase.")
    return tareas, cal, "\n".join(avisos)


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


def ejemplo_calendario(hoy=None):
    """Calendario FICTICIO para el ejemplo: planta 12 h todos los días; «Mantención» en ciclo 4x3."""
    h = hoy or dt.date.today()
    h = h - dt.timedelta(days=h.weekday() + 7)
    return Calendarios(Calendario(12, range(7)), {"Mantención": Calendario(12, (), (), (4, 3, h))})


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
