"""PSD Gantt — editor y visor de cronogramas (PySide6). Tema oscuro, Excel de entrada/salida,
exportación a PNG y PDF."""
import os
import sys
import datetime as dt

import json
from PySide6.QtCore import Qt, QSettings, QRectF, QPointF, QSize, QMarginsF, QDate, QTimer
from PySide6.QtGui import (QColor, QPainter, QPen, QBrush, QFont, QImage, QPdfWriter,
                           QPageSize, QPageLayout, QPalette, QFontMetrics)
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QSplitter, QTableWidget,
                               QTableWidgetItem, QScrollArea, QToolBar, QFileDialog, QMessageBox,
                               QComboBox, QLabel, QHeaderView, QToolTip, QAbstractItemView,
                               QCheckBox, QDialog, QDialogButtonBox, QDoubleSpinBox, QDateEdit, QLineEdit,
                               QPushButton, QHBoxLayout, QVBoxLayout, QGridLayout, QGroupBox, QListWidget, QSpinBox)
import gantt_core as core

BLUE, YELLOW, GREEN, RED, ORANGE = "#3a8ae6", "#f2b632", "#2fbf71", "#ef5350", "#e8622c"
INK, INK2, GRID, SURFACE, MUTED = "#f2f2f0", "#c9c8c3", "#3d3d3a", "#1a1a19", "#8a8984"
BANDA = "#202020"
RH, HH = 28, 44                       # alto de fila y de cabecera (px)
ZOOM = {"Día": 22, "Semana": 9, "Mes": 3}
NOLAB, FERIADO, NOLAB_REC = "#262624", "#35302a", "#2c2a38"   # sin trabajo / feriado / día libre del recurso


def _c(x):
    return QColor(x)


def pintar(p, filas, d0, ndias, ppd, x0=0, y0=0, hoy=None, cal=None):
    """Dibuja cabecera, cuadrícula, barras y dependencias. Reutilizado por pantalla y exportación."""
    w = ndias * ppd
    alto = HH + len(filas) * RH
    p.fillRect(QRectF(x0, y0, w, alto), _c(SURFACE))
    f = p.font()
    fb = QFont(f)
    fb.setBold(True)
    # bandas alternas + días sin trabajo + cabecera
    for i in range(len(filas)):
        if i % 2:
            p.fillRect(QRectF(x0, y0 + HH + i * RH, w, RH), _c(BANDA))
    if cal:
        for k in range(ndias):
            d = d0 + dt.timedelta(days=k)
            if not cal.es_habil(d):
                p.fillRect(QRectF(x0 + k * ppd, y0 + HH, ppd, alto - HH),
                           _c(FERIADO if d in cal._s else NOLAB))
        memo = {}                                  # sombreado propio de cada tarea (calendario de sus recursos)
        for i, fl in enumerate(filas):
            if fl.resumen:
                continue
            ct = cal.de(fl.t)
            if ct is cal.base:
                continue
            clave = tuple(core.nombres_recurso(fl.t.resp))
            if clave not in memo:
                tr, k = [], 0
                while k < ndias:
                    d = d0 + dt.timedelta(days=k)
                    if cal.base.es_habil(d) and not ct.es_habil(d):
                        k1 = k
                        while k1 + 1 < ndias and cal.base.es_habil(d0 + dt.timedelta(days=k1 + 1)) \
                                and not ct.es_habil(d0 + dt.timedelta(days=k1 + 1)):
                            k1 += 1
                        tr.append((k, k1))
                        k = k1
                    k += 1
                memo[clave] = tr
            for k, k1 in memo[clave]:
                p.fillRect(QRectF(x0 + k * ppd, y0 + HH + i * RH, (k1 - k + 1) * ppd, RH), _c(NOLAB_REC))
    p.setPen(_c(INK2))
    ult_mes = None
    for k in range(ndias):
        d = d0 + dt.timedelta(days=k)
        x = x0 + k * ppd
        if d.day == 1 or k == 0:
            p.setPen(QPen(_c(GRID)))
            p.drawLine(QPointF(x, y0), QPointF(x, y0 + alto))
            p.setPen(_c(INK))
            p.setFont(fb)
            p.drawText(QRectF(x + 3, y0, 120, HH / 2), Qt.AlignVCenter | Qt.AlignLeft, d.strftime("%m-%y"))
            p.setFont(f)
        if ppd >= 14:
            p.setPen(_c(MUTED if d.weekday() >= 5 else INK2))
            p.drawText(QRectF(x, y0 + HH / 2, ppd, HH / 2), Qt.AlignCenter, str(d.day))
            if d.weekday() == 0:
                p.setPen(QPen(_c(GRID), 1, Qt.DotLine))
                p.drawLine(QPointF(x, y0 + HH / 2), QPointF(x, y0 + alto))
        elif d.weekday() == 0 and ppd >= 5:
            p.setPen(_c(MUTED))
            p.drawText(QRectF(x, y0 + HH / 2, 30, HH / 2), Qt.AlignVCenter | Qt.AlignLeft, str(d.day))
    p.setPen(QPen(_c(GRID)))
    p.drawLine(QPointF(x0, y0 + HH), QPointF(x0 + w, y0 + HH))

    pos = {}
    for i, fl in enumerate(filas):
        y = y0 + HH + i * RH
        if not (fl.ini and fl.fin):
            continue
        xa = x0 + (fl.ini - d0).days * ppd
        xb = x0 + ((fl.fin - d0).days + 1) * ppd
        pos[fl.t.id] = (xa, xb, y + RH / 2)
        r = QRectF(xa, y + 6, max(xb - xa, 3), RH - 12)
        atras = hoy and fl.fin < hoy and fl.av < 100
        if fl.resumen:
            p.setPen(Qt.NoPen)
            p.setBrush(_c(INK2))
            p.drawRect(QRectF(xa, y + 8, max(xb - xa, 3), 7))
            p.drawPolygon([QPointF(xa, y + 8), QPointF(xa + 6, y + 8), QPointF(xa, y + 18)])
            p.drawPolygon([QPointF(xb, y + 8), QPointF(xb - 6, y + 8), QPointF(xb, y + 18)])
            p.setBrush(_c(GREEN))
            p.drawRect(QRectF(xa, y + 8, max(xb - xa, 3) * fl.av / 100, 7))
        else:
            base = YELLOW if fl.t.opcional else BLUE
            col = _c(base)
            p.setPen(QPen(_c(RED), 2) if atras else Qt.NoPen)
            fondo = QColor(col)
            fondo.setAlpha(110)
            p.setBrush(fondo)
            p.drawRoundedRect(r, 3, 3)
            p.setPen(Qt.NoPen)
            p.setBrush(col)
            if fl.av > 0:
                p.drawRoundedRect(QRectF(r.x(), r.y(), r.width() * fl.av / 100, r.height()), 3, 3)
        if ppd >= 9 and (xb - xa) > 40:
            p.setPen(_c("#ffffff"))
            if not fl.resumen:
                p.drawText(r.adjusted(9, 0, -2, 0), Qt.AlignVCenter | Qt.AlignLeft, f"{fl.av:.0f}%")
    # dependencias fin→inicio
    p.setPen(QPen(_c("#a8a7a2"), 1))
    p.setBrush(Qt.NoBrush)
    for fl in filas:
        if fl.t.id not in pos:
            continue
        for pr in core._preds(fl.t):
            if pr in pos:
                xa, xb, ya = pos[pr]
                xs, _, ys = pos[fl.t.id]
                xm = xb + 5
                p.drawLine(QPointF(xb, ya), QPointF(xm, ya))
                p.drawLine(QPointF(xm, ya), QPointF(xm, ys))
                p.drawLine(QPointF(xm, ys), QPointF(xs, ys))
                p.drawLine(QPointF(xs, ys), QPointF(xs - 4, ys - 3))
                p.drawLine(QPointF(xs, ys), QPointF(xs - 4, ys + 3))
    if hoy and d0 <= hoy < d0 + dt.timedelta(days=ndias):
        x = x0 + (hoy - d0).days * ppd + ppd / 2
        p.setPen(QPen(_c(ORANGE), 2))
        p.drawLine(QPointF(x, y0 + HH / 2), QPointF(x, y0 + alto))
    return w, alto


class GanttWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.filas, self.d0, self.ndias, self.ppd, self.cal = [], dt.date.today(), 30, 22, None
        self.setMouseTracking(True)

    def configurar(self, filas, ppd, cal):
        self.filas, self.ppd, self.cal = filas, ppd, cal
        fs = [d for f in filas for d in (f.ini, f.fin) if d]
        hoy = dt.date.today()
        if fs:
            self.d0 = min(min(fs), hoy) - dt.timedelta(days=min(fs).weekday() + 7)
            fin = max(max(fs), hoy) + dt.timedelta(days=14)
        else:
            self.d0, fin = hoy - dt.timedelta(days=7), hoy + dt.timedelta(days=60)
        self.ndias = (fin - self.d0).days
        self.setFixedSize(self.ndias * ppd, HH + len(filas) * RH)
        self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        pintar(p, self.filas, self.d0, self.ndias, self.ppd, hoy=dt.date.today(), cal=self.cal)

    def mouseMoveEvent(self, e):
        i = int((e.position().y() - HH) // RH)
        if 0 <= i < len(self.filas):
            f = self.filas[i]
            if f.ini and f.fin:
                t = f.t
                QToolTip.showText(e.globalPosition().toPoint(),
                                  f"<b>{t.id}</b> {t.nombre}<br>{core.fmt(f.ini)} → {core.fmt(f.fin)} "
                                  f"({self.cal.de(t).contar(f.ini, f.fin)} días háb. · {f.horas:g} h)<br>Avance {f.av:.0f}%"
                                  + (f"<br>Resp.: {t.resp}" if t.resp else "")
                                  + (f"<br>Pred.: {t.pred}" if t.pred else ""), self)
                return
        QToolTip.hideText()


COLS = ["ID", "Tarea", "Responsable", "Inicio", "Dur (d)", "Fin", "Horas", "Avance %", "Predecesora", "Opcional"]
ANCHOS = [80, 270, 110, 80, 60, 80, 60, 70, 90, 60]
SOLO_LECTURA_RESUMEN = (3, 4, 5, 6, 7)
PRESETS = {"Lun–Vie, 8 h": (8, (0, 1, 2, 3, 4), None),
           "Lun–Sáb, 10 h": (10, (0, 1, 2, 3, 4, 5), None),
           "Todos los días, 12 h": (12, tuple(range(7)), None),
           "Ciclo 4x3, 12 h": (12, (), (4, 3)),
           "Ciclo 7x7, 12 h": (12, (), (7, 7)),
           "Ciclo 14x14, 12 h": (12, (), (14, 14))}


class EditorCalendario(QWidget):
    """Edita una jornada: horas por día, días de la semana o ciclo rotativo (N trabaja / M libres)
    y fechas sin trabajo. Se usa para el calendario global y para cada recurso."""
    def __init__(self, cal, parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        gl = QGridLayout()
        self.horas = QDoubleSpinBox()
        self.horas.setRange(0.5, 24)
        self.horas.setSingleStep(0.5)
        self.horas.setSuffix(" h por día trabajado")
        gl.addWidget(QLabel("Horas de trabajo:"), 0, 0)
        gl.addWidget(self.horas, 0, 1, 1, 7)
        self.modo_sem = QCheckBox("Días de la semana:")
        self.modo_sem.setChecked(True)
        self.modo_sem.toggled.connect(self._modo)
        gl.addWidget(self.modo_sem, 1, 0)
        self.chk = []
        for i, n in enumerate(core.DIAS):
            c = QCheckBox(n)
            self.chk.append(c)
            gl.addWidget(c, 1, i + 1)
        self.modo_cic = QCheckBox("Ciclo rotativo:")
        self.modo_cic.toggled.connect(lambda on: self.modo_sem.setChecked(not on))
        gl.addWidget(self.modo_cic, 2, 0)
        self.n = QSpinBox()
        self.n.setRange(1, 60)
        self.n.setSuffix(" días trabaja")
        self.m = QSpinBox()
        self.m.setRange(0, 60)
        self.m.setSuffix(" días libres")
        self.f0 = QDateEdit(QDate.currentDate())
        self.f0.setDisplayFormat("dd-MM-yy")
        self.f0.setCalendarPopup(True)
        gl.addWidget(self.n, 2, 1, 1, 2)
        gl.addWidget(self.m, 2, 3, 1, 2)
        gl.addWidget(QLabel("desde"), 2, 5)
        gl.addWidget(self.f0, 2, 6, 1, 2)
        lay.addLayout(gl)
        fila = QHBoxLayout()
        for nom, (h, ds, cic) in PRESETS.items():
            b = QPushButton(nom)
            b.clicked.connect(lambda _=False, h=h, ds=ds, cic=cic: self._poner(h, ds, cic, None))
            fila.addWidget(b)
        lay.addLayout(fila)
        lay.addWidget(QLabel("Días sin trabajo propios (vacaciones, licencias, detenciones):"))
        self.lista = QTableWidget(0, 2)
        self.lista.setHorizontalHeaderLabels(["Fecha", "Motivo"])
        self.lista.horizontalHeader().setStretchLastSection(True)
        self.lista.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.lista.setEditTriggers(QAbstractItemView.NoEditTriggers)
        lay.addWidget(self.lista)
        add = QHBoxLayout()
        self.fecha = QDateEdit(QDate.currentDate())
        self.fecha.setDisplayFormat("dd-MM-yy")
        self.fecha.setCalendarPopup(True)
        self.fin = QDateEdit(QDate.currentDate())
        self.fin.setDisplayFormat("dd-MM-yy")
        self.fin.setCalendarPopup(True)
        self.motivo = QLineEdit()
        self.motivo.setPlaceholderText("Motivo (opcional)")
        b1, b2 = QPushButton("Agregar"), QPushButton("Quitar seleccionados")
        b1.clicked.connect(self._agregar)
        b2.clicked.connect(self._quitar)
        for w_ in (QLabel("Desde"), self.fecha, QLabel("hasta"), self.fin, self.motivo, b1, b2):
            add.addWidget(w_)
        lay.addLayout(add)
        self.cargar(cal)

    def _modo(self, on):
        self.modo_cic.setChecked(not on)
        for c in self.chk:
            c.setEnabled(on)
        for w_ in (self.n, self.m, self.f0):
            w_.setEnabled(not on)

    def _poner(self, h, ds, cic, f0):
        self.horas.setValue(h)
        self.modo_sem.setChecked(not cic)
        for i, c in enumerate(self.chk):
            c.setChecked(i in ds)
        if cic:
            self.n.setValue(cic[0])
            self.m.setValue(cic[1])
            if f0:
                self.f0.setDate(QDate(f0.year, f0.month, f0.day))
        self._modo(not cic)

    def cargar(self, cal):
        self.lista.setRowCount(0)
        c = cal.ciclo
        self._poner(cal.horas, cal.dias, (c[0], c[1]) if c else None, c[2] if c else None)
        for f, m in cal.libres:
            self._fila(f, m)

    def _fila(self, f, m):
        r = self.lista.rowCount()
        self.lista.insertRow(r)
        self.lista.setItem(r, 0, QTableWidgetItem(core.fmt(f)))
        self.lista.setItem(r, 1, QTableWidgetItem(m))
        self.lista.item(r, 0).setData(Qt.UserRole, f.toordinal())

    def _agregar(self):
        a, b = self.fecha.date().toPython(), self.fin.date().toPython()
        if b < a:
            a, b = b, a
        ya = {self.lista.item(r, 0).data(Qt.UserRole) for r in range(self.lista.rowCount())}
        for k in range((b - a).days + 1):
            d = a + dt.timedelta(days=k)
            if d.toordinal() not in ya:
                self._fila(d, self.motivo.text().strip())
        self.lista.sortItems(0)

    def _quitar(self):
        for r in sorted({i.row() for i in self.lista.selectedIndexes()}, reverse=True):
            self.lista.removeRow(r)

    def resultado(self):
        """Calendario editado, o None (con aviso) si es inválido."""
        dias = [i for i, c in enumerate(self.chk) if c.isChecked()]
        ciclo = None
        if self.modo_cic.isChecked():
            ciclo = (self.n.value(), self.m.value(), self.f0.date().toPython())
        elif not dias:
            QMessageBox.warning(self, "Calendario", "Marca al menos un día de trabajo a la semana, o usa un ciclo.")
            return None
        libres = [(dt.date.fromordinal(self.lista.item(r, 0).data(Qt.UserRole)), self.lista.item(r, 1).text())
                  for r in range(self.lista.rowCount())]
        return core.Calendario(self.horas.value(), dias, libres, ciclo)


class DialogoCalendario(QDialog):
    """Calendario global de la planta (jornada y feriados/detenciones que afectan a todos)."""
    def __init__(self, cal, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Calendario global de la planta")
        self.setMinimumWidth(640)
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("Jornada base y días sin trabajo de TODA la planta. Los recursos sin calendario propio "
                             "usan esta jornada."))
        self.ed = EditorCalendario(cal.base, self)
        lay.addWidget(self.ed)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        self.cal0, self.resultado = cal, None

    def _ok(self):
        b = self.ed.resultado()
        if b:
            self.resultado = core.Calendarios(b, self.cal0.recursos)
            self.accept()


class DialogoRecursos(QDialog):
    """Calendario propio por recurso. Los feriados globales aplican a todos además de los propios."""
    def __init__(self, cal, nombres, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Calendarios por recurso")
        self.setMinimumSize(860, 640)
        self.cal0 = cal
        self.edit = {n: cal.recursos.get(n) for n in dict.fromkeys(list(nombres) + list(cal.recursos))}
        self.actual = None
        lay = QHBoxLayout(self)
        izq = QVBoxLayout()
        izq.addWidget(QLabel("Recursos (de las tareas):"))
        self.lista = QListWidget()
        izq.addWidget(self.lista)
        self.nuevo = QLineEdit()
        self.nuevo.setPlaceholderText("Agregar recurso…")
        bn = QPushButton("Agregar")
        bn.clicked.connect(self._agregar)
        fn = QHBoxLayout()
        fn.addWidget(self.nuevo)
        fn.addWidget(bn)
        izq.addLayout(fn)
        lay.addLayout(izq, 1)
        der = QVBoxLayout()
        self.titulo = QLabel("Selecciona un recurso")
        der.addWidget(self.titulo)
        self.propio = QCheckBox("Usar calendario propio (si no, usa el calendario global)")
        self.propio.toggled.connect(self._toggle)
        der.addWidget(self.propio)
        self.ed = EditorCalendario(cal.base, self)
        der.addWidget(self.ed)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        der.addWidget(bb)
        lay.addLayout(der, 3)
        self._llenar()
        self.lista.currentTextChanged.connect(self._cambiar)
        self.ed.setEnabled(False)
        self.propio.setEnabled(False)
        self.resultado = None

    def _etiqueta(self, n):
        return ("● " if self.edit.get(n) else "   ") + n

    def _llenar(self, sel=None):
        self.lista.blockSignals(True)
        self.lista.clear()
        for n in self.edit:
            self.lista.addItem(self._etiqueta(n))
        self.lista.blockSignals(False)
        if sel:
            for r in range(self.lista.count()):
                if self.lista.item(r).text().strip("● ").strip() == sel:
                    self.lista.setCurrentRow(r)

    def _nombre(self, txt):
        return txt.replace("●", "").strip()

    def _guardar_actual(self):
        if self.actual is None:
            return True
        if self.propio.isChecked():
            c = self.ed.resultado()
            if c is None:
                return False
            self.edit[self.actual] = c
        else:
            self.edit[self.actual] = None
        return True

    def _cambiar(self, txt):
        if not self._guardar_actual():
            return
        n = self._nombre(txt)
        self.actual = n or None
        if not n:
            return
        c = self.edit.get(n)
        self.titulo.setText(f"Recurso: {n}")
        self.propio.setEnabled(True)
        self.propio.blockSignals(True)
        self.propio.setChecked(c is not None)
        self.propio.blockSignals(False)
        self.ed.cargar(c or self.cal0.base)
        self.ed.setEnabled(c is not None)
        # refresca marcas sin perder la selección
        self.lista.blockSignals(True)
        for r in range(self.lista.count()):
            self.lista.item(r).setText(self._etiqueta(self._nombre(self.lista.item(r).text())))
        self.lista.blockSignals(False)

    def _toggle(self, on):
        self.ed.setEnabled(on)
        if on and self.actual and self.edit.get(self.actual) is None:
            self.ed.cargar(self.cal0.base)

    def _agregar(self):
        n = self.nuevo.text().strip()
        if n and n not in self.edit:
            self.edit[n] = None
            self.nuevo.clear()
            self._llenar(n)

    def _ok(self):
        if not self._guardar_actual():
            return
        self.resultado = core.Calendarios(self.cal0.base, {n: c for n, c in self.edit.items() if c})
        self.accept()


def cal_a_json(cal):
    cal = core.como_calendarios(cal)
    def uno(c):
        return {"h": c.horas, "d": list(c.dias), "l": [[f.toordinal(), m] for f, m in c.libres],
                "c": [c.ciclo[0], c.ciclo[1], c.ciclo[2].toordinal()] if c.ciclo else None}
    return json.dumps({"base": uno(cal.base), "rec": {n: uno(c) for n, c in cal.recursos.items()}})


def cal_de_json(txt):
    def uno(o):
        c = o.get("c")
        return core.Calendario(o["h"], o["d"], [(dt.date.fromordinal(f), m) for f, m in o["l"]],
                               (c[0], c[1], dt.date.fromordinal(c[2])) if c else None)
    try:
        o = json.loads(txt)
        if "base" not in o:                                   # formato anterior (solo calendario global)
            return core.Calendarios(uno(o))
        return core.Calendarios(uno(o["base"]), {n: uno(c) for n, c in o["rec"].items()})
    except Exception:
        return None


class Ventana(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PSD Gantt")
        self.resize(1400, 760)
        self.ajustes = QSettings("LuxMinerals", "PSD_Gantt", self)
        self.tareas, self.ruta, self.sucio, self._reconstruyendo = [], None, False, False
        self.cal = cal_de_json(self.ajustes.value("calendario", "")) or core.Calendarios()

        self.tabla = QTableWidget(0, len(COLS))
        self.tabla.setHorizontalHeaderLabels(COLS)
        self.tabla.verticalHeader().hide()
        self.tabla.verticalHeader().setDefaultSectionSize(RH)
        self.tabla.horizontalHeader().setFixedHeight(HH)
        self.tabla.horizontalHeader().setStretchLastSection(False)
        for i, a in enumerate(ANCHOS):
            self.tabla.setColumnWidth(i, a)
        self.tabla.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.tabla.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.tabla.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tabla.itemChanged.connect(self.editado)

        self.gantt = GanttWidget()
        self.area = QScrollArea()
        self.area.setWidget(self.gantt)
        self.area.setFrameShape(QScrollArea.NoFrame)
        self.area.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.tabla.verticalScrollBar().valueChanged.connect(self.area.verticalScrollBar().setValue)
        self.area.verticalScrollBar().valueChanged.connect(self.tabla.verticalScrollBar().setValue)

        sp = QSplitter()
        sp.addWidget(self.tabla)
        sp.addWidget(self.area)
        sp.setStretchFactor(1, 1)
        sp.setSizes([760, 640])
        self.setCentralWidget(sp)

        tb = QToolBar("Principal")
        tb.setMovable(False)
        self.addToolBar(tb)
        for it_ in (("Nuevo", self.nuevo), ("Abrir Excel…", self.abrir), ("Guardar Excel…", self.guardar),
                        None, ("Calendario global…", self.editar_calendario),
                        ("Calendarios por recurso…", self.editar_recursos), None, ("+ Tarea", lambda: self.agregar(False)), ("+ Subtarea", lambda: self.agregar(True)),
                        ("Eliminar", self.eliminar), ("Ordenar por ID", self.ordenar),
                        ("Replanificar por predecesoras", self.replanificar)):
            if it_ is None:
                tb.addSeparator()
            else:
                tb.addAction(it_[0], it_[1])
        self.addToolBarBreak()
        tb = QToolBar("Salida")
        tb.setMovable(False)
        self.addToolBar(tb)
        for it_ in (("Exportar PNG…", lambda: self.exportar("png")), ("Exportar PDF…", lambda: self.exportar("pdf")),
                    None, ("Copiar para Project", self.copiar_project), ("Exportar para Project…", self.exportar_project),
                    None, ("Cargar ejemplo", self.cargar_ejemplo)):
            if it_ is None:
                tb.addSeparator()
            else:
                tb.addAction(it_[0], it_[1])
        tb.addSeparator()
        tb.addWidget(QLabel(" Escala: "))
        self.zoom = QComboBox()
        self.zoom.addItems(list(ZOOM))
        self.zoom.setCurrentText(self.ajustes.value("zoom", "Semana"))
        self.zoom.currentTextChanged.connect(lambda z: (self.ajustes.setValue("zoom", z), self.refrescar()))
        tb.addWidget(self.zoom)
        self.lbl = QLabel("")
        self.statusBar().addPermanentWidget(self.lbl)
        self.statusBar().showMessage("Azul = tarea · Amarillo = opcional · Barra gris = resumen · Borde rojo = atrasada · Línea naranja = hoy · Fondo oscuro = día sin trabajo (violeta = libre del recurso)")

        ult = self.ajustes.value("ultimo", "")
        cargado = False
        if ult and os.path.isfile(ult):
            try:
                self.tareas, cal = core.leer_excel(ult)
                self.ruta, cargado = ult, True
                self.cal = cal or self.cal
            except Exception:
                pass
        self.refrescar()
        if not cargado:
            QTimer.singleShot(300, self.comenzar)

    # ---- tabla ↔ datos
    def refrescar(self):
        filas = core.calcular(self.tareas, self.cal)
        self._reconstruyendo = True
        self.tabla.setRowCount(len(filas))
        fb = QFont(self.tabla.font())
        fb.setBold(True)
        for r, fl in enumerate(filas):
            t = fl.t
            dur = (self.cal.contar(fl.ini, fl.fin) if fl.ini else "") if fl.resumen else t.dur
            vals = [t.id, ("    " * fl.nivel) + t.nombre, t.resp, core.fmt(fl.ini), str(dur), core.fmt(fl.fin),
                    f"{fl.horas:g}", f"{fl.av:.0f}", t.pred, "Sí" if t.opcional else ""]
            for c, v in enumerate(vals):
                it = QTableWidgetItem(v)
                if fl.resumen:
                    it.setFont(fb)
                    if c in SOLO_LECTURA_RESUMEN:
                        it.setFlags(it.flags() & ~Qt.ItemIsEditable)
                        it.setForeground(_c(MUTED))
                if c == 6:
                    it.setFlags(it.flags() & ~Qt.ItemIsEditable)
                    it.setForeground(_c(MUTED))
                if 3 <= c <= 7:
                    it.setTextAlignment(Qt.AlignCenter)
                self.tabla.setItem(r, c, it)
        sr = {t.id for t in core.sin_recurso(self.tareas)}
        for r, fl in enumerate(filas):
            if fl.t.id in sr:
                it = self.tabla.item(r, 2)
                it.setBackground(_c("#5a2b2b"))
                it.setToolTip("Sin recurso: usa el calendario global")
        self._reconstruyendo = False
        self.gantt.configurar(filas, ZOOM[self.zoom.currentText()], self.cal)
        
        pr = core.validar(self.tareas)
        self.lbl.setText(("⚠ " + pr[0] + (f" (+{len(pr)-1})" if len(pr) > 1 else "")) if pr
                         else f"{len(self.tareas)} tareas" + (f" · ⚠ {len(sr)} sin recurso" if sr else "")
                         + f" · {self.cal.resumen()}")
        self.setWindowTitle(f"PSD Gantt — {os.path.basename(self.ruta) if self.ruta else 'sin guardar'}"
                            + (" *" if self.sucio else ""))

    def editado(self, it):
        if self._reconstruyendo:
            return
        r, c, v = it.row(), it.column(), it.text().strip()
        if r >= len(self.tareas):
            return
        t = self.tareas[r]
        try:
            if c == 0:
                t.id = v
            elif c == 1:
                t.nombre = v.strip()
            elif c == 2:
                t.resp = v
            elif c == 3:
                f = core.parse_fecha(v)
                t.inicio = self.cal.de(t).desde(f) if f else None
            elif c == 4:
                t.dur = int(float(v.replace(",", ".")))
                if t.dur < 1:
                    t.dur = 1
                    raise ValueError("La duración mínima es 1 día hábil.")
            elif c == 5:
                f = core.parse_fecha(v)
                ini = self.cal.de(t).desde(t.inicio) if t.inicio else None
                if f and ini:
                    if f < ini:
                        raise ValueError("El fin no puede ser anterior al inicio.")
                    t.dur = self.cal.de(t).contar(ini, f)
            elif c == 7:
                t.avance = max(0.0, min(100.0, float(v.replace(",", ".") or 0)))
            elif c == 8:
                t.pred = v
            elif c == 9:
                t.opcional = v.lower() in ("sí", "si", "x", "1")
        except ValueError as e:
            QMessageBox.warning(self, "Dato no válido", str(e))
        self.sucio = True
        self.refrescar()

    # ---- acciones
    def _confirmar(self):
        if not self.sucio:
            return True
        r = QMessageBox.question(self, "Cambios sin guardar", "Hay cambios sin guardar. ¿Continuar y perderlos?")
        return r == QMessageBox.Yes

    def comenzar(self):
        """Pantalla de inicio: desde cero, leer Excel o ejemplo."""
        if self.tareas or self.ruta:
            return
        m = QMessageBox(self)
        m.setWindowTitle("PSD Gantt")
        m.setText("¿Cómo quieres comenzar?")
        b0 = m.addButton("Comenzar desde cero", QMessageBox.AcceptRole)
        b1 = m.addButton("Leer archivo Excel…", QMessageBox.AcceptRole)
        b2 = m.addButton("Ver ejemplo", QMessageBox.ActionRole)
        m.addButton("Cancelar", QMessageBox.RejectRole)
        m.exec()
        c = m.clickedButton()
        if c is b0:
            self.nuevo(confirmar=False)
        elif c is b1:
            self.abrir(confirmar=False)
        elif c is b2:
            self.cargar_ejemplo(confirmar=False)

    def pedir_calendario(self, titulo_ok=True):
        dlg = DialogoCalendario(self.cal, self)
        dlg.setWindowTitle("Calendario de trabajo" + (" — define horas y días de trabajo" if titulo_ok else ""))
        if dlg.exec() == QDialog.Accepted:
            self.cal = dlg.resultado
            self.ajustes.setValue("calendario", cal_a_json(self.cal))
            return True
        return False

    def editar_calendario(self):
        if self.pedir_calendario(False):
            self.sucio = True
            self.refrescar()

    def editar_recursos(self):
        dlg = DialogoRecursos(self.cal, core.recursos_en_uso(self.tareas), self)
        if dlg.exec() == QDialog.Accepted:
            self.cal = dlg.resultado
            self.ajustes.setValue("calendario", cal_a_json(self.cal))
            self.sucio = True
            self.refrescar()

    def _confirmar_sin_recurso(self, accion):
        """«Obligar a asignar recurso»: se avisa y se pide confirmar antes de replanificar o exportar."""
        sr = core.sin_recurso(self.tareas)
        if not sr:
            return True
        lista = "\n".join(f"  {t.id}  {t.nombre[:50]}" for t in sr[:10]) + (f"\n  … y {len(sr) - 10} más" if len(sr) > 10 else "")
        r = QMessageBox.question(self, "Tareas sin recurso",
                                 f"{len(sr)} tarea(s) no tienen recurso asignado y usan el calendario global:\n{lista}\n\n"
                                 f"¿{accion} igual?")
        return r == QMessageBox.Yes

    def nuevo(self, confirmar=True):
        if confirmar and not self._confirmar():
            return
        if not self.pedir_calendario():
            return
        self.tareas, self.ruta, self.sucio = [], None, False
        self.refrescar()
        self.agregar(False)

    def cargar_ejemplo(self, confirmar=True):
        if confirmar and not self._confirmar():
            return
        self.cal = core.ejemplo_calendario()
        self.tareas, self.ruta, self.sucio = core.ejemplo(), None, True
        self.refrescar()

    def abrir(self, confirmar=True):
        if confirmar and not self._confirmar():
            return
        ruta, _ = QFileDialog.getOpenFileName(self, "Abrir cronograma", self.ajustes.value("ultimo", ""), "Excel (*.xlsx)")
        if not ruta:
            return
        if core.es_project(ruta):                      # exportación a Excel de Microsoft Project
            try:
                tareas, cal, aviso = core.leer_project(ruta)
            except Exception as e:
                QMessageBox.critical(self, "No se pudo leer el archivo de Project", str(e))
                return
            QMessageBox.information(self, "Importado desde Project", aviso + "\n\nRevisa el calendario en «Calendario…» "
                                    "antes de replanificar. Al guardar se crea un Excel nuevo; el original no se modifica.")
            self.cal, self.tareas, self.ruta, self.sucio = cal, tareas, None, True
            self.ajustes.setValue("calendario", cal_a_json(cal))
            self.refrescar()
            return
        try:
            tareas, cal = core.leer_excel(ruta)
        except Exception as e:
            QMessageBox.critical(self, "No se pudo leer", str(e))
            return
        if cal is None:                                # el archivo no trae calendario → lo defines tú
            QMessageBox.information(self, "Calendario", "Este archivo no trae hoja «Calendario». "
                                    "Define las horas por día y los días sin trabajo.")
            if not self.pedir_calendario():
                return
        else:
            self.cal = cal
            self.ajustes.setValue("calendario", cal_a_json(cal))
        self.tareas, self.ruta, self.sucio = tareas, ruta, cal is None
        self.ajustes.setValue("ultimo", ruta)
        self.refrescar()

    def guardar(self):
        ruta, _ = QFileDialog.getSaveFileName(self, "Guardar cronograma", self.ruta or "cronograma_psd.xlsx", "Excel (*.xlsx)")
        if not ruta:
            return
        if not ruta.lower().endswith(".xlsx"):
            ruta += ".xlsx"
        try:
            core.guardar_excel(self.tareas, self.cal, ruta)
        except Exception as e:
            QMessageBox.critical(self, "No se pudo guardar", str(e))
            return
        self.ruta, self.sucio = ruta, False
        self.ajustes.setValue("ultimo", ruta)
        self.refrescar()

    # ---- Microsoft Project
    def copiar_project(self):
        pr = core.validar(self.tareas)
        if pr:
            QMessageBox.warning(self, "Revisa los datos", "\n".join(pr[:8]))
            return
        if not self._confirmar_sin_recurso("Copiar"):
            return
        QApplication.clipboard().setText(core.texto_project(self.tareas, self.cal))
        self.statusBar().showMessage("Copiado. En Project muestra, en este orden, las columnas: " + " · ".join(core.PROJECT_COLS)
                                     + " — selecciona la primera celda de «Nivel de esquema» y pega (Cmd/Ctrl+V).", 15000)

    def exportar_project(self):
        pr = core.validar(self.tareas)
        if pr:
            QMessageBox.warning(self, "Revisa los datos", "\n".join(pr[:8]))
            return
        if not self._confirmar_sin_recurso("Exportar"):
            return
        base = os.path.splitext(self.ruta or "cronograma_psd")[0] + "_Pegar_en_Project.xlsx"
        ruta, _ = QFileDialog.getSaveFileName(self, "Exportar para Project", base, "Excel (*.xlsx)")
        if ruta:
            if not ruta.lower().endswith(".xlsx"):
                ruta += ".xlsx"
            core.guardar_project(self.tareas, self.cal, ruta)
            self.statusBar().showMessage("Exportado: " + ruta + " (mismo formato que tus planillas «Pegar_en_Project»)", 10000)

    def _fila_sel(self):
        s = self.tabla.selectionModel().selectedRows()
        return s[0].row() if s else None

    def agregar(self, sub):
        r = self._fila_sel()
        hoy = dt.date.today()
        ids = {t.id for t in self.tareas}
        if sub and r is not None:
            base, n = self.tareas[r].id, 1
            while f"{base}-{n}" in ids:
                n += 1
            nid, pos = f"{base}-{n}", r + 1
            while pos < len(self.tareas) and self.tareas[pos].id.startswith(base + "-"):
                pos += 1
        else:
            n = 1
            while f"T-{n:02d}" in ids:
                n += 1
            nid, pos = f"T-{n:02d}", len(self.tareas)
        self.tareas.insert(pos, core.Tarea(nid, "Nueva tarea", "", self.cal.desde(hoy), 5))
        self.sucio = True
        self.refrescar()
        self.tabla.selectRow(pos)

    def eliminar(self):
        r = self._fila_sel()
        if r is None:
            return
        t = self.tareas[r]
        sub = [x for x in self.tareas if x.id.startswith(t.id + "-")]
        msg = f"¿Eliminar {t.id} «{t.nombre}»" + (f" y sus {len(sub)} subtareas?" if sub else "?")
        if QMessageBox.question(self, "Eliminar", msg) == QMessageBox.Yes:
            self.tareas = [x for x in self.tareas if x is not t and x not in sub]
            self.sucio = True
            self.refrescar()

    def ordenar(self):
        self.tareas.sort(key=lambda t: core.clave_natural(t.id))
        self.sucio = True
        self.refrescar()

    def replanificar(self):
        pr = core.validar(self.tareas)
        if pr:
            QMessageBox.warning(self, "Revisa los datos", "\n".join(pr[:8]))
            return
        if not self._confirmar_sin_recurso("Replanificar"):
            return
        try:
            n = core.replanificar(self.tareas, self.cal)
        except ValueError as e:
            QMessageBox.warning(self, "Replanificar", str(e))
            return
        self.sucio = self.sucio or n > 0
        self.refrescar()
        self.statusBar().showMessage(f"{n} tarea(s) movida(s) según predecesoras (Fin→Inicio, solo hacia adelante).", 6000)

    # ---- exportación
    def _dibujar_completo(self, p, ancho_panel):
        filas = self.gantt.filas
        f0 = p.font()
        fb = QFont(f0)
        fb.setBold(True)
        h = HH + len(filas) * RH
        p.fillRect(QRectF(0, 0, ancho_panel, h), _c(SURFACE))
        cols = [("ID", 70), ("Tarea", 250), ("Resp.", 90), ("Inicio", 62), ("Fin", 62), ("Días", 36), ("Horas", 44), ("%", 34)]
        x = 0
        p.setPen(_c(INK))
        p.setFont(fb)
        for nom, a in cols:
            p.drawText(QRectF(x + 3, 0, a, HH), Qt.AlignVCenter | Qt.AlignLeft, nom)
            x += a
        p.setFont(f0)
        for i, fl in enumerate(filas):
            y = HH + i * RH
            if i % 2:
                p.fillRect(QRectF(0, y, ancho_panel, RH), _c(BANDA))
            t = fl.t
            dias = self.cal.de(fl.t).contar(fl.ini, fl.fin) if fl.ini else ""
            vals = [t.id, ("  " * fl.nivel) + t.nombre, t.resp, core.fmt(fl.ini), core.fmt(fl.fin), str(dias),
                    f"{fl.horas:g}", f"{fl.av:.0f}"]
            x = 0
            p.setFont(fb if fl.resumen else f0)
            p.setPen(_c(INK))
            for (nom, a), v in zip(cols, vals):
                txt = QFontMetrics(p.font()).elidedText(v, Qt.ElideRight, a - 6)
                p.drawText(QRectF(x + 3, y, a - 3, RH), Qt.AlignVCenter | Qt.AlignLeft, txt)
                x += a
        p.setFont(f0)
        p.setPen(QPen(_c(GRID)))
        p.drawLine(QPointF(ancho_panel, 0), QPointF(ancho_panel, h))
        g = self.gantt
        return pintar(p, filas, g.d0, g.ndias, g.ppd, x0=ancho_panel, hoy=dt.date.today(), cal=self.cal)

    def exportar(self, fmt):
        if not self.tareas:
            QMessageBox.information(self, "Exportar", "No hay tareas para exportar.")
            return
        base = os.path.splitext(self.ruta or "cronograma_psd")[0] + "." + fmt
        ruta, _ = QFileDialog.getSaveFileName(self, "Exportar", base, f"{fmt.upper()} (*.{fmt})")
        if not ruta:
            return
        if not ruta.lower().endswith("." + fmt):
            ruta += "." + fmt
        panel = 648
        g = self.gantt
        W, H = panel + g.ndias * g.ppd, HH + len(g.filas) * RH
        if fmt == "png":
            esc = 2
            img = QImage(W * esc, H * esc, QImage.Format_ARGB32)
            img.fill(_c(SURFACE))
            p = QPainter(img)
            p.scale(esc, esc)
            p.setRenderHint(QPainter.Antialiasing)
            self._dibujar_completo(p, panel)
            p.end()
            ok = img.save(ruta)
        else:
            pw = QPdfWriter(ruta)
            pw.setPageSize(QPageSize(QSizeF_(W + 40, H + 40), QPageSize.Point))
            pw.setPageMargins(QMarginsF(0, 0, 0, 0))
            pw.setResolution(72)
            p = QPainter(pw)
            p.fillRect(QRectF(0, 0, W + 40, H + 40), _c(SURFACE))
            p.translate(20, 20)
            p.setRenderHint(QPainter.Antialiasing)
            self._dibujar_completo(p, panel)
            p.end()
            ok = True
        self.statusBar().showMessage(("Exportado: " + ruta) if ok else "No se pudo exportar.", 8000)

    def closeEvent(self, e):
        if self._confirmar():
            e.accept()
        else:
            e.ignore()


def QSizeF_(w, h):
    from PySide6.QtCore import QSizeF
    return QSizeF(w, h)


def aplicar_tema(app):
    app.setStyle("Fusion")
    p = QPalette()
    for rol, col in ((QPalette.Window, "#232322"), (QPalette.WindowText, INK), (QPalette.Base, SURFACE),
                     (QPalette.AlternateBase, "#242423"), (QPalette.Text, INK), (QPalette.Button, "#2f2f2d"),
                     (QPalette.ButtonText, INK), (QPalette.Highlight, BLUE), (QPalette.HighlightedText, "#ffffff"),
                     (QPalette.ToolTipBase, "#2b2b29"), (QPalette.ToolTipText, INK)):
        p.setColor(rol, _c(col))
    for rol in (QPalette.WindowText, QPalette.Text, QPalette.ButtonText):
        p.setColor(QPalette.Disabled, rol, _c("#7a7975"))
    app.setPalette(p)
    app.setStyleSheet("QToolTip{background:#2b2b29;color:#f2f2f0;border:1px solid #6b6a65;}"
                      "QHeaderView::section{background:#2f2f2d;color:#f2f2f0;padding:4px;border:1px solid #3d3d3a;}"
                      "QTableView{gridline-color:#3d3d3a;}")


def main(argv):
    app = QApplication(argv)
    aplicar_tema(app)
    f = app.font()
    f.setPointSize(10)
    app.setFont(f)
    v = Ventana()
    v.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main(sys.argv))
