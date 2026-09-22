"""
Detector e importador de fórmulas desde un JSON de cierre (format x-spreadsheet,
salida de `panelWdgt[id]['elements']`).

Sin modificar la posición ni la estructura del JSON: solo lee las celdas,
detecta las fórmulas fltFLoatTablaDinamica, les asigna un NOMBRE (etiqueta de
la fila) y su COORDENADA A1, e identifica qué fórmulas se "tocan" entre sí
(comparten referencias de fuente: codigos $CD, celdas [N~...~C_XXX], funciones
<T.func,...>), o que referencian el nombre/coordenada de otra.

Si una fórmula es editada, se recalcula la integridad: las fórmulas que
comparten/usan la misma fuente quedan marcadas como "requiere actualización".

Solo depende de PySide6 y la stdlib.
"""

import io
import json
import os
import re

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QTableWidget,
    QTableWidgetItem, QHeaderView, QAbstractItemView, QFileDialog, QSplitter,
    QListWidget, QListWidgetItem, QGroupBox, QTextEdit, QDialog,
    QDialogButtonBox, QMessageBox, QSizePolicy,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QBrush

# ──────────────────────────────────────────────
# Utilidades de celda / A1
# ──────────────────────────────────────────────
def col_letter(idx: int) -> str:
    """Índice 0-based → letra de columna (0→A, 25→Z, 26→AA)."""
    s = ""
    x = idx
    while True:
        s = chr(ord('A') + (x % 26)) + s
        x = x // 26 - 1
        if x < 0:
            break
    return s


def a1(row: int, col: int) -> str:
    return f"{col_letter(col)}{row + 1}"


_TEMP_COL = {"DIA": 0, "SEMANA": 1, "MES": 2, "AÑO": 3, "ANO": 3}


def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s)).strip()


_ACCENTS = {
    "Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U",
    "Ñ": "N", "Ü": "U",
}


def _upper_norm(s) -> str:
    """Normaliza a mayúsculas sin tildes (DÍA→DIA, AÑO→ANO)."""
    t = _norm(s).upper()
    for a, b in _ACCENTS.items():
        t = t.replace(a, b)
    return t


# ──────────────────────────────────────────────
# Parser del JSON de "elements"
# ──────────────────────────────────────────────
def load_elements(path: str):
    """Carga el JSON de elements del cierre.

    El archivo suele venir con el contenido con doble codificación
    (string JSON escapado dentro del JSON), así que se intenta 'json.loads'
    dos veces cuando el primer parseo devuelve un str.
    Devuelve: lista de hojas (dicts x-spreadsheet).
    """
    with io.open(path, "r", encoding="utf-8") as f:
        raw = f.read()
    data = json.loads(raw)
    if isinstance(data, str):
        data = json.loads(data)
    if isinstance(data, dict):
        data = [data]
    if not isinstance(data, list):
        raise ValueError("La raíz del JSON no es una lista de hojas.")
    return data


def _cell_text(rows, r, c) -> str:
    rv = rows.get(str(r))
    if not isinstance(rv, dict):
        return ""
    cv = rv.get("cells", {}).get(str(c), {})
    if not isinstance(cv, dict):
        return ""
    return cv.get("text", "") or ""


def _is_formula(text: str) -> bool:
    t = _norm(text).lower()
    if not t:
        return False
    if "fltfloattabladinamica" in t:
        return True
    if t.startswith("="):
        return True
    if re.match(r"^\[\d+~[a-z_]+~C_[0-9a-z_]+\]", t):
        return True
    return False


def build_temp_map(rows) -> dict:
    """Detecta la fila de cabecera DÍA/SEMANA/MES/AÑO y mapea columna→temporalidad."""
    for r in range(0, 20):
        found = {}
        for c in range(0, 20):
            t = _upper_norm(_cell_text(rows, r, c))
            if t in _TEMP_COL:
                found[c] = t
        if len(found) >= 2:
            return found
    return {}


def _indicator_name(rows, r, c) -> str:
    """Busca la etiqueta del indicador en la misma fila (cols 0..3, preferida col 1)."""
    best = ""
    for cc in (1, 0, 2, 3):
        if cc == c:
            continue
        t = _norm(_cell_text(rows, r, cc))
        if t and not _is_formula(t) and _upper_norm(t) not in _TEMP_COL:
            best = t
            break
    return best


# ──────────────────────────────────────────────
# Extracción de referencias / tokens de fuente
# ──────────────────────────────────────────────
_CODIGO_RE = re.compile(r"\$([A-Za-z0-9_]+),")
_CELDA_RE  = re.compile(r"\[(\d+)~([A-Za-z0-9_]+)~([A-Za-z0-9_]+)\]")
_FUNC_RE   = re.compile(r"<T\.([A-Za-z0-9_]+),")


def extract_refs(text: str) -> set:
    """Devuelve el conjunto de tokens de fuente usados por la fórmula."""
    if not isinstance(text, str):
        return set()
    refs = set()
    for m in _CODIGO_RE.finditer(text):
        refs.add("$" + m.group(1))
    for m in _CELDA_RE.finditer(text):
        refs.add(f"REF:{m.group(2)}/{m.group(3)}")
    for m in _FUNC_RE.finditer(text):
        refs.add("T." + m.group(1))
    return refs


def extract_formulas(sheet) -> list:
    """Detecta todas las celdas con fórmula en la hoja.

    Respeta la estructura original (no la modifica). Devuelve una lista de
    dicts con: r, c, a1, name, temp, text, refs.
    """
    rows = sheet.get("rows", {})
    temp_map = build_temp_map(rows)
    out = []
    for rk, rv in rows.items():
        if rk == "len" or not isinstance(rv, dict):
            continue  # 'len' especial de x-spreadsheet
        r = int(rk)
        cells = rv.get("cells", {})
        if not isinstance(cells, dict):
            continue
        for ck, cv in cells.items():
            if not isinstance(cv, dict):
                continue
            text = cv.get("text", "")
            if not _is_formula(text):
                continue
            c = int(ck)
            temp = temp_map.get(c, "")
            out.append({
                "r": r,
                "c": c,
                "a1": a1(r, c),
                "name": _indicator_name(rows, r, c) or f"Fila {r + 1}",
                "temp": temp,
                "text": text,
                "refs": extract_refs(text),
                "modificada": False,
                "estado": "original",
            })
    out.sort(key=lambda e: (e["r"], e["c"]))
    return out


# ──────────────────────────────────────────────
# Relaciones / integridad
# ──────────────────────────────────────────────
def touches(a, b) -> bool:
    """Verdadero si a y b comparten fuente, o a referencia el nombre/coord de b."""
    if a is b:
        return True
    if a["refs"] & b["refs"]:
        return True
    # a referencia el nombre de b (como término literal)
    b_name = a_norm_name(b["name"])
    if b_name and len(b_name) >= 3 and b_name in _norm(a["text"]).lower():
        return True
    # a referencia la coordenada de b
    if b["a1"] in a["text"]:
        return True
    return False


def a_norm_name(name) -> str:
    return _norm(name).lower()


def touching_indices(records: list, i: int) -> list:
    sel = records[i]
    return [j for j, r in enumerate(records) if j != i and touches(sel, r)]


# ──────────────────────────────────────────────
# Diálogo de edición de fórmula
# ──────────────────────────────────────────────
class FormulaEditDialog(QDialog):
    def __init__(self, title, initial="", parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(900, 420)
        lay = QVBoxLayout(self)
        info = QLabel(
            "Edita la fórmula. Al guardar, la fórmula quedará marcada como "
            "modificada y se recalculará la integridad:\nlas fórmulas que comparten "
            "la misma fuente se marcarán como 'requiere actualización'."
        )
        info.setWordWrap(True)
        info.setStyleSheet("color:#8B949E; font-size:11px;")
        self.edit = QTextEdit()
        self.edit.setPlainText(initial)
        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        lay.addWidget(info)
        lay.addWidget(self.edit)
        lay.addWidget(btns)

    def text(self):
        return self.edit.toPlainText()


# ──────────────────────────────────────────────
# Widget de la pestaña
# ──────────────────────────────────────────────
class ImportadorCierreTab(QWidget):
    COLUMNS = ["Celda", "Nombre", "Temp.", "Fuentes", "Estado", "Vista previa"]

    def __init__(self, builder=None, parent=None):
        super().__init__(parent)
        self.builder = builder
        self.records = []
        self._gray = QColor("#21262D")
        self._touch_color = QColor("#16324B")
        self._sel_color = QColor("#1F3A58")
        self._mod_color = QColor("#3A2A16")
        self._need_color = QColor("#3A1616")
        self._build_ui()

    # ── UI ────────────────────────────────────
    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(8)
        root.addWidget(self._build_toolbar())

        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.setHandleWidth(3)
        self.splitter.addWidget(self._build_grid())
        self.splitter.addWidget(self._build_detail())
        self.splitter.setSizes([760, 400])
        root.addWidget(self.splitter, 1)

    def _btn(self, text, border, hover=None, fg=None, padding="5px 12px", fs=11, bold=False):
        from formula_builder import make_btn  # lazy: evita import circular
        fg = fg or border
        return f"""
            QPushButton {{
                background-color: #1C2128;
                color: {fg};
                border: 1px solid {border};
                border-radius: 6px;
                padding: {padding};
                font-size: {fs}px;
                font-weight: {'600' if bold else '400'};
                font-family: 'Consolas','Courier New',monospace;
            }}
            QPushButton:hover {{ background-color: {hover or border}; color: #E6EDF3; }}
            QPushButton:pressed {{ background-color: {border}; }}
        """

    def _build_toolbar(self):
        bar = QWidget()
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(8)

        b_open = QPushButton("📂 Abrir elements JSON")
        b_open.setStyleSheet(self._btn("", "#BC8CFF", "#1A1A3A", bold=True))
        b_open.setToolTip("Carga el JSON del cierre (panelWdgt[id]['elements'])")
        b_open.clicked.connect(self._load_json)

        b_integridad = QPushButton("🔄 Verificar integridad")
        b_integridad.setStyleSheet(self._btn("", "#58A6FF", "#1A3A5A"))
        b_integridad.clicked.connect(self._recompute_integridad)

        b_reset = QPushButton("↺ Reset ediciones")
        b_reset.setStyleSheet(self._btn("", "#7A828E", "#3A3F48"))
        b_reset.clicked.connect(self._reset_ediciones)

        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

        self.lbl_status = QLabel("Sin archivo cargado")
        self.lbl_status.setStyleSheet("color:#8B949E; font-size:11px;")

        lay.addWidget(b_open)
        lay.addWidget(b_integridad)
        lay.addWidget(b_reset)
        lay.addWidget(spacer)
        lay.addWidget(self.lbl_status)
        return bar

    def _build_grid(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(4)

        self.table = QTableWidget(0, len(self.COLUMNS))
        self.table.setHorizontalHeaderLabels(self.COLUMNS)
        self.table.setStyleSheet(f"""
            QTableWidget {{
                background: #1C2128;
                alternate-background-color: #161B22;
                color: #E6EDF3;
                border: 1px solid #30363D;
                border-radius: 6px;
                font-family: 'Menlo','Consolas',monospace;
                font-size: 10px;
                gridline-color: #30363D;
            }}
            QTableWidget::item:selected {{
                background: #58A6FF;
                color: #0D1117;
            }}
            QHeaderView::section {{
                background: #161B22;
                color: #8B949E;
                border: none;
                border-right: 1px solid #30363D;
                padding: 6px;
                font-weight: 600;
            }}
        """)
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        hh.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        vh = self.table.verticalHeader()
        vh.setDefaultSectionSize(40)
        self.table.setColumnWidth(0, 70)
        self.table.setColumnWidth(1, 230)
        self.table.setColumnWidth(2, 60)
        self.table.setColumnWidth(3, 150)
        self.table.setColumnWidth(4, 110)
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.itemSelectionChanged.connect(self._refresh_detail)
        self.table.cellDoubleClicked.connect(self._edit_record)
        self.table.cellClicked.connect(self._on_cell_clicked)
        lay.addWidget(self.table)

        hint = QLabel(
            "Doble clic en una fila edita la fórmula. Clic simple resalta en la tabla "
            "las fórmulas que se tocan (comparten fuente o referencian su nombre/coordenada)."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color:#484F58; font-size:10px;")
        lay.addWidget(hint)
        return w

    def _build_detail(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(10, 0, 0, 0)
        lay.setSpacing(6)

        self.lbl_sel = QLabel("Selecciona una fórmula…")
        self.lbl_sel.setStyleSheet("color:#58A6FF; font-size:13px; font-weight:700;")
        self.lbl_sel.setWordWrap(True)
        lay.addWidget(self.lbl_sel)

        self.lbl_resumen = QLabel("")
        self.lbl_resumen.setWordWrap(True)
        self.lbl_resumen.setStyleSheet(
            "color:#8B949E; font-size:11px; background:#21262D; padding:6px 8px; border-radius:4px;")
        lay.addWidget(self.lbl_resumen)

        self.detail_edit = QTextEdit()
        self.detail_edit.setReadOnly(True)
        self.detail_edit.setPlaceholderText("Texto de la fórmula…")
        self.detail_edit.setStyleSheet(f"""
            QTextEdit {{
                background:#21262D; color:#3FB950; border:1px solid #30363D;
                border-radius:6px; padding:8px; font-family:'Consolas',monospace; font-size:11px;
            }}
        """)
        lay.addWidget(self.detail_edit)

        self.grp_touch = QGroupBox("Se tocan (comparten fuente / referencia)")
        tl = QVBoxLayout(self.grp_touch)
        self.list_touch = QListWidget()
        self.list_touch.itemClicked.connect(self._goto)
        tl.addWidget(self.list_touch)
        lay.addWidget(self.grp_touch)

        self.grp_need = QGroupBox("Requieren actualización")
        nl = QVBoxLayout(self.grp_need)
        self.list_need = QListWidget()
        self.list_need.itemClicked.connect(self._goto)
        nl.addWidget(self.list_need)
        lay.addWidget(self.grp_need)

        btn_edit = QPushButton("✏ Editar esta fórmula")
        btn_edit.setStyleSheet(self._btn("", "#3FB950", "#1A3A2A", bold=True))
        btn_edit.clicked.connect(self._edit_current)
        lay.addWidget(btn_edit)
        return w

    # ── Carga ─────────────────────────────────
    def _load_json(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Abrir elements JSON", "", "JSON (*.json *.txt);;Todos (*.*)")
        if not path:
            return
        try:
            sheets = load_elements(path)
        except Exception as ex:
            QMessageBox.critical(self, "Error al abrir", str(ex))
            return
        if not sheets:
            QMessageBox.warning(self, "Warning", "El JSON no contiene hojas.")
            return
        self.sheets = sheets
        self.current_sheet = 0
        self._load_sheet()
        self.lbl_status.setText(
            f"{os.path.basename(path)}  ·  {len(sheets)} hoja(s)  ·  {len(self.records)} fórmula(s)")

    def _load_sheet(self):
        if not getattr(self, "sheets", None):
            return
        if not getattr(self, "current_sheet", None):
            self.current_sheet = 0
        sheet = self.sheets[self.current_sheet]
        self.records = extract_formulas(sheet)
        self._fill_table()

    # ── Tabla ─────────────────────────────────
    def _fill_table(self):
        prev_idx = self._selected_index()
        self.table.setRowCount(0)
        for i, rec in enumerate(self.records):
            r = self.table.rowCount()
            self.table.insertRow(r)
            items = [
                rec["a1"],
                rec["name"],
                rec["temp"] or "—",
                ", ".join(sorted(rec["refs"])),
                rec["estado"],
                rec["text"].replace("\n", " "),
            ]
            for c, val in enumerate(items):
                it = QTableWidgetItem(val)
                it.setData(Qt.ItemDataRole.UserRole, i)
                if c == 0:
                    it.setForeground(QBrush(QColor("#58A6FF")))
                    it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if c == 4:
                    it.setForeground(QBrush(self._estado_color(rec["estado"])))
                self.table.setItem(r, c, it)
        if prev_idx is not None and prev_idx >= 0 and prev_idx < self.table.rowCount():
            self.table.setCurrentCell(prev_idx, 0)
        self._colorize()

    def _estado_color(self, estado):
        mapa = {
            "original": QColor("#3FB950"),
            "MODIFICADA": QColor("#D29922"),
            "REQUIERE ACTUALIZACIÓN": QColor("#F85149"),
        }
        return mapa.get(estado, QColor("#8B949E"))

    def _colorize(self):
        sel_rows = set()
        sel = self._selected_index()
        if sel is not None:
            sel_rows = set(touching_indices(self.records, sel))
        for r in range(self.table.rowCount()):
            i = self.table.item(r, 0).data(Qt.ItemDataRole.UserRole)
            rec = self.records[i]
            if i == sel:
                bg = self._sel_color
            elif rec["estado"] == "REQUIERE ACTUALIZACIÓN":
                bg = self._need_color
            elif rec["estado"] == "MODIFICADA":
                bg = self._mod_color
            elif i in sel_rows:
                bg = self._touch_color
            else:
                bg = self._gray
            for c in range(self.table.columnCount()):
                it = self.table.item(r, c)
                if it:
                    it.setBackground(QBrush(bg))

    def _selected_index(self):
        r = self.table.currentRow()
        if r < 0:
            return None
        it = self.table.item(r, 0)
        if it is None:
            return None
        return it.data(Qt.ItemDataRole.UserRole)

    def _on_cell_clicked(self, r, c):
        self._colorize()

    # ── Detalle ───────────────────────────────
    def _refresh_detail(self):
        self._colorize()
        need = [j for j, r in enumerate(self.records)
                if r["estado"] == "REQUIERE ACTUALIZACIÓN"]
        self.list_need.clear()
        if not need:
            self.list_need.addItem("(sin fórmulas pendientes de actualización)")
        else:
            for j in need:
                e = self.records[j]
                self.list_need.addItem(self._entry(e, j))

        idx = self._selected_index()
        if idx is None:
            self.lbl_sel.setText("Selecciona una fórmula…")
            self.lbl_resumen.setText("")
            self.detail_edit.clear()
            self.list_touch.clear()
            return
        rec = self.records[idx]
        self.lbl_sel.setText(
            f"{rec['name']}  ·  {rec['temp'] or 'sin temp.'}  ·  {rec['a1']}")
        fuente = ", ".join(sorted(rec["refs"])) or "—"
        self.lbl_resumen.setText(
            f"Estado: {rec['estado']}\nFuentes: {fuente}\nCant. refs: {len(rec['refs'])}")
        self.detail_edit.setPlainText(rec["text"])

        touch = touching_indices(self.records, idx)
        self.list_touch.clear()
        if not touch:
            self.list_touch.addItem("(ninguna fórmula se toca con esta)")
        else:
            for j in touch:
                e = self.records[j]
                self.list_touch.addItem(self._entry(e, j))

    def _entry(self, rec, j):
        it = QListWidgetItem(f"[{rec['a1']}]  {rec['name']}  ·  {rec['temp'] or '—'}")
        it.setData(Qt.ItemDataRole.UserRole, j)
        it.setToolTip(rec["text"][:500])
        return it

    def _goto(self, item):
        j = item.data(Qt.ItemDataRole.UserRole)
        self.table.setCurrentCell(j, 0)
        self.table.scrollToItem(self.table.item(j, 0))
        self._colorize()

    # ── Edición ───────────────────────────────
    def _edit_current(self):
        idx = self._selected_index()
        if idx is None:
            QMessageBox.information(self, "Info", "Selecciona una fórmula primero.")
            return
        self._edit_idx(idx)

    def _edit_record(self, r, c):
        """Llamado por cellDoubleClicked: la fila r == índice de registro."""
        if not (0 <= r < len(self.records)):
            return
        self._edit_idx(r)

    def _edit_idx(self, idx):
        rec = self.records[idx]
        title = f"Editar fórmula  ·  {rec['name']}  ·  {rec['a1']}"
        dlg = FormulaEditDialog(title, rec["text"], self)
        if not dlg.exec():
            return
        new_text = dlg.text().strip()
        if new_text == rec["text"]:
            return
        rec["text"] = new_text
        rec["refs"] = extract_refs(new_text)
        rec["modificada"] = True
        rec["estado"] = "MODIFICADA"
        self._recompute_integridad()
        self._fill_table()
        self.table.setCurrentCell(idx, 0)
        self._refresh_detail()
        self.lbl_status.setText(
            f"Fórmula {rec['a1']} modificada → se re-calcularon las dependencias.")

    # ── Integridad ────────────────────────────
    def _recompute_integridad(self):
        """Marca como 'REQUIERE ACTUALIZACIÓN' toda fórmula que comparte fuente
        con alguna fórmula modificada (integridad del conjunto)."""
        modificadas = [i for i, r in enumerate(self.records) if r["modificada"]]
        for r in self.records:
            if r["estado"] != "MODIFICADA":
                r["estado"] = "original"
        need_set = set()
        for i in modificadas:
            for j in touching_indices(self.records, i):
                if j != i:
                    need_set.add(j)
        for j in need_set:
            if not self.records[j]["modificada"]:
                self.records[j]["estado"] = "REQUIERE ACTUALIZACIÓN"
        self._fill_table()
        self._refresh_detail()
        if need_set:
            self.lbl_status.setText(
                f"Integridad: {len(need_set)} fórmula(s) requieren actualización.")
        else:
            self.lbl_status.setText("Integridad verificada: sin fórmulas pendientes.")

    def _reset_ediciones(self):
        for r in self.records:
            r["modificada"] = False
            r["estado"] = "original"
        self.lbl_status.setText("Ediciones restablecidas.")
        self._fill_table()
        self._refresh_detail()