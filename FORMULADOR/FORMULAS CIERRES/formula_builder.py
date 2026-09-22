"""
Constructor Visual de Fórmulas - fltFLoatTablaDinamica
Requiere: pip install PySide6 (ya incluido en venv)
"""

import sys
import json
import os
import re
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTextEdit, QLabel, QComboBox, QScrollArea,
    QFrame, QMessageBox, QSizePolicy, QSplitter, QLineEdit,
    QGroupBox, QTabWidget, QCheckBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QFileDialog, QDialog, QDialogButtonBox, QFormLayout,
    QAbstractItemView, QListWidget, QListWidgetItem
)
from PySide6.QtCore import Qt, QMimeData, QPropertyAnimation, QEasingCurve, Signal, QEvent
from PySide6.QtGui import (
    QFont, QColor, QPalette, QDrag, QPixmap, QPainter,
    QLinearGradient, QBrush, QIcon, QTextCharFormat, QSyntaxHighlighter, QTextCursor
)

# ──────────────────────────────────────────────
# PALETA DE COLORES
# ──────────────────────────────────────────────
COLORS = {
    "bg_dark":      "#0D1117",
    "bg_panel":     "#161B22",
    "bg_card":      "#1C2128",
    "bg_input":     "#21262D",
    "border":       "#30363D",
    "border_focus": "#58A6FF",
    "accent_blue":  "#58A6FF",
    "accent_green": "#3FB950",
    "accent_red":   "#F85149",
    "accent_amber": "#D29922",
    "accent_purple":"#BC8CFF",
    "text_primary": "#E6EDF3",
    "text_secondary":"#8B949E",
    "text_muted":   "#484F58",
    "op_plus":      "#1A3A2A",
    "op_minus":     "#3A1A1A",
    "op_div":       "#1A2A3A",
    "op_cond":      "#2A1A3A",
    "op_plus_border":"#3FB950",
    "op_minus_border":"#F85149",
    "op_div_border": "#58A6FF",
    "op_cond_border":"#BC8CFF",
}

OP_COLORS = {
    "+":    (COLORS["op_plus"],    COLORS["op_plus_border"],   "#3FB950"),
    "-":    (COLORS["op_minus"],   COLORS["op_minus_border"],  "#F85149"),
    "/":    (COLORS["op_div"],     COLORS["op_div_border"],    "#58A6FF"),
    "cond": (COLORS["op_cond"],    COLORS["op_cond_border"],   "#BC8CFF"),
}

OP_LABELS = {
    "+":    "＋  SUMAR",
    "-":    "－  RESTAR",
    "/":    "÷  DIVIDIR",
    "cond": "⟐  CONDICIONAL",
}

STYLESHEET = f"""
QMainWindow, QWidget {{
    background-color: {COLORS['bg_dark']};
    color: {COLORS['text_primary']};
    font-family: 'Consolas', 'Courier New', monospace;
}}
QScrollArea {{
    border: none;
    background: transparent;
}}
QScrollBar:vertical {{
    background: {COLORS['bg_panel']};
    width: 8px;
    border-radius: 4px;
}}
QScrollBar::handle:vertical {{
    background: {COLORS['border']};
    border-radius: 4px;
    min-height: 20px;
}}
QScrollBar::handle:vertical:hover {{
    background: {COLORS['text_secondary']};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0px;
}}
QScrollBar:horizontal {{
    background: {COLORS['bg_panel']};
    height: 8px;
    border-radius: 4px;
}}
QScrollBar::handle:horizontal {{
    background: {COLORS['border']};
    border-radius: 4px;
}}
QTextEdit {{
    background-color: {COLORS['bg_input']};
    color: {COLORS['text_primary']};
    border: 1px solid {COLORS['border']};
    border-radius: 6px;
    padding: 8px;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 12px;
    selection-background-color: {COLORS['accent_blue']};
}}
QTextEdit:focus {{
    border: 1px solid {COLORS['border_focus']};
}}
QLineEdit {{
    background-color: {COLORS['bg_input']};
    color: {COLORS['text_primary']};
    border: 1px solid {COLORS['border']};
    border-radius: 6px;
    padding: 6px 10px;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 12px;
}}
QLineEdit:focus {{
    border: 1px solid {COLORS['border_focus']};
}}
QComboBox {{
    background-color: {COLORS['bg_input']};
    color: {COLORS['text_primary']};
    border: 1px solid {COLORS['border']};
    border-radius: 6px;
    padding: 5px 10px;
    font-size: 12px;
    min-width: 160px;
}}
QComboBox:focus {{
    border: 1px solid {COLORS['border_focus']};
}}
QComboBox::drop-down {{
    border: none;
    padding-right: 8px;
}}
QComboBox QAbstractItemView {{
    background-color: {COLORS['bg_card']};
    color: {COLORS['text_primary']};
    border: 1px solid {COLORS['border']};
    selection-background-color: {COLORS['accent_blue']};
    border-radius: 4px;
}}
QLabel {{
    color: {COLORS['text_primary']};
    background: transparent;
}}
QTabWidget::pane {{
    border: 1px solid {COLORS['border']};
    background: {COLORS['bg_panel']};
    border-radius: 8px;
}}
QTabBar::tab {{
    background: {COLORS['bg_card']};
    color: {COLORS['text_secondary']};
    padding: 8px 20px;
    border: 1px solid {COLORS['border']};
    border-bottom: none;
    border-radius: 6px 6px 0 0;
    margin-right: 2px;
    font-size: 12px;
}}
QTabBar::tab:selected {{
    background: {COLORS['bg_panel']};
    color: {COLORS['accent_blue']};
    border-bottom: 2px solid {COLORS['accent_blue']};
}}
QGroupBox {{
    border: 1px solid {COLORS['border']};
    border-radius: 8px;
    margin-top: 12px;
    padding-top: 8px;
    font-size: 11px;
    color: {COLORS['text_secondary']};
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 2px 8px;
    color: {COLORS['text_secondary']};
}}
QCheckBox {{
    color: {COLORS['text_secondary']};
    font-size: 11px;
    spacing: 6px;
}}
QCheckBox::indicator {{
    width: 14px;
    height: 14px;
    border: 1px solid {COLORS['border']};
    border-radius: 3px;
    background: {COLORS['bg_input']};
}}
QCheckBox::indicator:checked {{
    background: {COLORS['accent_blue']};
    border-color: {COLORS['accent_blue']};
}}
QSplitter::handle {{
    background: {COLORS['border']};
    width: 2px;
}}
"""

def make_btn(text, color_bg, color_border, color_text, hover_bg=None, font_size=12, padding="8px 16px", bold=False):
    hover = hover_bg or color_border
    fw = "600" if bold else "400"
    return f"""
        QPushButton {{
            background-color: {color_bg};
            color: {color_text};
            border: 1px solid {color_border};
            border-radius: 6px;
            padding: {padding};
            font-size: {font_size}px;
            font-weight: {fw};
            font-family: 'Consolas', 'Courier New', monospace;
        }}
        QPushButton:hover {{
            background-color: {hover};
            color: {COLORS['text_primary']};
        }}
        QPushButton:pressed {{
            background-color: {color_border};
        }}
        QPushButton:disabled {{
            background-color: {COLORS['bg_card']};
            color: {COLORS['text_muted']};
            border-color: {COLORS['text_muted']};
        }}
    """


# ──────────────────────────────────────────────
# MOTOR DE CONSTRUCCIÓN DE FÓRMULAS
# ──────────────────────────────────────────────
class FormulaEngine:
    WRAPPER_START = "fltFLoatTablaDinamica(''+ ( "
    WRAPPER_END   = " ) + '',\"float\",0)"

    _placeholders: dict = {}

    @classmethod
    def set_placeholders(cls, placeholders: dict):
        cls._placeholders = dict(placeholders)

    @classmethod
    def get_placeholders(cls) -> dict:
        return dict(cls._placeholders)

    @classmethod
    def resolve(cls, text: str) -> str:
        if not text or not cls._placeholders:
            return text
        result = text
        sorted_names = sorted(cls._placeholders.keys(), key=len, reverse=True)
        for name in sorted_names:
            val = cls._placeholders[name]
            result = result.replace(name, f"( {val} )")
        return result

    @staticmethod
    def wrap(inner: str) -> str:
        return f"{FormulaEngine.WRAPPER_START}{inner}{FormulaEngine.WRAPPER_END}"

    @staticmethod
    def paren(expr: str) -> str:
        e = expr.strip()
        if e.startswith("(") and e.endswith(")"):
            return e
        return f"( {e} )"

    @staticmethod
    def combine(blocks: list) -> str:
        """
        blocks: lista de dicts con keys:
          - 'expr': str  (la expresión sin wrapper)
          - 'op':   str  ('+', '-', '/', 'cond')
          - [si cond] 'cond_expr': str, 'cond_true': str, 'cond_false': str
        Devuelve la expresión combinada lista para envolver.
        """
        if not blocks:
            return ""

        parts = []
        for i, b in enumerate(blocks):
            op   = b.get("op", "+")
            expr = b["expr"].strip()

            if op == "cond":
                # Formato: ((<condición>) > 0) ? <verdadero> : <falso>
                cond_expr  = b.get("cond_expr", expr).strip()
                cond_true  = b.get("cond_true",  expr).strip()
                cond_false = b.get("cond_false", "0").strip()
                piece = (
                    f"({FormulaEngine.paren(cond_expr)} > 0) "
                    f"? {FormulaEngine.paren(cond_true)} "
                    f": {cond_false}"
                )
            else:
                piece = FormulaEngine.paren(expr)

            if i == 0:
                parts.append(piece)
            else:
                parts.append(f" {op} {piece}")

        combined = "".join(parts)
        return FormulaEngine.paren(combined)

    @staticmethod
    def build_final(blocks: list) -> str:
        inner = FormulaEngine.combine(blocks)
        if not inner:
            return ""
        return FormulaEngine.wrap(inner)


# ──────────────────────────────────────────────
# WIDGET DE BLOQUE INDIVIDUAL
# ──────────────────────────────────────────────
class BlockWidget(QFrame):
    deleted  = Signal(object)
    moved_up = Signal(object)
    moved_dn = Signal(object)

    def __init__(self, index: int, parent=None):
        super().__init__(parent)
        self.index = index
        self._build_ui()

    def _build_ui(self):
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setStyleSheet(f"""
            BlockWidget {{
                background-color: {COLORS['bg_card']};
                border: 1px solid {COLORS['border']};
                border-radius: 8px;
                margin: 3px 0px;
            }}
        """)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(8)

        # ── Cabecera ──
        header = QHBoxLayout()
        header.setSpacing(8)

        self.lbl_num = QLabel(f"#{self.index + 1}")
        self.lbl_num.setStyleSheet(f"color:{COLORS['text_muted']}; font-size:11px;")
        self.lbl_num.setFixedWidth(26)

        self.lbl_name = QLabel(f"Bloque {self.index + 1}")
        self.lbl_name.setStyleSheet(f"color:{COLORS['text_secondary']}; font-size:11px; font-weight:600;")

        # Selector de operador (solo desde bloque 2 en adelante)
        self.op_combo = QComboBox()
        for key, label in OP_LABELS.items():
            self.op_combo.addItem(label, key)
        self.op_combo.currentIndexChanged.connect(self._on_op_changed)

        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

        btn_up = QPushButton("↑")
        btn_up.setFixedSize(26, 26)
        btn_up.setStyleSheet(make_btn("↑", COLORS['bg_input'], COLORS['border'], COLORS['text_secondary'], padding="2px"))
        btn_up.clicked.connect(lambda: self.moved_up.emit(self))

        btn_dn = QPushButton("↓")
        btn_dn.setFixedSize(26, 26)
        btn_dn.setStyleSheet(make_btn("↓", COLORS['bg_input'], COLORS['border'], COLORS['text_secondary'], padding="2px"))
        btn_dn.clicked.connect(lambda: self.moved_dn.emit(self))

        btn_del = QPushButton("✕")
        btn_del.setFixedSize(26, 26)
        btn_del.setStyleSheet(make_btn("✕", COLORS['bg_input'], COLORS['accent_red'], COLORS['accent_red'], hover_bg="#3A1A1A", padding="2px"))
        btn_del.clicked.connect(lambda: self.deleted.emit(self))

        header.addWidget(self.lbl_num)
        header.addWidget(self.lbl_name)
        header.addWidget(self.op_combo)
        header.addWidget(spacer)
        header.addWidget(btn_up)
        header.addWidget(btn_dn)
        header.addWidget(btn_del)
        root.addLayout(header)

        # ── Área de expresión ──
        self.expr_edit = QTextEdit()
        self.expr_edit.setPlaceholderText(
            "Pega aquí la expresión interna (sin el wrapper fltFLoatTablaDinamica)…\n"
            "Ejemplo: $CD_16086POLI,FECHINFPOL,dataLast$"
        )
        self.expr_edit.setMinimumHeight(70)
        self.expr_edit.setMaximumHeight(120)
        root.addWidget(self.expr_edit)

        # ── Panel condicional (oculto por defecto) ──
        self.cond_panel = QWidget()
        cond_layout = QVBoxLayout(self.cond_panel)
        cond_layout.setContentsMargins(0, 4, 0, 0)
        cond_layout.setSpacing(6)

        lbl_cond = QLabel("⟐  Configuración del condicional")
        lbl_cond.setStyleSheet(f"color:{COLORS['accent_purple']}; font-size:11px; font-weight:600;")
        cond_layout.addWidget(lbl_cond)

        lbl_c1 = QLabel("Expresión a evaluar  (¿es > 0?)")
        lbl_c1.setStyleSheet(f"color:{COLORS['text_secondary']}; font-size:10px;")
        self.cond_expr = QTextEdit()
        self.cond_expr.setPlaceholderText("Expresión que se evalúa > 0 …")
        self.cond_expr.setMinimumHeight(50)
        self.cond_expr.setMaximumHeight(80)

        lbl_c2 = QLabel("Resultado si  > 0  (verdadero)")
        lbl_c2.setStyleSheet(f"color:{COLORS['accent_green']}; font-size:10px;")
        self.cond_true = QTextEdit()
        self.cond_true.setPlaceholderText("Valor/expresión cuando la condición es verdadera…")
        self.cond_true.setMinimumHeight(50)
        self.cond_true.setMaximumHeight(80)

        lbl_c3 = QLabel("Resultado si  ≤ 0  (falso)")
        lbl_c3.setStyleSheet(f"color:{COLORS['accent_red']}; font-size:10px;")
        self.cond_false = QLineEdit()
        self.cond_false.setPlaceholderText("0")
        self.cond_false.setText("0")

        for w in [lbl_c1, self.cond_expr, lbl_c2, self.cond_true, lbl_c3, self.cond_false]:
            cond_layout.addWidget(w)

        self.cond_panel.setVisible(False)
        root.addWidget(self.cond_panel)

        # ── Previsualización del bloque ──
        self.preview_lbl = QLabel("")
        self.preview_lbl.setStyleSheet(
            f"color:{COLORS['text_muted']}; font-size:10px; "
            f"background:{COLORS['bg_input']}; padding:4px 8px; border-radius:4px;"
        )
        self.preview_lbl.setWordWrap(True)
        self.preview_lbl.setVisible(False)
        root.addWidget(self.preview_lbl)

        self._on_op_changed()
        self.expr_edit.textChanged.connect(self._update_preview)

        self._setup_paren_highlight(self.expr_edit)
        self._setup_paren_highlight(self.cond_expr)
        self._setup_paren_highlight(self.cond_true)

    def _on_op_changed(self):
        op = self.current_op()
        bg, border, txt = OP_COLORS.get(op, (COLORS['bg_card'], COLORS['border'], COLORS['text_primary']))
        self.setStyleSheet(f"""
            BlockWidget {{
                background-color: {bg};
                border: 1px solid {border};
                border-radius: 8px;
                margin: 3px 0px;
            }}
        """)
        is_cond = (op == "cond")
        self.cond_panel.setVisible(is_cond)
        if is_cond:
            self.expr_edit.setPlaceholderText(
                "Este campo es opcional para condicionales.\n"
                "Rellena los campos de abajo para configurar la condición."
            )
        else:
            self.expr_edit.setPlaceholderText(
                "Pega aquí la expresión interna (sin el wrapper)…"
            )
        self._update_preview()

    def _update_preview(self):
        data = self.get_data()
        if not data["expr"] and self.current_op() != "cond":
            self.preview_lbl.setVisible(False)
            return
        try:
            op = data["op"]
            if op == "cond":
                ce = FormulaEngine.resolve(data.get("cond_expr", "").strip())
                ct = FormulaEngine.resolve(data.get("cond_true", "").strip())
                cf = data.get("cond_false", "0").strip()
                if ce:
                    preview = (
                        f"({FormulaEngine.paren(ce)} > 0) "
                        f"? {FormulaEngine.paren(ct) if ct else '…'} : {cf}"
                    )
                    self.preview_lbl.setText(f"↳  {preview}")
                    self.preview_lbl.setVisible(True)
                else:
                    self.preview_lbl.setVisible(False)
            else:
                e = FormulaEngine.resolve(data["expr"].strip())
                if e:
                    self.preview_lbl.setText(f"↳  {FormulaEngine.paren(e)}")
                    self.preview_lbl.setVisible(True)
                else:
                    self.preview_lbl.setVisible(False)
        except Exception:
            self.preview_lbl.setVisible(False)

    # ── Highlight de paréntesis ──────────────
    def _setup_paren_highlight(self, editor: QTextEdit):
        editor.textChanged.connect(lambda: self._apply_paren_highlight(editor))
        editor.cursorPositionChanged.connect(lambda: self._apply_paren_highlight(editor))

    def _apply_paren_highlight(self, editor: QTextEdit):
        text = editor.toPlainText()
        if not text:
            editor.setExtraSelections([])
            return

        matches = {}
        stack = []
        for i, ch in enumerate(text):
            if ch == '(':
                stack.append(i)
            elif ch == ')':
                if stack:
                    j = stack.pop()
                    matches[j] = i
                    matches[i] = j

        unmatched = {i for i, ch in enumerate(text) if ch in '()' and i not in matches}

        cursor = editor.textCursor()
        pos = cursor.position()
        highlight_pair = None
        for check_pos in (pos - 1, pos):
            if 0 <= check_pos < len(text) and check_pos in matches:
                highlight_pair = (check_pos, matches[check_pos])
                break

        extra = []
        doc = editor.document()

        if highlight_pair:
            fmt_pair = QTextCharFormat()
            fmt_pair.setBackground(QColor(COLORS['accent_blue']))
            fmt_pair.setForeground(QColor(COLORS['text_primary']))
            for p in highlight_pair:
                es = QTextEdit.ExtraSelection()
                es.format = fmt_pair
                c = QTextCursor(doc)
                c.setPosition(p)
                c.movePosition(QTextCursor.MoveOperation.Right, QTextCursor.MoveMode.KeepAnchor, 1)
                es.cursor = c
                extra.append(es)

        if unmatched:
            fmt_err = QTextCharFormat()
            fmt_err.setUnderlineStyle(QTextCharFormat.UnderlineStyle.SpellCheckUnderline)
            fmt_err.setUnderlineColor(QColor(COLORS['accent_red']))
            fmt_err.setForeground(QColor(COLORS['accent_red']))
            for p in unmatched:
                if highlight_pair and p in highlight_pair:
                    continue
                es = QTextEdit.ExtraSelection()
                es.format = fmt_err
                c = QTextCursor(doc)
                c.setPosition(p)
                c.movePosition(QTextCursor.MoveOperation.Right, QTextCursor.MoveMode.KeepAnchor, 1)
                es.cursor = c
                extra.append(es)

        editor.setExtraSelections(extra)

    # ── API pública ──
    def current_op(self) -> str:
        return self.op_combo.currentData()

    def set_index(self, idx: int):
        self.index = idx
        self.lbl_num.setText(f"#{idx + 1}")
        self.lbl_name.setText(f"Bloque {idx + 1}")
        # El primer bloque no tiene operador previo visible pero lo guardamos igual
        op_visible = idx > 0
        self.op_combo.setVisible(op_visible)

    def refresh_preview(self):
        self._update_preview()

    def get_data(self) -> dict:
        d = {
            "expr": self.expr_edit.toPlainText().strip(),
            "op":   self.current_op(),
        }
        if d["op"] == "cond":
            d["cond_expr"]  = self.cond_expr.toPlainText().strip()
            d["cond_true"]  = self.cond_true.toPlainText().strip()
            d["cond_false"] = self.cond_false.text().strip() or "0"
        return d


# ──────────────────────────────────────────────
# VENTANA PRINCIPAL
# ──────────────────────────────────────────────
class FormulaBuilder(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Constructor de Fórmulas  ·  fltFLoatTablaDinamica")
        self.resize(1100, 820)
        self.setMinimumSize(800, 600)
        self.setStyleSheet(STYLESHEET)
        self.blocks: list[BlockWidget] = []
        self._init_ui()
        self._add_block()   # Empezar con un bloque

    # ── UI principal ──────────────────────────
    def _init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main = QVBoxLayout(central)
        main.setContentsMargins(0, 0, 0, 0)
        main.setSpacing(0)

        # Barra superior
        main.addWidget(self._build_topbar())

        # Splitter: izquierda = bloques, derecha = resultado
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.setHandleWidth(3)

        # Panel izquierdo
        left_widget = QWidget()
        left_widget.setStyleSheet(f"background:{COLORS['bg_dark']};")
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(16, 12, 8, 12)
        left_layout.setSpacing(10)

        # ── Sección Placeholders ──
        lbl_ph = QLabel("PLACEHOLDERS")
        lbl_ph.setStyleSheet(f"color:{COLORS['accent_purple']}; font-size:10px; font-weight:600; letter-spacing:2px;")
        left_layout.addWidget(lbl_ph)

        self.ph_scroll = QScrollArea()
        self.ph_scroll.setWidgetResizable(True)
        self.ph_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.ph_scroll.setMaximumHeight(120)

        self.ph_container = QWidget()
        self.ph_container.setStyleSheet(f"background:{COLORS['bg_dark']};")
        self.ph_layout = QVBoxLayout(self.ph_container)
        self.ph_layout.setContentsMargins(0, 0, 6, 0)
        self.ph_layout.setSpacing(3)
        self.ph_layout.addStretch()

        self.ph_scroll.setWidget(self.ph_container)
        left_layout.addWidget(self.ph_scroll)

        btn_add_ph = QPushButton("＋  Agregar Placeholder")
        btn_add_ph.setStyleSheet(make_btn(
            "＋  Agregar Placeholder",
            COLORS['bg_card'], COLORS['accent_purple'], COLORS['accent_purple'],
            hover_bg="#1A1A3A", padding="4px 12px", font_size=11
        ))
        btn_add_ph.clicked.connect(lambda: self._add_placeholder())
        left_layout.addWidget(btn_add_ph)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet(f"color:{COLORS['border']};")
        sep.setFixedHeight(1)
        left_layout.addWidget(sep)

        # ── Sección Bloques ──
        lbl_bloques = QLabel("BLOQUES DE FÓRMULA")
        lbl_bloques.setStyleSheet(f"color:{COLORS['text_muted']}; font-size:10px; font-weight:600; letter-spacing:2px;")
        left_layout.addWidget(lbl_bloques)

        # Scroll para los bloques
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        self.blocks_container = QWidget()
        self.blocks_container.setStyleSheet(f"background:{COLORS['bg_dark']};")
        self.blocks_layout = QVBoxLayout(self.blocks_container)
        self.blocks_layout.setContentsMargins(0, 0, 6, 0)
        self.blocks_layout.setSpacing(4)
        self.blocks_layout.addStretch()

        self.scroll_area.setWidget(self.blocks_container)
        left_layout.addWidget(self.scroll_area)

        # Botones de acción inferiores
        left_layout.addWidget(self._build_action_bar())

        # Panel derecho
        right_widget = QWidget()
        right_widget.setStyleSheet(f"background:{COLORS['bg_panel']};")
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(12, 12, 16, 12)
        right_layout.setSpacing(10)

        lbl_res = QLabel("RESULTADO GENERADO")
        lbl_res.setStyleSheet(f"color:{COLORS['text_muted']}; font-size:10px; font-weight:600; letter-spacing:2px;")
        right_layout.addWidget(lbl_res)

        right_layout.addWidget(self._build_result_panel())

        self.splitter.addWidget(left_widget)
        self.splitter.addWidget(right_widget)
        self.splitter.setSizes([580, 520])

        main.addWidget(self.splitter, 1)

    def _build_topbar(self) -> QWidget:
        bar = QFrame()
        bar.setFixedHeight(52)
        bar.setStyleSheet(f"""
            QFrame {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {COLORS['bg_panel']}, stop:1 {COLORS['bg_dark']});
                border-bottom: 1px solid {COLORS['border']};
            }}
        """)
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(20, 0, 20, 0)

        icon_lbl = QLabel("⟨ƒ⟩")
        icon_lbl.setStyleSheet(f"color:{COLORS['accent_blue']}; font-size:20px; font-weight:700;")

        title = QLabel("Formula Builder")
        title.setStyleSheet(f"color:{COLORS['text_primary']}; font-size:16px; font-weight:700;")

        sub = QLabel("fltFLoatTablaDinamica · Constructor de expresiones")
        sub.setStyleSheet(f"color:{COLORS['text_muted']}; font-size:11px;")

        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

        lbl_ver = QLabel("v1.0")
        lbl_ver.setStyleSheet(f"color:{COLORS['text_muted']}; font-size:10px;")

        layout.addWidget(icon_lbl)
        layout.addSpacing(8)
        layout.addWidget(title)
        layout.addSpacing(12)
        layout.addWidget(sub)
        layout.addWidget(spacer)
        layout.addWidget(lbl_ver)

        return bar

    def _build_action_bar(self) -> QWidget:
        bar = QWidget()
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(0, 4, 0, 0)
        layout.setSpacing(8)

        btn_add = QPushButton("＋  Agregar Bloque")
        btn_add.setStyleSheet(make_btn(
            "＋  Agregar Bloque",
            COLORS['bg_card'], COLORS['accent_green'], COLORS['accent_green'],
            hover_bg="#1A3A2A", padding="8px 18px", bold=True
        ))
        btn_add.clicked.connect(self._add_block)

        btn_clear = QPushButton("⊘  Limpiar Todo")
        btn_clear.setStyleSheet(make_btn(
            "⊘  Limpiar Todo",
            COLORS['bg_card'], COLORS['border'], COLORS['text_secondary'],
            hover_bg=COLORS['bg_input'], padding="8px 14px"
        ))
        btn_clear.clicked.connect(self._clear_all)

        btn_build = QPushButton("▶  Construir Fórmula")
        btn_build.setStyleSheet(make_btn(
            "▶  Construir Fórmula",
            "#0D2137", COLORS['accent_blue'], COLORS['accent_blue'],
            hover_bg="#1A3A5A", padding="8px 18px", bold=True
        ))
        btn_build.clicked.connect(self._build_formula)

        layout.addWidget(btn_add)
        layout.addWidget(btn_clear)
        layout.addStretch()
        layout.addWidget(btn_build)

        return bar

    def _build_result_panel(self) -> QWidget:
        tabs = QTabWidget()

        # ── Tab 1: Fórmula completa
        tab_full = QWidget()
        t1 = QVBoxLayout(tab_full)
        t1.setContentsMargins(8, 8, 8, 8)
        t1.setSpacing(8)

        self.result_full = QTextEdit()
        self.result_full.setReadOnly(True)
        self.result_full.setPlaceholderText("Presiona  ▶ Construir Fórmula  para generar el resultado…")
        self.result_full.setStyleSheet(f"""
            QTextEdit {{
                background: {COLORS['bg_input']};
                color: {COLORS['accent_green']};
                border: 1px solid {COLORS['border']};
                border-radius: 6px;
                padding: 10px;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 12px;
            }}
        """)
        t1.addWidget(self.result_full)

        row = QHBoxLayout()
        row.setSpacing(8)
        btn_copy_full = QPushButton("⎘  Copiar")
        btn_copy_full.setStyleSheet(make_btn("⎘  Copiar", COLORS['bg_card'], COLORS['accent_blue'], COLORS['accent_blue'], padding="6px 14px"))
        btn_copy_full.clicked.connect(lambda: self._copy(self.result_full.toPlainText()))

        self.lbl_char = QLabel("0 caracteres")
        self.lbl_char.setStyleSheet(f"color:{COLORS['text_muted']}; font-size:10px;")

        row.addWidget(btn_copy_full)
        row.addStretch()
        row.addWidget(self.lbl_char)
        t1.addLayout(row)

        # ── Tab 2: Solo expresión interna
        tab_inner = QWidget()
        t2 = QVBoxLayout(tab_inner)
        t2.setContentsMargins(8, 8, 8, 8)
        t2.setSpacing(8)

        self.result_inner = QTextEdit()
        self.result_inner.setReadOnly(True)
        self.result_inner.setPlaceholderText("La expresión interna aparecerá aquí…")
        self.result_inner.setStyleSheet(f"""
            QTextEdit {{
                background: {COLORS['bg_input']};
                color: {COLORS['accent_amber']};
                border: 1px solid {COLORS['border']};
                border-radius: 6px;
                padding: 10px;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 12px;
            }}
        """)
        t2.addWidget(self.result_inner)

        btn_copy_inner = QPushButton("⎘  Copiar expresión interna")
        btn_copy_inner.setStyleSheet(make_btn("⎘  Copiar", COLORS['bg_card'], COLORS['accent_amber'], COLORS['accent_amber'], padding="6px 14px"))
        btn_copy_inner.clicked.connect(lambda: self._copy(self.result_inner.toPlainText()))
        t2.addWidget(btn_copy_inner)

        # ── Tab 3: Historial
        tab_hist = QWidget()
        t3 = QVBoxLayout(tab_hist)
        t3.setContentsMargins(8, 8, 8, 8)
        t3.setSpacing(8)

        self.hist_edit = QTextEdit()
        self.hist_edit.setReadOnly(True)
        self.hist_edit.setPlaceholderText("El historial de fórmulas construidas aparecerá aquí…")
        t3.addWidget(self.hist_edit)

        btn_clear_hist = QPushButton("Limpiar historial")
        btn_clear_hist.setStyleSheet(make_btn("Limpiar", COLORS['bg_card'], COLORS['border'], COLORS['text_secondary'], padding="6px 12px"))
        btn_clear_hist.clicked.connect(lambda: self.hist_edit.clear())
        t3.addWidget(btn_clear_hist)

        tabs.addTab(tab_full,  "Fórmula Completa")
        tabs.addTab(tab_inner, "Expresión Interna")
        tabs.addTab(tab_hist,  "Historial")
        self.library_tab = FormulaLibraryTab(self)
        tabs.addTab(self.library_tab, "📚 Librería de Fórmulas")
        self.excel_tab = ExcelExprTab(self)
        tabs.addTab(self.excel_tab, "∑  Unir celdas (tipo Excel)")

        return tabs

    # ── Gestión de bloques ────────────────────
    def _add_block(self):
        idx = len(self.blocks)
        block = BlockWidget(idx)
        block.deleted.connect(self._remove_block)
        block.moved_up.connect(self._move_up)
        block.moved_dn.connect(self._move_dn)

        self.blocks.append(block)
        # Insertar antes del stretch
        self.blocks_layout.insertWidget(self.blocks_layout.count() - 1, block)
        block.set_index(idx)

        # Scroll al fondo
        self.scroll_area.verticalScrollBar().setValue(
            self.scroll_area.verticalScrollBar().maximum()
        )

    def _remove_block(self, block: BlockWidget):
        if len(self.blocks) == 1:
            QMessageBox.information(self, "Info", "Debe haber al menos un bloque.")
            return
        self.blocks.remove(block)
        self.blocks_layout.removeWidget(block)
        block.deleteLater()
        self._reindex()

    def _move_up(self, block: BlockWidget):
        idx = self.blocks.index(block)
        if idx == 0:
            return
        self.blocks[idx], self.blocks[idx - 1] = self.blocks[idx - 1], self.blocks[idx]
        self._rebuild_layout()

    def _move_dn(self, block: BlockWidget):
        idx = self.blocks.index(block)
        if idx == len(self.blocks) - 1:
            return
        self.blocks[idx], self.blocks[idx + 1] = self.blocks[idx + 1], self.blocks[idx]
        self._rebuild_layout()

    def _rebuild_layout(self):
        for b in self.blocks:
            self.blocks_layout.removeWidget(b)
        for b in self.blocks:
            self.blocks_layout.insertWidget(self.blocks_layout.count() - 1, b)
        self._reindex()

    def _reindex(self):
        for i, b in enumerate(self.blocks):
            b.set_index(i)

    def _clear_all(self):
        reply = QMessageBox.question(
            self, "Confirmar", "¿Limpiar todos los bloques y placeholders?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            # Limpiar placeholders
            for i in reversed(range(self.ph_layout.count())):
                item = self.ph_layout.itemAt(i)
                if item and item.widget():
                    item.widget().deleteLater()
            FormulaEngine.set_placeholders({})

            # Limpiar bloques
            for b in self.blocks[:]:
                self.blocks_layout.removeWidget(b)
                b.deleteLater()
            self.blocks.clear()
            self.result_full.clear()
            self.result_inner.clear()
            self.lbl_char.setText("0 caracteres")
            self._add_block()

    # ── Gestión de placeholders ──────────────
    def _add_placeholder(self, name="", expr=""):
        row = QWidget()
        row.setStyleSheet(f"background:transparent;")
        h = QHBoxLayout(row)
        h.setContentsMargins(0, 1, 0, 1)
        h.setSpacing(4)

        name_edit = QLineEdit(name)
        name_edit.setPlaceholderText("Nombre del placeholder…")
        name_edit.setMinimumWidth(100)
        name_edit.setStyleSheet(f"font-size:11px; padding:3px 6px;")

        expr_edit = QLineEdit(expr)
        expr_edit.setPlaceholderText("Expresión (ej: $CD_16086POLI,FECHINFPOL,dataLast$)…")
        expr_edit.setStyleSheet(f"font-size:11px; padding:3px 6px;")

        btn_del = QPushButton("✕")
        btn_del.setFixedSize(22, 22)
        btn_del.setStyleSheet(make_btn("✕", COLORS['bg_input'], COLORS['accent_red'], COLORS['accent_red'],
                                        hover_bg="#3A1A1A", padding="1px", font_size=10))

        def on_delete():
            self.ph_layout.removeWidget(row)
            row.deleteLater()
            self._sync_placeholders()

        def on_change(*_):
            self._sync_placeholders()

        btn_del.clicked.connect(on_delete)
        name_edit.textChanged.connect(on_change)
        expr_edit.textChanged.connect(on_change)

        h.addWidget(name_edit)
        h.addWidget(expr_edit)
        h.addWidget(btn_del)

        # Insertar antes del stretch
        self.ph_layout.insertWidget(self.ph_layout.count() - 1, row)
        self._sync_placeholders()

    def _sync_placeholders(self):
        ph = {}
        for i in range(self.ph_layout.count()):
            item = self.ph_layout.itemAt(i)
            if item and item.widget():
                w = item.widget()
                children = w.findChildren(QLineEdit)
                if len(children) >= 2:
                    n = children[0].text().strip()
                    e = children[1].text().strip()
                    if n:
                        ph[n] = e
        FormulaEngine.set_placeholders(ph)
        # Refrescar previsualizaciones de todos los bloques
        for b in self.blocks:
            b.refresh_preview()

    # ── Construcción ─────────────────────────
    def _build_formula(self):
        data_list = [b.get_data() for b in self.blocks]

        # Resolver placeholders en cada expresión
        for d in data_list:
            d["expr"] = FormulaEngine.resolve(d["expr"])
            if "cond_expr" in d:
                d["cond_expr"] = FormulaEngine.resolve(d["cond_expr"])
            if "cond_true" in d:
                d["cond_true"] = FormulaEngine.resolve(d["cond_true"])

        # Validación básica
        errors = []
        for i, d in enumerate(data_list):
            op = d["op"]
            if op == "cond":
                if not d.get("cond_expr") and not d.get("cond_true"):
                    errors.append(f"Bloque #{i+1}: falta expresión condicional.")
            else:
                if not d["expr"]:
                    errors.append(f"Bloque #{i+1}: la expresión está vacía.")

        if errors:
            QMessageBox.warning(self, "Bloques incompletos", "\n".join(errors))
            return

        try:
            inner   = FormulaEngine.combine(data_list)
            formula = FormulaEngine.wrap(inner)

            self.result_full.setPlainText(formula)
            self.result_inner.setPlainText(inner)
            self.lbl_char.setText(f"{len(formula):,} caracteres")

            # Historial
            separator = "─" * 60
            self.hist_edit.append(f"\n{separator}")
            self.hist_edit.append(f"[{len(self.blocks)} bloque(s)]")
            self.hist_edit.append(formula)

        except Exception as ex:
            QMessageBox.critical(self, "Error al construir", str(ex))

    # ── Utilidades ───────────────────────────
    def _copy(self, text: str):
        if text:
            QApplication.clipboard().setText(text)
            # Feedback visual breve: título cambia
            original = self.windowTitle()
            self.setWindowTitle("✓ Copiado al portapapeles!")
            from PySide6.QtCore import QTimer
            QTimer.singleShot(1800, lambda: self.setWindowTitle(original))


# ──────────────────────────────────────────────
# LIBRERÍA DE FÓRMULAS (vista tipo Excel)
# ──────────────────────────────────────────────
TEMPORALITIES = ["DIA", "SEMANA", "MES", "AÑO"]
MIN_MATCH = 15
LIB_DEFAULT_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)), "libreria_formulas.json")


def _norm(s):
    return re.sub(r"\s+", " ", str(s)).strip()


_PREFIX_RE = re.compile(r"^\s*fltfloattabladinamica\s*\(\s*''\s*\+\s*\(", re.I)
_SUFFIX_RE = re.compile(r"\s*\)\s*\+\s*''\s*,\s*[\"']float[\"']\s*,\s*\d+\s*\)\s*$", re.I)


def extract_inner(text):
    """Extrae la expresión interna de fltFLoatTablaDinamica(''+ ( X ) + '',"float",N).
    Si no reconoce el wrapper, devuelve el texto normalizado tal cual."""
    t = _norm(text)
    m = _PREFIX_RE.match(t)
    if not m:
        return t
    s = _SUFFIX_RE.search(t)
    if not s or s.start() < m.end():
        return t
    return t[m.end():s.start()]


def strip_outer(expr):
    """Quita el paréntesis externo equilibrado que envuelve toda la expresión."""
    e = expr.strip()
    while len(e) >= 2 and e[0] == "(" and e[-1] == ")":
        depth = 0
        wrapped = True
        for i, ch in enumerate(e):
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0 and i != len(e) - 1:
                    wrapped = False
                    break
        if not wrapped or depth != 0:
            break
        e = e[1:-1].strip()
    return e


def split_top_level(expr):
    """Divide por operadores + - * / a nivel 0 de paréntesis. Devuelve [[op, texto], ...]"""
    parts = []
    depth = 0
    cur = []
    op = "+"
    for ch in expr:
        if ch == "(":
            depth += 1
            cur.append(ch)
        elif ch == ")":
            depth -= 1
            cur.append(ch)
        elif ch in "+-*/" and depth == 0:
            text = "".join(cur).strip()
            if text:
                parts.append([op, text])
            op = ch
            cur = []
        else:
            cur.append(ch)
    text = "".join(cur).strip()
    if text:
        parts.append([op, text])
    return parts


def decompose(expr, max_level=2, _budget=None):
    """Descompone una fórmula en términos jerárquicos: [(nivel, operador, texto), ...].
    Limita profundidad y cantidad para no explotar con fórmulas gigantes."""
    if _budget is None:
        _budget = [60]
    items = []
    e = strip_outer(expr)
    for op, text in split_top_level(e):
        if _budget[0] <= 0:
            break
        _budget[0] -= 1
        items.append((0, op, text))
        if len(text) > 200 and max_level > 1:
            for lvl, op2, t2 in decompose(text, max_level - 1, _budget):
                items.append((lvl + 1, op2, t2))
    return items


def term_kind(term):
    if "* 100" in term:
        return "Escala % (×100)"
    if "?" in term and ":" in term:
        return "Condicional"
    if " && " in term or " || " in term:
        return "Lógica"
    if "/" in term:
        return "División"
    if " * " in term:
        return "Multiplicación"
    return "Suma/Resta"


def describe_inner(inner):
    """Bosquejo simple en palabras: cuántas sumas, restas, condicionales, etc."""
    desc = []
    plus = inner.count("+")
    minus = inner.count("-")
    div = inner.count("/")
    mul = inner.count("*")
    cond = inner.count("?") + inner.count("&&") + inner.count("||")
    if plus:   desc.append(f"{plus} suma(s)")
    if minus:  desc.append(f"{minus} resta(s)")
    if div:    desc.append(f"{div} división(es)")
    if mul:    desc.append(f"{mul} multiplicación(es)")
    if cond:   desc.append(f"{cond} condición(es)")
    if "* 100" in inner:
        desc.append("escala a % (×100)")
    return desc or ["(operación simple)"]


def _balanced(s):
    depth = 0
    for ch in s:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth < 0:
                return False
    return depth == 0


def generate_formula(expr0, repo, dec="2"):
    """Construye la fórmula fltFLoatTablaDinamica a partir de una expresión tipo Excel.
    expr0: texto del usuario (ya sin '=' inicial). repo: lista de dicts {name, temp, inner}.
    Acepta referencias por nombre (Nombre:TEMP) y por dirección A1 (B1, C3…, B..E columna).
    Devuelve (formula, error). error es None si todo bien, si no un mensaje."""
    expr0 = _norm(expr0)
    if not expr0:
        return None, "Escribe una expresión primero."

    known = {}
    addr_map = {}
    for e in repo:
        name = e['name'].strip()
        if not name:
            continue
        key = f"{name}:{e['temp']}"
        if key not in known:
            known[key] = e["inner"]
        if 1 <= e['col'] <= 4:
            addr = f"{chr(ord('B') + e['col'] - 1)}{e['row'] + 1}"
            addr_map[addr] = key

    addr_pat = re.compile(r"\b([B-E])(\d+)\b", re.I)

    def _addr_repl(m):
        a = f"{m.group(1).upper()}{int(m.group(2))}"
        return addr_map.get(a, m.group(0))

    expr1 = addr_pat.sub(_addr_repl, expr0)

    # Detección de referencias sin resolver: enmascara lo que SÍ coincide
    masked = list(expr1)
    covered = [False] * len(expr1)
    low = expr1.lower()
    for key in sorted(known, key=len, reverse=True):
        klow = key.lower()
        it = low.find(klow)
        while it != -1:
            for p in range(it, it + len(key)):
                covered[p] = True
            it = low.find(klow, it + len(key))
    for i, cov in enumerate(covered):
        if cov:
            masked[i] = "X"
    masked_str = "".join(masked)

    missing = []
    pat = re.compile(r"([^:()+\-*/=]+):(DIA|SEMANA|MES|AÑO)\b", re.I)
    missing += sorted({f"{m.group(1).strip()}:{m.group(2).upper()}" for m in pat.finditer(masked_str)})
    for m in addr_pat.finditer(masked_str):
        a = f"{m.group(1).upper()}{int(m.group(2))}"
        if a not in addr_map:
            missing.append(f"celda {a} (no existe)")
    if missing:
        return None, ("Referencias sin resolver:\n  " + "\n  ".join(sorted(set(missing))) +
                      "\n\nRevisa la dirección o haz clic sobre las celdas de la librería.")

    expr = expr1
    for key in sorted(known, key=len, reverse=True):
        expr = re.sub(re.escape(key), f"( {known[key]} )", expr, flags=re.IGNORECASE)

    if not _balanced(expr):
        return None, "La expresión tiene paréntesis desbalanceados."

    formula = f"fltFLoatTablaDinamica(''+ ( {expr} ) + '',\"float\",{dec})"
    return formula, None


class FormulaEditorDialog(QDialog):
    def __init__(self, title, initial="", parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(900, 560)
        lay = QVBoxLayout(self)
        lab = QLabel("Pega la fórmula aquí y pulsa Aceptar para guardarla en la celda.")
        lab.setStyleSheet(f"color:{COLORS['text_secondary']}; font-size:11px;")
        self.edit = QTextEdit()
        self.edit.setPlainText(initial)
        self.edit.setStyleSheet(f"""
            QTextEdit {{
                background: {COLORS['bg_input']};
                color: {COLORS['accent_green']};
                border: 1px solid {COLORS['border']};
                border-radius: 6px;
                padding: 10px;
                font-family: 'Menlo', 'Consolas', monospace;
                font-size: 12px;
            }}
        """)
        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        lay.addWidget(lab)
        lay.addWidget(self.edit)
        lay.addWidget(btns)

    def text(self):
        return self.edit.toPlainText()


class ExcelExprTab(QWidget):
    """Constructor tipo Excel: =celda+celda-celda usando los nombres de la librería."""
    def __init__(self, builder, parent=None):
        super().__init__(parent)
        self.builder = builder
        self.insert_mode = False
        self._build_ui()

    # ── UI ────────────────────────────────────
    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(8)

        hint = QLabel(
            "Escribe una expresión tipo Excel referenciando celdas de la librería con  <Nombre>:<TEMPORALIDAD>.\n"
            "Método corto: escribe '=' y haz clic sobre una celda en 📚 Librería — se inserta sola.\n"
            "Ejemplo:\n"
            "  =INVENTARIO INICIAL ARICHE:DIA + ARICHE - OBTENIDO:DIA - ARICHE DESPACHADO:DIA + [7~dataLast~C_3176]\n"
            "Cada referencia se sustituye por su expresión interna, y el conjunto se envuelve en "
            "fltFLoatTablaDinamica(''+ (  ) + '',\"float\",N)."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet(f"color:{COLORS['text_secondary']}; font-size:11px;")
        root.addWidget(hint)

        self.input = QTextEdit()
        self.input.setPlaceholderText(
            "=INVENTARIO INICIAL ARICHE:DIA + ARICHE - OBTENIDO:DIA - ARICHE DESPACHADO:DIA")
        self.input.setMinimumHeight(90)
        self.input.textChanged.connect(self._on_input_changed)
        root.addWidget(self.input)

        row = QHBoxLayout()
        row.setSpacing(6)
        self.lbl_mode = QLabel("Insertar celda: INACTIVO")
        self.lbl_mode.setStyleSheet(f"color:{COLORS['text_muted']}; font-size:10px;")
        lbl = QLabel("Insertar celda:")
        lbl.setStyleSheet(f"color:{COLORS['text_secondary']}; font-size:11px;")
        self.cell_combo = QComboBox()
        self.cell_combo.setMinimumWidth(240)
        btn_refresh = QPushButton("⟳")
        btn_refresh.setFixedSize(26, 26)
        btn_refresh.setStyleSheet(make_btn("⟳", COLORS['bg_input'], COLORS['border'], COLORS['text_secondary'], padding="1px"))
        btn_refresh.setToolTip("Actualizar lista de celdas de la librería")
        btn_refresh.clicked.connect(self._reload_cells)
        btn_ins = QPushButton("＋ insertar")
        btn_ins.setStyleSheet(make_btn("＋ insertar", COLORS['bg_card'], COLORS['accent_purple'], COLORS['accent_purple'], padding="5px 10px", font_size=10))
        btn_ins.clicked.connect(self._insert_ref)
        lbl_n = QLabel("float:")
        lbl_n.setStyleSheet(f"color:{COLORS['text_secondary']}; font-size:11px;")
        self.dec_edit = QLineEdit("2")
        self.dec_edit.setFixedWidth(40)
        self.dec_edit.setToolTip('Parámetro "float",N del wrapper (0 o 2 normalmente)')
        btn_gen = QPushButton("▶  Generar fórmula")
        btn_gen.setStyleSheet(make_btn("▶  Generar fórmula", "#0D2137", COLORS['accent_blue'], COLORS['accent_blue'], hover_bg="#1A3A5A", padding="7px 16px", bold=True))
        btn_gen.clicked.connect(self._build)

        row.addWidget(self.lbl_mode)
        row.addWidget(lbl)
        row.addWidget(self.cell_combo)
        row.addWidget(btn_refresh)
        row.addWidget(btn_ins)
        row.addWidget(lbl_n)
        row.addWidget(self.dec_edit)
        row.addStretch()
        row.addWidget(btn_gen)
        root.addLayout(row)

        self.output = QTextEdit()
        self.output.setReadOnly(True)
        self.output.setPlaceholderText("La fórmula generada aparecerá aquí…")
        self.output.setStyleSheet(f"""
            QTextEdit {{
                background: {COLORS['bg_input']};
                color: {COLORS['accent_green']};
                border: 1px solid {COLORS['border']};
                border-radius: 6px;
                padding: 10px;
                font-family: 'Menlo', 'Consolas', monospace;
                font-size: 12px;
            }}
        """)
        root.addWidget(self.output, 1)

        row2 = QHBoxLayout()
        row2.setSpacing(8)
        btn_copy = QPushButton("⎘ Copiar")
        btn_copy.setStyleSheet(make_btn("⎘ Copiar", COLORS['bg_card'], COLORS['accent_blue'], COLORS['accent_blue'], padding="6px 14px"))
        btn_copy.clicked.connect(lambda: self._copy(self.output.toPlainText()))
        self.lbl_char = QLabel("0 caracteres")
        self.lbl_char.setStyleSheet(f"color:{COLORS['text_muted']}; font-size:10px;")
        btn_save_lib = QPushButton("💾 Guardar en la celda seleccionada de la Librería")
        btn_save_lib.setStyleSheet(make_btn("💾 Guardar en la celda seleccionada de la Librería", COLORS['bg_card'], COLORS['accent_green'], COLORS['accent_green'], padding="6px 12px"))
        btn_save_lib.clicked.connect(self._save_to_library)
        row2.addWidget(btn_copy)
        row2.addWidget(self.lbl_char)
        row2.addStretch()
        row2.addWidget(btn_save_lib)
        root.addLayout(row2)

        self._reload_cells()

    # ── Acciones ──────────────────────────────
    def set_insert_mode(self, on):
        self.insert_mode = bool(on)
        if self.lbl_mode:
            if on:
                self.lbl_mode.setText("Insertar celda: ACTIVO — haz clic en 📚 Librería")
                self.lbl_mode.setStyleSheet(f"color:{COLORS['accent_purple']}; font-size:10px; font-weight:600;")
            else:
                self.lbl_mode.setText("Insertar celda: INACTIVO")
                self.lbl_mode.setStyleSheet(f"color:{COLORS['text_muted']}; font-size:10px;")
        lib = self.builder.library_tab
        if lib.b_pick.isChecked() != on:
            lib.b_pick.setChecked(on)

    def _on_input_changed(self):
        if self.input.toPlainText().lstrip().startswith("=") and not self.insert_mode:
            self.set_insert_mode(True)

    def insert_cell_ref(self, name, temp):
        key = f"{name.strip()}:{temp}"
        tc = self.input.textCursor()
        tc.insertText(key)
        self.input.setTextCursor(tc)
        self.input.setFocus()

    def _reload_cells(self):
        self.cell_combo.clear()
        repo = self.builder.library_tab._repo()
        for e in repo:
            key = f"{e['name'].strip()}:{e['temp']}"
            self.cell_combo.addItem(f"{e['name'].strip()}  ·  {e['temp']}", key)
        if self.cell_combo.count() == 0:
            self.cell_combo.addItem("(sin celdas en la librería)", None)

    def _insert_ref(self):
        key = self.cell_combo.currentData()
        if key:
            self.input.insertPlainText(key)

    def _build(self):
        raw = self.input.toPlainText()
        expr0 = _norm(raw)
        if expr0.startswith("="):
            expr0 = expr0[1:]
        elif expr0.startswith("+="):
            expr0 = expr0[2:]
        repo = self.builder.library_tab._repo()
        formula, err = generate_formula(expr0, repo, self.dec_edit.text().strip() or "2")
        if err:
            QMessageBox.warning(self, "No se pudo construir", err)
            return
        self.output.setPlainText(formula)
        self.lbl_char.setText(f"{len(formula):,} caracteres")

    def _save_to_library(self):
        formula = self.output.toPlainText()
        if not formula:
            QMessageBox.information(self, "Nada que guardar", "Primero genera la fórmula.")
            return
        tab = self.builder.library_tab
        r = tab.table.currentRow()
        c = tab.table.currentColumn()
        if r < 0 or c <= 0:
            QMessageBox.information(self, "Selecciona una celda",
                                    "En la pestaña 📚 Librería, selecciona la celda de fórmula (DIA/SEMANA/MES/AÑO) donde quieres guardarla y pulsa de nuevo.")
            return
        item = tab.table.item(r, c)
        if item is None:
            item = QTableWidgetItem("")
            tab.table.setItem(r, c, item)
        item.setText(formula)
        tab.table.resizeRowToContents(r)
        tab._refresh_detail()
        tab.lbl_status.setText("Fórmula guardada ✓ en la celda seleccionada")

    def _copy(self, text):
        if text:
            QApplication.clipboard().setText(text)


class ReplaceDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Reemplazo masivo en la librería")
        self.setMinimumWidth(540)
        lay = QVBoxLayout(self)
        form = QFormLayout()
        self.find_edit = QLineEdit()
        self.find_edit.setPlaceholderText("Texto a buscar (ej: dataLast)")
        self.rep_edit = QLineEdit()
        self.rep_edit.setPlaceholderText("Texto de reemplazo")
        form.addRow("Buscar:", self.find_edit)
        form.addRow("Reemplazar:", self.rep_edit)
        lay.addLayout(form)
        note = QLabel("Solo afecta las celdas de fórmulas (DIA / SEMANA / MES / AÑO).")
        note.setStyleSheet(f"color:{COLORS['text_muted']}; font-size:10px;")
        lay.addWidget(note)
        btn = QPushButton("Reemplazar todo")
        btn.setStyleSheet(make_btn("Reemplazar todo", COLORS['bg_card'], COLORS['accent_amber'], COLORS['accent_amber'], padding="7px 14px"))
        btn.clicked.connect(self.accept)
        lay.addWidget(btn)

    def values(self):
        return self.find_edit.text(), self.rep_edit.text()


class FormulaLibraryTab(QWidget):
    def __init__(self, builder=None, parent=None):
        super().__init__(parent)
        self.builder = builder
        self._term_hits = []
        self.fx_active = False
        self.fx_target = None
        self._build_ui()
        if os.path.exists(LIB_DEFAULT_JSON):
            try:
                with open(LIB_DEFAULT_JSON, "r", encoding="utf-8") as f:
                    self._load_data(json.load(f))
            except Exception:
                pass

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
        self.splitter.setSizes([560, 430])
        root.addWidget(self.splitter, 1)

    def _build_toolbar(self):
        bar = QWidget()
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)

        b_add = QPushButton("＋ Indicador")
        b_add.setStyleSheet(make_btn("＋ Indicador", COLORS['bg_card'], COLORS['accent_green'], COLORS['accent_green'], hover_bg="#1A3A2A", padding="5px 12px", font_size=11, bold=True))
        b_add.setToolTip("Agrega una fila nueva (un indicador con sus columnas DIA/SEMANA/MES/AÑO)")
        b_add.clicked.connect(self._add_row)

        self.b_pick = QPushButton("🖱 Insertar al hacer clic")
        self.b_pick.setCheckable(True)
        self.b_pick.setToolTip(
            "Modo Excel: cuando está activo (o escribes '=' en ∑ Unir celdas), "
            "hacer clic en una celda de fórmula la agrega a tu expresión.")
        self.b_pick.toggled.connect(self._on_pick_toggled)

        b_del = QPushButton("✂ Eliminar fila")
        b_del.setStyleSheet(make_btn("✂ Eliminar fila", COLORS['bg_card'], COLORS['accent_red'], COLORS['accent_red'], hover_bg="#3A1A1A", padding="5px 12px", font_size=11))
        b_del.clicked.connect(self._del_row)

        b_open = QPushButton("📂 Abrir JSON")
        b_open.setStyleSheet(make_btn("📂 Abrir JSON", COLORS['bg_card'], COLORS['accent_purple'], COLORS['accent_purple'], padding="5px 12px", font_size=11))
        b_open.clicked.connect(self._load_json)

        b_save = QPushButton("💾 Guardar JSON")
        b_save.setStyleSheet(make_btn("💾 Guardar JSON", COLORS['bg_card'], COLORS['accent_purple'], COLORS['accent_purple'], padding="5px 12px", font_size=11))
        b_save.clicked.connect(self._save_json)

        b_rep = QPushButton("🔄 Reemplazo masivo")
        b_rep.setStyleSheet(make_btn("🔄 Reemplazo masivo", COLORS['bg_card'], COLORS['accent_amber'], COLORS['accent_amber'], padding="5px 12px", font_size=11))
        b_rep.clicked.connect(self._do_replace)

        b_clr = QPushButton("🧹 Limpiar")
        b_clr.setStyleSheet(make_btn("🧹 Limpiar", COLORS['bg_card'], COLORS['border'], COLORS['text_secondary'], padding="5px 12px", font_size=11))
        b_clr.clicked.connect(self._clear_all)

        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

        self.lbl_status = QLabel("")
        self.lbl_status.setStyleSheet(f"color:{COLORS['text_secondary']}; font-size:10px;")

        lay.addWidget(b_add)
        lay.addWidget(self.b_pick)
        lay.addWidget(b_del)
        lay.addWidget(b_open)
        lay.addWidget(b_save)
        lay.addWidget(b_rep)
        lay.addWidget(b_clr)
        lay.addWidget(spacer)
        lay.addWidget(self.lbl_status)
        return bar

    def _build_grid(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(4)

        # ── Barra fx (constructor tipo Excel dentro de la librería) ──
        fx = QHBoxLayout()
        fx.setSpacing(6)
        fx_lbl = QLabel("fx")
        fx_lbl.setStyleSheet(f"color:{COLORS['accent_green']}; font-size:14px; font-weight:700;")
        fx_lbl.setFixedWidth(24)
        self.fx_edit = QLineEdit()
        self.fx_edit.setPlaceholderText("Escribe = en una celda y haz clic sobre otras…")
        self.fx_edit.setEnabled(False)
        self.fx_edit.setStyleSheet(f"""
            QLineEdit {{
                background: {COLORS['bg_input']};
                color: {COLORS['accent_green']};
                border: 1px solid {COLORS['border_focus']};
                border-radius: 6px;
                padding: 6px 10px;
                font-family: 'Menlo', 'Consolas', monospace;
                font-size: 12px;
            }}
        """)
        self.fx_edit.returnPressed.connect(self._commit_fx)
        self.fx_edit.installEventFilter(self)
        self.fx_btn_ok = QPushButton("✓")
        self.fx_btn_ok.setFixedSize(30, 30)
        self.fx_btn_ok.setEnabled(False)
        self.fx_btn_ok.setStyleSheet(make_btn("✓", "#1A3A2A", COLORS['accent_green'], COLORS['accent_green'], padding="1px", bold=True))
        self.fx_btn_ok.setToolTip("Construir y guardar en la celda (Enter)")
        self.fx_btn_ok.clicked.connect(self._commit_fx)
        self.fx_btn_cancel = QPushButton("✕")
        self.fx_btn_cancel.setFixedSize(30, 30)
        self.fx_btn_cancel.setEnabled(False)
        self.fx_btn_cancel.setStyleSheet(make_btn("✕", "#3A1A1A", COLORS['accent_red'], COLORS['accent_red'], padding="1px"))
        self.fx_btn_cancel.setToolTip("Cancelar (Esc)")
        self.fx_btn_cancel.clicked.connect(self._cancel_fx)
        fx.addWidget(fx_lbl)
        fx.addWidget(self.fx_edit, 1)
        fx.addWidget(self.fx_btn_ok)
        fx.addWidget(self.fx_btn_cancel)
        lay.addLayout(fx)

        self.table = QTableWidget(0, 1 + len(TEMPORALITIES))
        self.table.setHorizontalHeaderLabels(["Indicador (nombre)"] + TEMPORALITIES)
        self.table.setStyleSheet(f"""
            QTableWidget {{
                background: {COLORS['bg_card']};
                alternate-background-color: {COLORS['bg_panel']};
                color: {COLORS['text_primary']};
                border: 1px solid {COLORS['border']};
                border-radius: 6px;
                font-family: 'Menlo', 'Consolas', monospace;
                font-size: 10px;
                gridline-color: {COLORS['border']};
            }}
            QTableWidget::item:selected {{
                background: {COLORS['accent_blue']};
                color: #0D1117;
            }}
            QHeaderView::section {{
                background: {COLORS['bg_panel']};
                color: {COLORS['text_secondary']};
                border: none;
                border-right: 1px solid {COLORS['border']};
                padding: 6px;
                font-weight: 600;
            }}
        """)
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        vh = self.table.verticalHeader()
        vh.setDefaultSectionSize(46)
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(True)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.table.itemSelectionChanged.connect(self._refresh_detail)
        self.table.cellDoubleClicked.connect(self._edit_cell)
        self.table.cellClicked.connect(self._on_cell_clicked)
        self.table.installEventFilter(self)
        self.table.viewport().installEventFilter(self)
        lay.addWidget(self.table)

        hint = QLabel("Doble clic en una celda de fórmula para pegarla/verla ampliada. Cada fila = un indicador; cada columna = una temporaliad.")
        hint.setWordWrap(True)
        hint.setStyleSheet(f"color:{COLORS['text_muted']}; font-size:10px;")
        lay.addWidget(hint)
        return w

    def _build_detail(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(10, 0, 0, 0)
        lay.setSpacing(6)

        self.lbl_sel = QLabel("Selecciona una celda de fórmula…")
        self.lbl_sel.setStyleSheet(f"color:{COLORS['accent_blue']}; font-size:13px; font-weight:700;")
        self.lbl_sel.setWordWrap(True)

        self.lbl_resumen = QLabel("")
        self.lbl_resumen.setStyleSheet(f"color:{COLORS['text_secondary']}; font-size:11px; background:{COLORS['bg_input']}; padding:6px 8px; border-radius:4px;")
        self.lbl_resumen.setWordWrap(True)

        self.grp_used = QGroupBox("¿Dónde se usa esta fórmula?")
        ul = QVBoxLayout(self.grp_used)
        self.btn_back = QPushButton("↩ Ver usos de la fórmula completa")
        self.btn_back.setStyleSheet(make_btn("↩ Volver", COLORS['bg_input'], COLORS['border'], COLORS['text_secondary'], padding="3px 8px", font_size=10))
        self.btn_back.setVisible(False)
        self.btn_back.clicked.connect(self._reset_used_view)
        self.list_used = QListWidget()
        self.list_used.itemClicked.connect(self._on_used_clicked)
        ul.addWidget(self.btn_back)
        ul.addWidget(self.list_used)

        self.grp_cont = QGroupBox("Qué fórmulas completas incluye")
        cl = QVBoxLayout(self.grp_cont)
        self.list_cont = QListWidget()
        self.list_cont.itemClicked.connect(self._on_used_clicked)
        cl.addWidget(self.list_cont)

        self.grp_terms = QGroupBox("Bosquejo · estructura por términos")
        tl = QVBoxLayout(self.grp_terms)
        self.list_terms = QListWidget()
        self.list_terms.itemClicked.connect(self._on_term_clicked)
        self.list_terms.setStyleSheet(f"font-family:'Menlo','Consolas',monospace; font-size:10px;")
        tl.addWidget(self.list_terms)

        lay.addWidget(self.lbl_sel)
        lay.addWidget(self.lbl_resumen)
        lay.addWidget(self.grp_used)
        lay.addWidget(self.grp_cont)
        lay.addWidget(self.grp_terms)
        return w

    # ── Datos ─────────────────────────────────
    def _repo(self):
        rows = []
        for r in range(self.table.rowCount()):
            name = ""
            it0 = self.table.item(r, 0)
            if it0:
                name = it0.text().strip()
            for ci, temp in enumerate(TEMPORALITIES, 1):
                it = self.table.item(r, ci)
                raw = it.text().strip() if it else ""
                if not raw:
                    continue
                rows.append({"row": r, "col": ci, "name": name, "temp": temp,
                             "raw": raw, "inner": extract_inner(raw)})
        return rows

    def _selected_entry(self, repo):
        r = self.table.currentRow()
        c = self.table.currentColumn()
        if r < 0 or c < 0:
            return None
        if c == 0:
            return next((e for e in repo if e["row"] == r), None)
        return next((e for e in repo if e["row"] == r and e["col"] == c), None)

    # ── Detalle / análisis ────────────────────
    def _refresh_detail(self, *_):
        repo = self._repo()
        sel = self._selected_entry(repo)
        self.btn_back.setVisible(False)
        self.grp_used.setTitle("¿Dónde se usa esta fórmula?")
        self._term_hits = []

        if not sel:
            self.lbl_sel.setText("Selecciona una celda de fórmula…")
            self.lbl_resumen.setText("")
            self.list_used.clear()
            self.list_cont.clear()
            self.list_terms.clear()
            return

        self.lbl_sel.setText(f"{sel['name'] or '(sin nombre)'}  ·  {sel['temp']}")
        desc = describe_inner(sel["inner"])
        self.lbl_resumen.setText("Bosquejo:  " + " + ".join(desc))

        used_in = []
        contains = []
        for e in repo:
            if e is sel:
                continue
            if len(sel["inner"]) >= MIN_MATCH and sel["inner"] in e["inner"]:
                used_in.append(e)
            if len(e["inner"]) >= MIN_MATCH and e["inner"] in sel["inner"]:
                contains.append(e)

        self.list_used.clear()
        if not used_in:
            self.list_used.addItem("(ninguna otra fórmula la usa completa)")
        else:
            for e in used_in:
                self.list_used.addItem(self._entry_item(e))

        self.list_cont.clear()
        if not contains:
            self.list_cont.addItem("(no incluye fórmulas completas de la librería)")
        else:
            for e in contains:
                self.list_cont.addItem(self._entry_item(e))

        items = decompose(sel["inner"])
        self._term_hits = []
        self.list_terms.clear()

        def preview(t):
            t = _norm(t)
            return t if len(t) <= 70 else t[:67] + "…"

        for i, (level, op, text) in enumerate(items):
            t = _norm(text)
            hits = [e for e in repo if e is not sel and len(t) >= MIN_MATCH and t in e["inner"]]
            self._term_hits.append(hits)
            indent = "  " * level
            mark = f"   ✓ en {len(hits)} fórmula(s)" if hits else ""
            item = QListWidgetItem(f"{indent}#{i + 1}  [{term_kind(t).upper()}]  {preview(t)}{mark}")
            item.setData(Qt.ItemDataRole.UserRole, i)
            item.setToolTip(t)
            self.list_terms.addItem(item)

        if not items:
            self.list_terms.addItem("(no se pudieron separar términos)")

    def _entry_item(self, e):
        it = QListWidgetItem(f"{e['name'] or '(sin nombre)'}  ·  {e['temp']}")
        it.setData(Qt.ItemDataRole.UserRole, (e["row"], e["col"]))
        it.setToolTip(e["inner"][:400])
        return it

    def _on_term_clicked(self, item):
        idx = item.data(Qt.ItemDataRole.UserRole)
        if idx is None or not (0 <= idx < len(self._term_hits)):
            return
        hits = self._term_hits[idx]
        self.btn_back.setVisible(True)
        self.grp_used.setTitle(f"Término #{idx + 1} — se usa en:")
        self.list_used.clear()
        if not hits:
            self.list_used.addItem("(este término no aparece en otras fórmulas)")
        else:
            for e in hits:
                self.list_used.addItem(self._entry_item(e))

    def _reset_used_view(self):
        self._refresh_detail()

    def _on_used_clicked(self, item):
        data = item.data(Qt.ItemDataRole.UserRole)
        if not data:
            return
        r, c = data
        self.table.setCurrentCell(r, c)

    def _on_pick_toggled(self, on):
        self._update_pick_style()
        if self.builder and getattr(self.builder, "excel_tab", None):
            self.builder.excel_tab.set_insert_mode(on)

    def _update_pick_style(self):
        if self.b_pick.isChecked():
            self.b_pick.setStyleSheet(make_btn(
                "🖱 Insertar al hacer clic", "#1A1A3A", COLORS['accent_purple'], COLORS['accent_purple'],
                padding="5px 12px", font_size=11, bold=True))
        else:
            self.b_pick.setStyleSheet(make_btn(
                "🖱 Insertar al hacer clic", COLORS['bg_card'], COLORS['border'], COLORS['text_secondary'],
                padding="5px 12px", font_size=11))

    def _on_cell_clicked(self, r, c):
        if c <= 0:
            return
        item0 = self.table.item(r, 0)
        name = item0.text().strip() if item0 else ""
        if not name:
            return
        if self.fx_active:
            self._fx_insert(f"{name}:{TEMPORALITIES[c - 1]}")
            self.fx_edit.setFocus()
            return
        if self.builder and getattr(self.builder, "excel_tab", None) \
                and self.builder.excel_tab.insert_mode:
            self.builder.excel_tab.insert_cell_ref(name, TEMPORALITIES[c - 1])

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.KeyPress:
            key = event.key()
            if obj is self.fx_edit and self.fx_active and key == Qt.Key.Key_Escape:
                self._cancel_fx()
                return True
            if self.fx_active and key in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
                self._commit_fx()
                return True
            if obj in (self.table, self.table.viewport()) and event.text() == "=" and not self.fx_active:
                r = self.table.currentRow()
                c = self.table.currentColumn()
                if r >= 0 and c > 0 and self.table.state() != QAbstractItemView.State.EditingState:
                    self._start_fx(r, c)
                    return True
        return super().eventFilter(obj, event)

    # ── Modo fx (tipo Excel) ──────────────────
    def _start_fx(self, r, c):
        self.fx_target = (r, c)
        self.fx_active = True
        self.fx_edit.setEnabled(True)
        self.fx_edit.clear()
        self.fx_edit.setText("=")
        self.fx_edit.setCursorPosition(1)
        self.fx_btn_ok.setEnabled(True)
        self.fx_btn_cancel.setEnabled(True)
        self.lbl_status.setText("Construyendo fórmula: haz clic en las celdas que quieras sumar y termina con Enter.")
        self.fx_edit.setFocus()

    def _fx_insert(self, text):
        pos = self.fx_edit.cursorPosition()
        full = self.fx_edit.text()
        self.fx_edit.setText(full[:pos] + text + full[pos:])
        self.fx_edit.setCursorPosition(pos + len(text))

    def _commit_fx(self):
        if not self.fx_active or not self.fx_target:
            return
        r, c = self.fx_target
        raw = self.fx_edit.text().strip()
        expr0 = raw[1:] if raw.startswith("=") else raw
        repo = self._repo()
        formula, err = generate_formula(expr0, repo, "2")
        if err:
            QMessageBox.warning(self, "No se pudo construir", err)
            return
        item = self.table.item(r, c)
        if item is None:
            item = QTableWidgetItem("")
            self.table.setItem(r, c, item)
        item.setText(formula)
        self.table.resizeRowToContents(r)
        name = self.table.item(r, 0).text().strip() if self.table.item(r, 0) else ""
        self._refresh_detail()
        self.lbl_status.setText(f"✓ Construida y guardada en '{name or '(sin nombre)'}' · {TEMPORALITIES[c - 1]}")
        self._cancel_fx()
        self.table.setCurrentCell(r, c)

    def _cancel_fx(self):
        self.fx_active = False
        self.fx_target = None
        self.fx_edit.clear()
        self.fx_edit.setEnabled(False)
        self.fx_btn_ok.setEnabled(False)
        self.fx_btn_cancel.setEnabled(False)

    # ── Operaciones de filas / celdas ─────────
    def _add_row(self):
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QTableWidgetItem(""))
        self.table.resizeRowToContents(r)
        self.table.setCurrentCell(r, 0)
        self._refresh_detail()
        self.table.editItem(self.table.item(r, 0))

    def _del_row(self):
        r = self.table.currentRow()
        if r < 0:
            QMessageBox.information(self, "Info", "Selecciona una fila primero.")
            return
        name = self.table.item(r, 0).text() if self.table.item(r, 0) else ""
        reply = QMessageBox.question(
            self, "Confirmar", f"¿Eliminar la fila '{name or ('#' + str(r + 1))}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply != QMessageBox.StandardButton.Yes:
            return
        self.table.removeRow(r)
        self._refresh_detail()

    def _edit_cell(self, r, c):
        col0 = self.table.item(r, 0)
        if col0 is None:
            self.table.setItem(r, 0, QTableWidgetItem(""))
            col0 = self.table.item(r, 0)
        if c == 0:
            self.table.editItem(col0)
            self._refresh_detail()
            return
        item = self.table.item(r, c)
        cur = item.text() if item else ""
        title = f"Fórmula  ·  {col0.text() or '(sin nombre)'}  ·  {TEMPORALITIES[c - 1]}"
        dlg = FormulaEditorDialog(title, cur, self)
        if dlg.exec():
            if item is None:
                item = QTableWidgetItem("")
                self.table.setItem(r, c, item)
            item.setText(dlg.text())
            self.table.resizeRowToContents(r)
            self._refresh_detail()

    # ── Persistencia ──────────────────────────
    def _table_rows(self):
        rows = []
        for r in range(self.table.rowCount()):
            name = self.table.item(r, 0).text().strip() if self.table.item(r, 0) else ""
            formulas = {}
            for ci, temp in enumerate(TEMPORALITIES, 1):
                it = self.table.item(r, ci)
                formulas[temp] = it.text().strip() if it else ""
            rows.append({"name": name, "formulas": formulas})
        return rows

    def _save_json(self):
        path, _ = QFileDialog.getSaveFileName(self, "Guardar librería de fórmulas", LIB_DEFAULT_JSON, "JSON (*.json)")
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump({"version": 1, "rows": self._table_rows()}, f, ensure_ascii=False, indent=2)
        except Exception as ex:
            QMessageBox.critical(self, "Error al guardar", str(ex))
            return
        self.lbl_status.setText(f"Guardado ✓  {os.path.basename(path)}")

    def _load_json(self):
        path, _ = QFileDialog.getOpenFileName(self, "Abrir librería de fórmulas", LIB_DEFAULT_JSON, "JSON (*.json)")
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as ex:
            QMessageBox.critical(self, "Error al abrir", str(ex))
            return
        self._load_data(data)
        self.lbl_status.setText(f"Cargado ✓  {len(data.get('rows', []))} filas  ·  {os.path.basename(path)}")

    def _load_data(self, data):
        self.table.setRowCount(0)
        for row in data.get("rows", []):
            r = self.table.rowCount()
            self.table.insertRow(r)
            self.table.setItem(r, 0, QTableWidgetItem(row.get("name", "")))
            for temp, formula in (row.get("formulas") or {}).items():
                if temp in TEMPORALITIES:
                    ci = TEMPORALITIES.index(temp) + 1
                    self.table.setItem(r, ci, QTableWidgetItem(formula))
        self._refresh_detail()

    def _do_replace(self):
        dlg = ReplaceDialog(self)
        if not dlg.exec():
            return
        find_text, repl = dlg.values()
        if not find_text:
            return
        count = 0
        for r in range(self.table.rowCount()):
            for ci in range(1, self.table.columnCount()):
                it = self.table.item(r, ci)
                if it and find_text in it.text():
                    count += it.text().count(find_text)
                    it.setText(it.text().replace(find_text, repl))
        if count == 0:
            QMessageBox.information(self, "Reemplazo masivo", "No se encontró el texto buscado.")
            return
        self._refresh_detail()
        QMessageBox.information(self, "Reemplazo masivo", f"Se reemplazaron {count} coincidencia(s).")

    def _clear_all(self):
        reply = QMessageBox.question(
            self, "Confirmar", "¿Vaciar toda la librería?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply != QMessageBox.StandardButton.Yes:
            return
        self.table.setRowCount(0)
        self.lbl_status.setText("")
        self._refresh_detail()


# ──────────────────────────────────────────────
# PUNTO DE ENTRADA
# ──────────────────────────────────────────────
def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Formula Builder")

    # Fuente por defecto
    font = QFont("Consolas", 10)
    app.setFont(font)

    window = FormulaBuilder()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
