import sys
import json
import copy
from dataclasses import dataclass
from typing import List, Dict, Optional, Tuple
from PySide6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem, QListWidget, QListWidgetItem, QPushButton, QFileDialog, QLabel, QSplitter, QMessageBox, QFormLayout, QLineEdit, QGroupBox, QCheckBox, QScrollArea, QTabWidget, QSpinBox, QAbstractItemView
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QKeySequence, QShortcut

@dataclass
class CellItem:
    id_form: int
    label: str
    codigo: str
    tipo: int
    deci: int
    posicion: str
    valor: str

def parse_pos(pos: str) -> Tuple[int, int]:
    r, c = pos.split(":")
    return int(r), int(c)

def fmt_pos(r: int, c: int) -> str:
    return f"{r}:{c}"

def col_name(idx: int) -> str:
    s = ""
    x = idx
    while True:
        s = chr(ord('A') + (x % 26)) + s
        x = x // 26 - 1
        if x < 0:
            break
    return s

class GridEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Cierres Maker")
        self.items: List[CellItem] = []
        self.items_by_codigo: Dict[str, CellItem] = {}
        self.pos_to_item: Dict[str, CellItem] = {}
        self.current_codigo: Optional[str] = None
        self.groups: Dict[str, Dict[str, CellItem]] = {}
        self.root_data = None
        self.global_ids: Dict[str, Optional[int]] = {"dia": None, "semana": None, "mes": None, "anio": None}
        self.updating = False
        self.undo_stack = []
        self.redo_stack = []
        self.copied_data: Optional[Dict] = None

        self.tabs = QTabWidget()
        self.tabs.currentChanged.connect(self.on_tab_changed)
        self.tables: Dict[str, QTableWidget] = {}
        self.current_period: str = "dia"
        self.init_table_for_period("dia")

        self.list = QListWidget()
        self.list.currentRowChanged.connect(self.on_list_change)

        self.load_btn = QPushButton("Cargar JSON")
        self.load_btn.clicked.connect(self.on_load_json)

        self.save_btn = QPushButton("Guardar JSON")
        self.save_btn.clicked.connect(self.on_save_json)

        self.add_row_btn = QPushButton("Agregar Fila")
        self.add_row_btn.clicked.connect(self.on_insert_row)
        self.add_col_btn = QPushButton("Agregar Columna")
        self.add_col_btn.clicked.connect(self.on_insert_col)
        
        self.undo_btn = QPushButton("Deshacer")
        self.undo_btn.clicked.connect(self.undo)
        self.redo_btn = QPushButton("Rehacer")
        self.redo_btn.clicked.connect(self.redo)
        
        self.copy_btn = QPushButton("Copiar Celda")
        self.copy_btn.clicked.connect(self.copy_selection)
        self.paste_btn = QPushButton("Pegar Celda")
        self.paste_btn.clicked.connect(self.paste_selection)
        
        self.clear_btn = QPushButton("Limpiar")
        self.clear_btn.clicked.connect(self.on_clear_all)
        
        QShortcut(QKeySequence("Ctrl+Z"), self, self.undo)
        QShortcut(QKeySequence("Ctrl+Y"), self, self.redo)
        QShortcut(QKeySequence("Ctrl+Shift+Z"), self, self.redo)
        QShortcut(QKeySequence("Ctrl+C"), self, self.copy_selection)
        QShortcut(QKeySequence("Ctrl+V"), self, self.paste_selection)

        self.move_mode = QCheckBox("Mover grupo con clic")
        self.move_mode.setChecked(False)

        self.current_label = QLabel("Sin selección")

        self.ids_box = QGroupBox("IDs de formularios")
        self.id_form_dia = QLineEdit()
        self.id_form_semana = QLineEdit()
        self.id_form_mes = QLineEdit()
        self.id_form_anio = QLineEdit()
        ids_form = QFormLayout()
        ids_form.addRow("Día", self.id_form_dia)
        ids_form.addRow("Semana", self.id_form_semana)
        ids_form.addRow("Mes", self.id_form_mes)
        ids_form.addRow("Año", self.id_form_anio)
        self.ids_box.setLayout(ids_form)
        self.update_ids_btn = QPushButton("Actualizar IDs")
        self.update_ids_btn.clicked.connect(self.on_update_ids)

        self.detail_box = QGroupBox("Detalle celda")
        self.det_label = QLineEdit()
        self.det_codigo = QLineEdit()
        self.det_posicion = QLineEdit(); self.det_posicion.setReadOnly(True)
        self.det_id = QLineEdit()
        self.det_tipo = QLineEdit()
        self.det_deci = QLineEdit()
        self.det_valor = QLineEdit()
        
        # Connect editing signals
        self.det_label.editingFinished.connect(self.on_detail_edited)
        self.det_codigo.editingFinished.connect(self.on_detail_edited)
        self.det_id.editingFinished.connect(self.on_detail_edited)
        self.det_tipo.editingFinished.connect(self.on_detail_edited)
        self.det_deci.editingFinished.connect(self.on_detail_edited)
        self.det_valor.editingFinished.connect(self.on_detail_edited)

        det_form = QFormLayout()
        det_form.addRow("Label", self.det_label)
        det_form.addRow("Código", self.det_codigo)
        det_form.addRow("Posición", self.det_posicion)
        det_form.addRow("Id form", self.det_id)
        det_form.addRow("Tipo", self.det_tipo)
        det_form.addRow("Deci", self.det_deci)
        det_form.addRow("Valor", self.det_valor)
        self.detail_box.setLayout(det_form)

        # Modernize button labels (visual only)
        self.load_btn.setText("📁 Cargar JSON")
        self.save_btn.setText("💾 Guardar JSON")
        self.add_row_btn.setText("➕ Agregar Fila")
        self.add_col_btn.setText("📊 Agregar Columna")
        self.undo_btn.setText("🗑️ Deshacer")
        self.redo_btn.setText("♻️ Rehacer")
        self.copy_btn.setText("📋 Copiar Celda")
        self.paste_btn.setText("📌 Pegar Celda")
        self.clear_btn.setText("🧹 Limpiar")
        self.move_mode.setText("👥 Mover grupo con clic")

        # Header
        header = QWidget()
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(8, 4, 8, 0)
        title_box = QWidget()
        title_layout = QVBoxLayout(title_box)
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_layout.setSpacing(0)
        title_main = QLabel("Cierres Maker")
        title_sub = QLabel("Sistema de Gestión de Inventarios")
        title_main.setFont(QFont("Segoe UI", 14, QFont.Bold))
        title_sub.setFont(QFont("Segoe UI", 9))
        title_layout.addWidget(title_main)
        title_layout.addWidget(title_sub)
        header_layout.addWidget(title_box, 1)
        self.count_label = QLabel("0 Items | Cargados")
        self.count_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        header_layout.addWidget(self.count_label, 0, Qt.AlignRight)

        # Toolbar (search + quick filters)
        toolbar = QWidget()
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(8, 0, 8, 4)
        
        # Row Height
        rh_lbl = QLabel("Alto:")
        self.row_height_spin = QSpinBox()
        self.row_height_spin.setRange(20, 200)
        self.row_height_spin.setValue(24)
        self.row_height_spin.valueChanged.connect(self.on_row_height_change)
        self.row_height_spin.setFixedWidth(60)
        toolbar_layout.addWidget(rh_lbl)
        toolbar_layout.addWidget(self.row_height_spin)
        
        search_box = QWidget()
        sb_layout = QHBoxLayout(search_box)
        sb_layout.setContentsMargins(8, 2, 8, 2)
        search_icon = QLabel("🔍")
        self.search_entry = QLineEdit()
        self.search_entry.setPlaceholderText("Buscar en datos...")
        self.search_entry.textChanged.connect(self.on_search_changed)
        self.search_entry.returnPressed.connect(self.on_search_return)
        sb_layout.addWidget(search_icon)
        sb_layout.addWidget(self.search_entry)
        toolbar_layout.addWidget(search_box, 1)
        filters_label = QLabel("Filtros:")
        toolbar_layout.addWidget(filters_label, 0)
        for name in ["INVENTARIO", "FRUTO", "INGRESO"]:
            btn = QPushButton(name)
            toolbar_layout.addWidget(btn, 0)

        ids_panel = QWidget()
        ids_layout = QVBoxLayout(ids_panel)
        ids_layout.addWidget(self.ids_box)
        ids_layout.addWidget(self.update_ids_btn)

        # Left sidebar (actions + list)
        left_controls = QWidget()
        left_controls_layout = QVBoxLayout(left_controls)
        left_controls_layout.setContentsMargins(8, 8, 8, 8)
        left_controls_layout.addWidget(self.load_btn)
        left_controls_layout.addWidget(self.save_btn)
        left_controls_layout.addWidget(self.add_row_btn)
        left_controls_layout.addWidget(self.add_col_btn)
        left_controls_layout.addWidget(self.undo_btn)
        left_controls_layout.addWidget(self.redo_btn)
        left_controls_layout.addWidget(self.copy_btn)
        left_controls_layout.addWidget(self.paste_btn)
        left_controls_layout.addWidget(self.clear_btn)
        left_controls_layout.addWidget(self.move_mode)
        left_controls_layout.addWidget(self.current_label)

        left_splitter = QSplitter(Qt.Vertical)
        left_splitter.addWidget(left_controls)
        left_splitter.addWidget(self.list)
        left_splitter.setStretchFactor(0, 0)
        left_splitter.setStretchFactor(1, 1)

        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setWidget(left_splitter)
        left_scroll.setMinimumWidth(220)

        # Right details panel (ids + details)
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(8, 8, 8, 8)
        right_layout.addWidget(ids_panel)
        right_layout.addWidget(self.detail_box)
        right_scroll = QScrollArea()
        right_scroll.setWidgetResizable(True)
        right_scroll.setWidget(right_panel)
        right_scroll.setMinimumWidth(240)

        splitter = QSplitter()
        splitter.addWidget(left_scroll)
        splitter.addWidget(self.tabs)
        splitter.addWidget(right_scroll)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setStretchFactor(2, 0)

        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)
        root_layout.addWidget(header)
        root_layout.addWidget(toolbar)
        root_layout.addWidget(splitter, 1)
        self.setCentralWidget(root)
        self.setStyleSheet(
            "QWidget{background:#0A0E27; color:#fff;}"
            "QScrollArea{background:#0A0E27; border:0;}"
            "QSplitter::handle{background:#1E1E2E;}"
            "QGroupBox{background:#1E1E2E; border:1px solid #2D2D44; padding:8px; margin-top:10px;}"
            "QGroupBox::title{subcontrol-origin: margin; left:10px; padding:0 6px;}"
            "QTableWidget{background:#1E1E2E; color:#fff; gridline-color:#2D2D44; selection-background-color: #5865F2; selection-color: #fff;}"
            "QTableWidget::item:selected{background:#5865F2; color:#fff;}"
            "QTableWidget::item:selected:!active{background:#5865F2; color:#fff;}"
            "QTableWidget::item{padding:4px; border:0px;}"
            "QHeaderView::section{background:#2D2D44; color:#fff; border:0; padding:6px;}"
            "QPushButton{background:#5865F2; color:#fff; border:0; padding:8px; font-weight:bold;}"
            "QPushButton:hover{background:#4752C4;}"
            "QListWidget{background:#1E1E2E; color:#fff; border:0;}"
            "QLineEdit{background:#2D2D44; color:#fff; border:0; padding:8px;}"
            "QLabel{color:#fff;}"
        )
    
    def period_title(self, p: str) -> str:
        return {"dia": "Día", "semana": "Semana", "mes": "Mes", "anio": "Año"}.get(p, p.capitalize())
    
    def init_table_for_period(self, period: str):
        if period in self.tables:
            return
        tbl = QTableWidget()
        tbl.setSelectionBehavior(QTableWidget.SelectItems)
        tbl.setSelectionMode(QTableWidget.SingleSelection)
        tbl.setFont(QFont("Arial", 11))
        tbl.setProperty("period", period)
        tbl.cellClicked.connect(lambda r, c, p=period: self.on_cell_clicked_tab(p, r, c))
        tbl.cellChanged.connect(lambda r, c, p=period: self.on_cell_changed_tab(p, r, c))
        self.tables[period] = tbl
        self.tabs.addTab(tbl, self.period_title(period))
        # Point helpers to the active table
        self.table = self.tables[self.current_period]
    
    def setup_tabs_from_items(self):
        present = set()
        for d in self.items:
            p = self.get_period(d) or "dia"
            present.add(p)
        order = ["dia", "mes", "anio", "semana"]
        # remove all tabs and rebuild to satisfy ordering and presence
        while self.tabs.count() > 0:
            self.tabs.removeTab(0)
        self.tables.clear()
        for p in order:
            if p in present or (not present and p == "dia"):
                self.init_table_for_period(p)
        # set current period to first tab
        if self.tabs.count() > 0:
            w = self.tabs.widget(0)
            self.current_period = w.property("period")
            self.table = self.tables[self.current_period]
    
    def on_tab_changed(self, idx: int):
        if idx < 0:
            return
        w = self.tabs.widget(idx)
        if not w:
            return
        period = w.property("period")
        self.current_period = period
        self.table = self.tables[period]
        
        # Update search results for the new period
        if hasattr(self, "search_entry"):
            self.on_search_changed(self.search_entry.text())
    
    def on_cell_clicked_tab(self, period: str, r0: int, c0: int):
        self.current_period = period
        self.table = self.tables[period]
        self.on_cell_clicked(r0, c0)
    
    def on_cell_changed_tab(self, period: str, r: int, c: int):
        self.current_period = period
        self.table = self.tables[period]
        self.on_cell_changed(r, c)

    def save_state(self):
        state = {
            'items': copy.deepcopy(self.items),
            'global_ids': copy.deepcopy(self.global_ids)
        }
        self.undo_stack.append(state)
        if len(self.undo_stack) > 50:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def restore_state(self, state):
        self.items = copy.deepcopy(state['items'])
        self.global_ids = copy.deepcopy(state['global_ids'])
        self.items_by_codigo = {d.codigo: d for d in self.items}
        self.pos_to_item = {d.posicion: d for d in self.items}
        self.build_groups()
        self.refresh_list()
        self.render_from_items()
        self.update_duplicates()
        
        # Restore selection if possible
        if self.current_codigo and self.current_codigo in self.items_by_codigo:
            item = self.items_by_codigo[self.current_codigo]
            r, c = parse_pos(item.posicion)
            p = self.get_period(item) or "dia"
            if p in self.tables:
                idx = [self.tables[t].property("period") for t in self.tables].index(p) if p in self.tables else 0
                # ensure tab is active
                for i in range(self.tabs.count()):
                    w = self.tabs.widget(i)
                    if w.property("period") == p:
                        self.tabs.setCurrentIndex(i)
                        break
                self.table = self.tables[p]
                self.table.setCurrentCell(r, c)
            self.show_cell_details(r, c)
            self.fill_id_fields(item)
        else:
            # If current selection is gone, try to keep table selection or clear details
            r = self.table.currentRow()
            c = self.table.currentColumn()
            if r >= 0 and c >= 0:
                self.show_cell_details(r, c)
            else:
                self.fill_id_fields(None)

    def undo(self):
        if not self.undo_stack:
            return
        
        current_state = {
            'items': copy.deepcopy(self.items),
            'global_ids': copy.deepcopy(self.global_ids)
        }
        self.redo_stack.append(current_state)
        
        state = self.undo_stack.pop()
        self.restore_state(state)

    def redo(self):
        if not self.redo_stack:
            return
            
        current_state = {
            'items': copy.deepcopy(self.items),
            'global_ids': copy.deepcopy(self.global_ids)
        }
        self.undo_stack.append(current_state)
        
        state = self.redo_stack.pop()
        self.restore_state(state)

    def on_list_change(self, row: int):
        # row param corresponds to visual row in QListWidget, which matches 
        # index in self.items ONLY if not filtered.
        # Now we use UserRole to get the absolute index in self.items
        
        item_widget = self.list.currentItem()
        if not item_widget:
            self.current_codigo = None
            self.current_label.setText("Sin selección")
            self.fill_id_fields(None)
            return
            
        idx = item_widget.data(Qt.UserRole)
        if idx is None or not isinstance(idx, int) or idx < 0 or idx >= len(self.items):
             return

        item = self.items[idx]
        self.current_codigo = item.codigo
        r, c = parse_pos(item.posicion)
        p = self.get_period(item) or "dia"
        # Switch to the item's period tab
        for i in range(self.tabs.count()):
            w = self.tabs.widget(i)
            if w.property("period") == p:
                self.tabs.setCurrentIndex(i)
                break
        tbl = self.tables.get(p)
        if tbl and not tbl.item(r, c):
            tbl.setItem(r, c, QTableWidgetItem(""))
        
        if tbl:
            # Force selection and scroll to item
            qitem = tbl.item(r, c)
            if qitem:
                tbl.setCurrentItem(qitem)
                tbl.scrollToItem(qitem, QAbstractItemView.PositionAtCenter)
            else:
                tbl.setCurrentCell(r, c)
                tbl.scrollToItem(tbl.item(r, c), QAbstractItemView.PositionAtCenter)
            
            self.table = tbl
            
        self.show_cell_details(r, c)
        self.current_label.setText(f"Seleccionado: {item.codigo} | Fila {r} Col {c}")
        self.fill_id_fields(item)

    def on_cell_clicked(self, r0: int, c0: int):
        r = r0
        c = c0
        self.table.setCurrentCell(r, c)
        self.show_cell_details(r, c)
        if self.current_codigo and self.move_mode.isChecked():
            self.save_state()
            itm = self.items_by_codigo[self.current_codigo]
            self.move_group_for_item(itm, r, c)
            self.update_duplicates()

    def copy_selection(self):
        r = self.table.currentRow()
        c = self.table.currentColumn()
        pos = fmt_pos(r, c)
        item = self.pos_to_item.get(pos)
        if item:
            self.copied_data = {
                "label": item.label,
                "codigo": item.codigo,
                "id_form": item.id_form,
                "tipo": item.tipo,
                "deci": item.deci,
                "valor": item.valor
            }
            self.current_label.setText(f"Copiado: {item.codigo}")
        else:
            self.copied_data = None
            self.current_label.setText("Nada para copiar")

    def paste_selection(self):
        if not hasattr(self, 'copied_data') or not self.copied_data:
            return
        
        r = self.table.currentRow()
        c = self.table.currentColumn()
        if r < 0 or c < 0:
            return
            
        self.save_state()
        
        pos = fmt_pos(r, c)
        item = self.pos_to_item.get(pos)
        
        data = self.copied_data
        
        if not item:
            item = CellItem(
                id_form=data["id_form"],
                label=data["label"],
                codigo=data["codigo"],
                tipo=data["tipo"],
                deci=data["deci"],
                posicion=pos,
                valor=data["valor"]
            )
            self.items.append(item)
            self.pos_to_item[pos] = item
            if item.codigo:
                self.items_by_codigo[item.codigo] = item
        else:
            item.label = data["label"]
            item.codigo = data["codigo"]
            item.id_form = data["id_form"]
            item.tipo = data["tipo"]
            item.deci = data["deci"]
            item.valor = data["valor"]
            if item.codigo:
                self.items_by_codigo[item.codigo] = item

        self.place_item(item) # Re-render cell
        self.show_cell_details(r, c)
        self.refresh_list()
        self.update_duplicates()

    def update_duplicates(self):
        counts = {}
        for d in self.items:
            if d.codigo:
                counts[d.codigo] = counts.get(d.codigo, 0) + 1
        
        for d in self.items:
            r, c = parse_pos(d.posicion)
            p = self.get_period(d) or "dia"
            tbl = self.tables.get(p)
            itm = tbl.item(r, c) if tbl else None
            if itm:
                if d.codigo and counts[d.codigo] > 1:
                    itm.setBackground(QColor("#ffcccc")) # Light red
                    itm.setForeground(QColor("black"))
                    itm.setToolTip(f"Código duplicado: {d.codigo}")
                else:
                    itm.setBackground(QColor("#1E1E2E")) # Restore default
                    itm.setForeground(QColor("white"))
                    itm.setToolTip("")

    def on_row_height_change(self, val: int):
        for tbl in self.tables.values():
            for r in range(tbl.rowCount()):
                tbl.setRowHeight(r, val)

    def on_search_changed(self, text: str):
        text = text.lower().strip()
        self.list.clear()
        
        # Filter list items
        for idx, d in enumerate(self.items):
            # Only show items belonging to the current active period
            p = self.get_period(d) or "dia"
            if p != self.current_period:
                continue
                
            # Match against code or label
            full_str = f"{d.codigo} | {d.label}"
            if not text or text in full_str.lower():
                item = QListWidgetItem(full_str)
                item.setData(Qt.UserRole, idx)
                self.list.addItem(item)
    
    def on_search_return(self):
        text = self.search_entry.text().lower().strip()
        if not text:
            return
            
        # Find first match in items that is NOT the current selection if possible,
        # or just the first match.
        # For simplicity, let's find the first match in the filtered list.
        
        if self.list.count() > 0:
            # Select first item in the list, which triggers on_list_change -> selecting cell
            self.list.setCurrentRow(0)
            self.list.setFocus()

    def on_load_json(self):
        path, _ = QFileDialog.getOpenFileName(self, "Abrir JSON", "", "Archivos (*.json *.txt);;Todos (*.*)")
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                raw = f.read()
            try:
                data = json.loads(raw)
            except Exception:
                lines = [json.loads(l) for l in raw.splitlines() if l.strip()]
                data = lines
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
            return
        def rows_from_datosAG(x):
            if isinstance(x, dict) and isinstance(x.get("datosAG"), list):
                out = []
                for group in x["datosAG"]:
                    if isinstance(group, list):
                        out.extend([d for d in group if isinstance(d, dict)])
                return out
            return None
        self.root_data = data
        rows = rows_from_datosAG(data)
        if rows is None:
            def collect_entries(x):
                out = []
                def rec(v):
                    if isinstance(v, dict):
                        if all(k in v for k in ("codigo", "posicion", "label")):
                            out.append(v)
                        else:
                            for vv in v.values():
                                rec(vv)
                    elif isinstance(v, list):
                        for e in v:
                            rec(e)
                rec(x)
                return out
            rows = collect_entries(data)
        items: List[CellItem] = []
        for d in rows:
            try:
                items.append(CellItem(
                    id_form=int(d.get("id_form", 0)),
                    label=str(d.get("label", "")),
                    codigo=str(d.get("codigo", "")),
                    tipo=int(d.get("tipo", 0)),
                    deci=int(d.get("deci", 0)),
                    posicion=str(d.get("posicion", "1:1")),
                    valor=str(d.get("valor", "")),
                ))
            except Exception:
                pass
        self.items = items
        self.items_by_codigo = {d.codigo: d for d in self.items}
        self.pos_to_item = {d.posicion: d for d in self.items}
        self.build_groups()
        self.extract_global_ids()
        self.refresh_list()
        self.render_from_items()
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.current_label.setText(f"Cargados: {len(self.items)} items")
        if hasattr(self, "count_label") and self.count_label:
            self.count_label.setText(f"{len(self.items)} Items | Cargados")
        if not self.items:
            QMessageBox.information(self, "Aviso", "No se encontraron items válidos en el JSON")

    def on_save_json(self):
        if not self.items:
            return
        path, _ = QFileDialog.getSaveFileName(self, "Guardar JSON", "", "JSON (*.json)")
        if not path:
            return
        try:
            data = [d.__dict__ for d in self.items]
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def refresh_list(self):
        if hasattr(self, "search_entry"):
            self.on_search_changed(self.search_entry.text())
        else:
            self.on_search_changed("")

    def render_from_items(self):
        # Ensure tabs exist for current items
        self.setup_tabs_from_items()
        # Calculate global grid size
        if not self.items:
            for tbl in self.tables.values():
                tbl.setRowCount(0)
                tbl.setColumnCount(0)
                tbl.clearContents()
            return
        max_r = 0
        max_c = 0
        for d in self.items:
            r, c = parse_pos(d.posicion)
            max_r = max(max_r, r)
            max_c = max(max_c, c)
        # Prepare each table
        for period, tbl in self.tables.items():
            tbl.setRowCount(max_r + 1)
            tbl.setColumnCount(max_c + 1)
            tbl.clearContents()
            for r in range(tbl.rowCount()):
                tbl.setRowHeight(r, 24)
            for c in range(tbl.columnCount()):
                tbl.setColumnWidth(c, 120)
            tbl.setHorizontalHeaderLabels([col_name(i) for i in range(tbl.columnCount())])
            tbl.setVerticalHeaderLabels([str(i + 1) for i in range(tbl.rowCount())])
        # Place items only into their period's table
        self.updating = True
        for d in self.items:
            self.place_item(d)
        self.update_duplicates()
        self.updating = False

    def place_item(self, d: CellItem):
        p = self.get_period(d) or "dia"
        tbl = self.tables.get(p)
        if not tbl:
            return
        r, c = parse_pos(d.posicion)
        item = QTableWidgetItem(d.label)
        item.setData(Qt.UserRole, d.codigo)
        tbl.setItem(r, c, item)

    def get_period(self, d: CellItem) -> Optional[str]:
        code = d.codigo.strip().upper()
        if len(code) >= 2:
            prefix = code[:2]
            m = {"CD": "dia", "CS": "semana", "CM": "mes", "CA": "anio"}
            if prefix in m:
                return m[prefix]
        for k in ["dia", "semana", "mes", "anio"]:
            val = self.global_ids.get(k)
            if val is not None and d.id_form == val:
                return k
        lbl = d.label.upper()
        for k in ["DIA", "SEMANA", "MES", "AÑO", "ANIO"]:
            if k in lbl:
                return {"DIA": "dia", "SEMANA": "semana", "MES": "mes", "AÑO": "anio", "ANIO": "anio"}[k]
        return None

    def normalize_label(self, lbl: str) -> str:
        u = lbl.upper().strip()
        for k in [" DIA", " SEMANA", " MES", " AÑO", " ANIO"]:
            if u.endswith(k):
                u = u[: -len(k)]
                break
        return u

    def build_groups(self):
        self.groups = {}
        for d in self.items:
            p = self.get_period(d)
            base = self.normalize_label(d.label)
            if not p:
                p = "dia"
            if base not in self.groups:
                self.groups[base] = {}
            if p not in self.groups[base]:
                self.groups[base][p] = []
            self.groups[base][p].append(d)

    def move_group_for_item(self, pivot: CellItem, new_r: int, new_c: int):
        base = self.normalize_label(pivot.label)
        grp = self.groups.get(base, {})
        pr_r, pr_c = parse_pos(pivot.posicion)
        deltas: Dict[str, int] = {}
        
        for k, items_list in grp.items():
            if items_list:
                # Use the first item to determine the current column of this period block
                r, c = parse_pos(items_list[0].posicion)
                deltas[k] = c - pr_c
                
        max_c = 0
        max_r = 0
        # Determine current global grid size across items
        for d in self.items:
            rr, cc = parse_pos(d.posicion)
            max_r = max(max_r, rr)
            max_c = max(max_c, cc)
        max_c = max(max_c, (new_c + (max(deltas.values()) if deltas else 0)) + 1)
        max_r = max(max_r, new_r + 1)
        # Apply to all tables
        for tbl in self.tables.values():
            tbl.setRowCount(max_r)
            tbl.setColumnCount(max_c)
        
        for k, items_list in grp.items():
            # Clear old positions first
            for it in items_list:
                old_r, old_c = parse_pos(it.posicion)
                if old_r >= 0 and old_c >= 0:
                    p_it = self.get_period(it) or "dia"
                    tbl_it = self.tables.get(p_it)
                    if tbl_it:
                        tbl_it.setItem(old_r, old_c, QTableWidgetItem(""))
            
            # Update to new positions
            delta = deltas.get(k, 0)
            target_c = new_c + delta
            
            for it in items_list:
                it.posicion = fmt_pos(new_r, target_c)
                self.place_item(it)
                
        self.pos_to_item = {d.posicion: d for d in self.items}

    def fill_id_fields(self, item: Optional[CellItem]):
        if not item:
            self.id_form_dia.setText("")
            self.id_form_semana.setText("")
            self.id_form_mes.setText("")
            self.id_form_anio.setText("")
            return
        base = self.normalize_label(item.label)
        grp = self.groups.get(base, {})
        
        def get_id(k):
            # Return ID from the first item in the list, or global fallback
            items_list = grp.get(k, [])
            if items_list:
                return str(items_list[0].id_form)
            val = self.global_ids.get(k)
            return str(val) if val else ""
            
        self.id_form_dia.setText(get_id("dia"))
        self.id_form_semana.setText(get_id("semana"))
        self.id_form_mes.setText(get_id("mes"))
        self.id_form_anio.setText(get_id("anio"))

    def on_update_ids(self):
        if not self.current_codigo:
            return
        ret = QMessageBox.question(self, "Confirmar", "¿Actualizar IDs del grupo?")
        if ret != QMessageBox.Yes:
            return
        self.save_state()
        item = self.items_by_codigo[self.current_codigo]
        base = self.normalize_label(item.label)
        grp = self.groups.get(base, {})
        def to_int(s: str) -> Optional[int]:
            s = s.strip()
            if not s:
                return None
            try:
                return int(s)
            except:
                return None
        v_d = to_int(self.id_form_dia.text())
        v_s = to_int(self.id_form_semana.text())
        v_m = to_int(self.id_form_mes.text())
        v_a = to_int(self.id_form_anio.text())
        
        if grp.get("dia") and v_d is not None:
            for it in grp["dia"]: it.id_form = v_d
        if grp.get("semana") and v_s is not None:
            for it in grp["semana"]: it.id_form = v_s
        if grp.get("mes") and v_m is not None:
            for it in grp["mes"]: it.id_form = v_m
        if grp.get("anio") and v_a is not None:
            for it in grp["anio"]: it.id_form = v_a
            
        self.global_ids = {
            "dia": v_d if v_d is not None else self.global_ids.get("dia"),
            "semana": v_s if v_s is not None else self.global_ids.get("semana"),
            "mes": v_m if v_m is not None else self.global_ids.get("mes"),
            "anio": v_a if v_a is not None else self.global_ids.get("anio"),
        }
        self.apply_global_ids_to_root()

    def extract_global_ids(self):
        ids = {"dia": None, "semana": None, "mes": None, "anio": None}
        root = self.root_data
        try:
            if isinstance(root, dict) and isinstance(root.get("formularioC"), list):
                arr = root["formularioC"][0]
                cf = arr.get("cod_fechas", [])
                for e in cf:
                    tv = e.get("tipo_val")
                    if tv == "d":
                        ids["dia"] = e.get("id_form")
                    elif tv == "s":
                        ids["semana"] = e.get("id_form")
                    elif tv == "m":
                        ids["mes"] = e.get("id_form")
                    elif tv == "a":
                        ids["anio"] = e.get("id_form")
        except Exception:
            pass
        self.global_ids = ids

    def apply_global_ids_to_root(self):
        root = self.root_data
        try:
            if isinstance(root, dict) and isinstance(root.get("formularioC"), list):
                arr = root["formularioC"][0]
                cf = arr.get("cod_fechas", [])
                for e in cf:
                    tv = e.get("tipo_val")
                    if tv == "d" and self.global_ids.get("dia") is not None:
                        e["id_form"] = int(self.global_ids["dia"])
                    elif tv == "s" and self.global_ids.get("semana") is not None:
                        e["id_form"] = int(self.global_ids["semana"])
                    elif tv == "m" and self.global_ids.get("mes") is not None:
                        e["id_form"] = int(self.global_ids["mes"])
                    elif tv == "a" and self.global_ids.get("anio") is not None:
                        e["id_form"] = int(self.global_ids["anio"])
        except Exception:
            pass

    def update_root_with_items_and_ids(self, root):
        self.apply_global_ids_to_root()
        mapping = {d.codigo: d.posicion for d in self.items}
        def rec(v):
            if isinstance(v, dict):
                if "codigo" in v and "posicion" in v:
                    code = v.get("codigo")
                    if code in mapping:
                        v["posicion"] = mapping[code]
                for k in list(v.keys()):
                    v[k] = rec(v[k])
                return v
            elif isinstance(v, list):
                return [rec(e) for e in v]
            else:
                return v
        try:
            return rec(root)
        except Exception:
            return [d.__dict__ for d in self.items]

    def on_insert_row(self):
        idx = self.table.currentRow()
        if idx < 0:
            return
        self.save_state()
        updated = []
        for d in self.items:
            r, c = parse_pos(d.posicion)
            if r >= idx:
                r += 1
            d.posicion = fmt_pos(r, c)
            updated.append(d)
        self.items = updated
        self.items_by_codigo = {d.codigo: d for d in self.items}
        self.pos_to_item = {d.posicion: d for d in self.items}
        self.render_from_items()
        for tbl in self.tables.values():
            for c in range(tbl.columnCount()):
                tbl.setItem(idx, c, QTableWidgetItem(""))

    def on_insert_col(self):
        idx = self.table.currentColumn()
        if idx < 0:
            return
        self.save_state()
        updated = []
        for d in self.items:
            r, c = parse_pos(d.posicion)
            if c >= idx:
                c += 1
            d.posicion = fmt_pos(r, c)
            updated.append(d)
        self.items = updated
        self.items_by_codigo = {d.codigo: d for d in self.items}
        self.pos_to_item = {d.posicion: d for d in self.items}
        self.render_from_items()
        for tbl in self.tables.values():
            for r in range(tbl.rowCount()):
                tbl.setItem(r, idx, QTableWidgetItem(""))

    def on_cell_changed(self, r: int, c: int):
        if self.updating:
            return
        itm = self.table.item(r, c)
        txt = itm.text().strip() if itm else ""
        pos = fmt_pos(r, c)
        existing = self.pos_to_item.get(pos)
        
        changed = False
        if not txt:
            if existing:
                changed = True
        else:
            if not existing:
                changed = True
            elif existing.label != txt:
                changed = True
        
        if changed:
            self.save_state()
            
        if not txt:
            if existing:
                try:
                    self.items.remove(existing)
                except ValueError:
                    pass
                self.pos_to_item.pop(pos, None)
                self.items_by_codigo.pop(existing.codigo, None)
            return
        if existing:
            existing.label = txt
        else:
            new = CellItem(
                id_form=0,
                label=txt,
                codigo="",
                tipo=0,
                deci=0,
                posicion=pos,
                valor="",
            )
            self.items.append(new)
            self.pos_to_item[pos] = new
        self.build_groups()
        self.show_cell_details(r, c)
        self.update_duplicates()
        self.refresh_list()

    def show_cell_details(self, r: int, c: int):
        pos = fmt_pos(r, c)
        it = self.pos_to_item.get(pos)
        if it:
            self.det_label.setText(it.label)
            self.det_codigo.setText(it.codigo)
            self.det_posicion.setText(it.posicion)
            self.det_id.setText(str(it.id_form))
            self.det_tipo.setText(str(it.tipo))
            self.det_deci.setText(str(it.deci))
            self.det_valor.setText(it.valor)
        else:
            self.det_label.setText("")
            self.det_codigo.setText("")
            self.det_posicion.setText(f"{r}:{c}")
            self.det_id.setText("")
            self.det_tipo.setText("")
            self.det_deci.setText("")
            self.det_valor.setText("")
            
    def on_detail_edited(self):
        if self.updating: return
        r = self.table.currentRow()
        c = self.table.currentColumn()
        if r < 0 or c < 0:
            return
        
        pos = fmt_pos(r, c)
        item = self.pos_to_item.get(pos)
        
        def safe_int(s):
            try: return int(s)
            except: return 0
            
        new_label = self.det_label.text()
        new_code = self.det_codigo.text()
        new_id = safe_int(self.det_id.text())
        new_tipo = safe_int(self.det_tipo.text())
        new_deci = safe_int(self.det_deci.text())
        new_valor = self.det_valor.text()
        
        if not item:
            if not any([new_label, new_code, new_valor, new_id, new_tipo, new_deci]):
                return
            
            self.save_state()
            item = CellItem(
                id_form=new_id,
                label=new_label,
                codigo=new_code,
                tipo=new_tipo,
                deci=new_deci,
                posicion=pos,
                valor=new_valor
            )
            self.items.append(item)
            self.pos_to_item[pos] = item
            if new_code:
                self.items_by_codigo[new_code] = item
            
            self.updating = True
            qitem = QTableWidgetItem(new_label)
            qitem.setData(Qt.UserRole, new_code)
            self.table.setItem(r, c, qitem)
            self.updating = False
            
            self.build_groups() # Rebuild groups for the new item
            self.refresh_list()
            self.current_label.setText(f"Cargados: {len(self.items)} items")
            if hasattr(self, "count_label") and self.count_label:
                self.count_label.setText(f"{len(self.items)} Items | Cargados")
            return

        if (item.label == new_label and
            item.codigo == new_code and
            item.id_form == new_id and
            item.tipo == new_tipo and
            item.deci == new_deci and
            item.valor == new_valor):
            return

        self.save_state()
        
        # Identify siblings BEFORE applying changes
        base = self.normalize_label(item.label)
        p = self.get_period(item) or "dia"
        siblings = []
        if base in self.groups and p in self.groups[base]:
            siblings = [x for x in self.groups[base][p] if x is not item]

        old_code = item.codigo
        if old_code != new_code:
            if old_code in self.items_by_codigo:
                del self.items_by_codigo[old_code]
            if new_code:
                self.items_by_codigo[new_code] = item
        
        # Update the main item
        item.label = new_label
        item.codigo = new_code
        item.id_form = new_id
        item.tipo = new_tipo
        item.deci = new_deci
        item.valor = new_valor
        
        # Update siblings (everything EXCEPT code and id_form)
        for sib in siblings:
            sib.label = new_label
            sib.tipo = new_tipo
            sib.deci = new_deci
            sib.valor = new_valor
            # Do NOT sync codigo or id_form
        
        self.updating = True
        qitem = self.table.item(r, c)
        if not qitem:
            qitem = QTableWidgetItem(new_label)
            self.table.setItem(r, c, qitem)
        else:
            qitem.setText(new_label)
            qitem.setData(Qt.UserRole, new_code)
        self.updating = False
        
        self.build_groups() # Rebuild groups as label might have changed
        self.refresh_list()
        self.update_duplicates()

    def on_clear_all(self):
        self.save_state()
        self.items = []
        self.items_by_codigo = {}
        self.pos_to_item = {}
        self.groups = {}
        for tbl in self.tables.values():
            tbl.clearContents()
            tbl.setRowCount(0)
            tbl.setColumnCount(0)
        self.list.clear()
        self.current_label.setText("Cargados: 0 items")
        if hasattr(self, "count_label") and self.count_label:
            self.count_label.setText("0 Items | Cargados")

def main():
    app = QApplication(sys.argv)
    w = GridEditor()
    w.resize(1200, 700)
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
