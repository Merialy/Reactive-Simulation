"""
Сторінка «Дані Excel»: серія сценаріїв з двома типами подій (8 параметрів у стовпці).
Таблицю можна відкрити з файлу, редагувати, доповнювати і зберігати.
Кожна комірка перевіряється одразу; останній рядок показує завантаження процесора U.
"""
import os

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (QAbstractItemView, QFileDialog, QHBoxLayout, QHeaderView, QMessageBox,
                               QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget)

from ui.core import PARAM_HINT, PARAM_SHORT, ParamTable, fmt, parse_number
from ui.widgets import Banner, Card, PageHeader, button, color, muted, ro_item
from ui.i18n import tr

U_ROW = 8


class ExcelPage(QWidget):
    run_requested = Signal()          # натиснуто «Запустити серію»
    table_changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("Page")
        self.data = ParamTable(columns=3)
        self._loading = False

        lay = QVBoxLayout(self)
        lay.setContentsMargins(28, 24, 28, 20)
        lay.setSpacing(14)

        self.header = PageHeader(tr("Серія сценаріїв з Excel"),
                                 tr("Кожен стовпець – окремий сценарій з двома типами подій (МРЧ і ЖРЧ). "
                                 "Комірки можна редагувати прямо в таблиці."))
        self.header.add_action(button(tr("Відкрити Excel…"), slot=self.open_file,
                                      tooltip=tr("Завантажити таблицю параметрів з .xlsx")))
        self.header.add_action(button(tr("Зберегти як…"), slot=self.save_file))
        self.run_btn = self.header.add_action(button(tr("▶  Запустити серію"), "Primary", slot=lambda: self.run_requested.emit()))
        lay.addWidget(self.header)

        self.banner = Banner()
        lay.addWidget(self.banner)

        card = Card()
        tools = QHBoxLayout()
        self.file_label = muted(tr("Файл не відкрито – можна заповнити таблицю вручну."))
        tools.addWidget(self.file_label, 1)
        tools.addWidget(button(tr("+ Стовпець"), slot=self.add_column, tooltip=tr("Додати порожній сценарій")))
        tools.addWidget(button(tr("Дублювати"), slot=self.duplicate_column,
                               tooltip=tr("Скопіювати виділений стовпець у новий сценарій")))
        tools.addWidget(button(tr("Видалити"), "Danger", slot=self.delete_columns,
                               tooltip=tr("Видалити виділені стовпці")))
        card.add(tools)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.table.horizontalHeader().setDefaultSectionSize(110)
        self.table.horizontalHeader().setMinimumSectionSize(80)
        self.table.verticalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        self.table.verticalHeader().setDefaultSectionSize(34)
        self.table.itemChanged.connect(self._on_item_changed)
        card.add(self.table, 1)
        card.add(muted(tr("Червоним підсвічено значення поза допустимими межами (підказка – при наведенні). "
                       "Рядок U – завантаження процесора: зелений < 0,95, жовтий 0,95–1, червоний ≥ 1 "
                       "(перевантаження). Порожні стовпці не моделюються."), hint=True))
        lay.addWidget(card, 1)
        self.refresh()

    # ------------------------------------------------------------------
    def refresh(self):
        """Перемалювати таблицю з self.data."""
        self._loading = True
        t = self.table
        n = self.data.n_cols
        t.clear()
        t.setRowCount(9)
        t.setColumnCount(n)
        t.setHorizontalHeaderLabels([tr("Сценарій {n}", n=c + 1) for c in range(n)])
        t.setVerticalHeaderLabels(PARAM_SHORT + [tr("U (завантаження)")])
        for r in range(8):
            t.verticalHeaderItem(r).setToolTip(PARAM_HINT[r])
        t.verticalHeaderItem(U_ROW).setToolTip(tr("U = λ1·(s1 + s2) + λ2·(s1 + s2b) – частка часу, коли процесор зайнятий"))
        for c in range(n):
            for r in range(8):
                it = QTableWidgetItem(fmt(self.data.values[r][c], 8))
                it.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                t.setItem(r, c, it)
                self._style_cell(r, c)
            self._update_u(c)
        self._loading = False
        self.file_label.setText(tr("Файл: {f}  ·  сценаріїв: {n}", f=self.data.source_name(), n=self._filled_count())
                                if self.data.path else
                                tr("Таблицю заповнено вручну  ·  сценаріїв: {n}", n=self._filled_count()))

    def _filled_count(self) -> int:
        return sum(1 for c in range(self.data.n_cols) if any(self.data.values[r][c] is not None for r in range(8)))

    def _style_cell(self, r: int, c: int):
        it = self.table.item(r, c)
        if it is None:
            return
        v = self.data.values[r][c]
        problem = self.data.bounds_problem(r, v)
        if problem:
            it.setBackground(color("err_bg"))
            it.setForeground(color("err"))
            it.setToolTip(problem)
        else:
            it.setData(Qt.ItemDataRole.BackgroundRole, None)
            it.setData(Qt.ItemDataRole.ForegroundRole, None)
            it.setToolTip(PARAM_HINT[r])

    def _update_u(self, c: int):
        u = self.data.column_u(c)
        it = ro_item("" if u is None else fmt(u, 3))
        f = it.font()
        f.setBold(True)
        it.setFont(f)
        if u is not None:
            lvl = "err" if u >= 1 else ("warn" if u > 0.95 else "ok")
            it.setBackground(color(lvl + "_bg"))
            it.setForeground(color(lvl))
            it.setToolTip({"ok": tr("Процесор не перевантажений"),
                           "warn": tr("Близько до перевантаження: потрібні більші N і R"),
                           "err": tr("Перевантаження: черга МРЧ росте необмежено")}[lvl])
        self.table.setItem(U_ROW, c, it)

    def _on_item_changed(self, it: QTableWidgetItem):
        if self._loading or it.row() >= 8:
            return
        r, c = it.row(), it.column()
        try:
            v = parse_number(it.text())
        except ValueError:
            self._loading = True
            it.setText(fmt(self.data.values[r][c], 8))
            self._loading = False
            self.banner.show_message(tr("«{t}» – не число. Використовуйте цифри, кома або крапка – десятковий роздільник.",
                                        t=it.text()), "err")
            return
        if r == 3 and v is not None:
            v = float(int(round(v)))
        self.data.set(r, c, v)
        self._loading = True
        it.setText(fmt(v, 8))
        self._style_cell(r, c)
        self._update_u(c)
        self._loading = False
        self.banner.hide()
        self.table_changed.emit()

    # ------------------------------------------------------------------
    def selected_columns(self):
        cols = sorted({i.column() for i in self.table.selectedIndexes()})
        if not cols and self.table.currentColumn() >= 0:
            cols = [self.table.currentColumn()]
        return cols

    def add_column(self):
        self.data.add_column()
        self.refresh()
        self.table.setCurrentCell(0, self.data.n_cols - 1)
        self.table_changed.emit()

    def duplicate_column(self):
        cols = self.selected_columns()
        if not cols:
            self.banner.show_message(tr("Спочатку виділіть комірку в стовпці, який треба скопіювати."), "warn")
            return
        self.data.add_column(copy_from=cols[0])
        self.refresh()
        self.table.setCurrentCell(0, self.data.n_cols - 1)
        self.table_changed.emit()

    def delete_columns(self):
        cols = self.selected_columns()
        if not cols:
            self.banner.show_message(tr("Виділіть комірки в стовпцях, які треба видалити."), "warn")
            return
        names = ", ".join(str(c + 1) for c in cols)
        if QMessageBox.question(self, tr("Видалити стовпці"), tr("Видалити сценарії {n}?", n=names)) \
                != QMessageBox.StandardButton.Yes:
            return
        self.data.remove_columns(cols)
        self.refresh()
        self.table_changed.emit()

    # ------------------------------------------------------------------
    def open_file(self, path: str = None):
        if not path:
            start = os.path.dirname(self.data.path) if self.data.path else os.getcwd()
            path, _ = QFileDialog.getOpenFileName(self, tr("Відкрити таблицю параметрів"), start,
                                                  tr("Excel (*.xlsx *.xls);;Усі файли (*)"))
        if not path:
            return False
        try:
            data, problems = ParamTable.load(path)
        except Exception as e:
            self.banner.show_message(tr("Не вдалося відкрити файл: {e}", e=e), "err")
            return False
        self.data = data
        self.refresh()
        real = problems[1:] if data.row_format else problems
        if real:
            self.banner.show_message(tr("Файл відкрито, але є зауваження:\n• ") + "\n• ".join(problems[:8]), "warn")
        elif problems:
            self.banner.show_message(tr("Відкрито «{f}»: {n} сценаріїв.", f=os.path.basename(path), n=self._filled_count())
                                     + " " + problems[0], "ok")
        else:
            self.banner.show_message(tr("Відкрито «{f}»: {n} сценаріїв.", f=os.path.basename(path), n=self._filled_count())
                                     + " " + tr("Натисніть «Запустити серію»."), "ok")
        self.table_changed.emit()
        return True

    def save_file(self):
        start = self.data.path or os.path.join(os.getcwd(), tr("параметри.xlsx"))
        path, _ = QFileDialog.getSaveFileName(self, tr("Зберегти таблицю параметрів"), start, "Excel (*.xlsx)")
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        try:
            self.data.save(path)
            self.refresh()
            self.banner.show_message(tr("Таблицю збережено: {p}", p=path), "ok")
        except Exception as e:
            self.banner.show_message(tr("Не вдалося зберегти: {e}", e=e), "err")
