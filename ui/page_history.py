"""
Сторінка «Історія»: результати, збережені в MongoDB (або в JSON-файлах, якщо MongoDB недоступна).
Можна знайти, відкрити, порівняти кілька результатів і видалити зайві.
"""
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (QAbstractItemView, QDialog, QHBoxLayout, QHeaderView, QLineEdit, QMessageBox,
                               QTableWidget, QVBoxLayout, QWidget)

from config import Config
from ui.core import METRIC_SHORT, fmt, fmt_ci
from ui.widgets import Banner, Card, PageHeader, button, muted, ro_item
from ui.i18n import tr, tr_msg

COLS = [tr("Збережено"), tr("Сценарій"), "R", "seed", "U", tr("Теги"), tr("Ідентифікатор")]


class HistoryPage(QWidget):
    open_result = Signal(dict)         # документ результату для сторінки «Результати»

    def __init__(self, get_repo, parent=None):
        super().__init__(parent)
        self.setObjectName("Page")
        self.get_repo = get_repo
        self.rows = []

        lay = QVBoxLayout(self)
        lay.setContentsMargins(28, 24, 28, 20)
        lay.setSpacing(14)
        self.header = PageHeader(tr("Історія результатів"), tr("Результати, збережені кнопкою «Зберегти в базу даних»."))
        self.header.add_action(button(tr("Оновити"), slot=self.reload, icon="refresh"))
        lay.addWidget(self.header)
        self.banner = Banner()
        lay.addWidget(self.banner)

        card = Card()
        tools = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText(tr("Пошук за назвою сценарію або тегом (наприклад, назвою Excel-файлу)…"))
        self.search.textChanged.connect(self._filter)
        tools.addWidget(self.search, 1)
        tools.addWidget(button(tr("Відкрити"), "Primary", slot=self.open_selected, icon="arrow",
                               tooltip=tr("Показати вибраний результат на сторінці «Результати»")))
        tools.addWidget(button(tr("Порівняти"), slot=self.compare_selected, icon="compare",
                               tooltip=tr("Виділіть 2 або більше рядків (Ctrl + клік)")))
        tools.addWidget(button(tr("Видалити"), "Danger", slot=self.delete_selected, icon="trash"))
        card.add(tools)
        self.table = QTableWidget(0, len(COLS))
        self.table.setHorizontalHeaderLabels(COLS)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(32)
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.doubleClicked.connect(lambda *_: self.open_selected())
        card.add(self.table, 1)
        self.count_label = muted("", hint=True)
        card.add(self.count_label)
        lay.addWidget(card, 1)

    # ------------------------------------------------------------------
    def reload(self):
        try:
            repo = self.get_repo()
            self.rows = repo.list_results(limit=2000)
            self.header.subtitle.setText(tr("Сховище: {d}", d=tr_msg(repo.describe())))
            self.banner.hide()
        except Exception as e:
            self.rows = []
            self.banner.show_message(tr("Не вдалося прочитати сховище: {e}", e=e), "err")
        self._fill()

    def _fill(self):
        t = self.table
        t.setRowCount(len(self.rows))
        for i, r in enumerate(self.rows):
            st = r.get("settings") or {}
            vals = [(r.get("created_at") or "").replace("T", " "), r.get("scenario_name", ""),
                    str(st.get("replications", "")), str(st.get("seed", "")), fmt(r.get("utilization"), 3),
                    ", ".join(r.get("tags") or []), r["id"]]
            for j, v in enumerate(vals):
                it = ro_item(v, align_right=j in (2, 3, 4))
                if j == 0:
                    it.setData(Qt.ItemDataRole.UserRole, r["id"])
                t.setItem(i, j, it)
        self._filter(self.search.text())

    def _filter(self, text):
        text = (text or "").lower().strip()
        shown = 0
        for i, r in enumerate(self.rows):
            hay = (r.get("scenario_name", "") + " " + " ".join(r.get("tags") or [])).lower()
            hide = bool(text) and text not in hay
            self.table.setRowHidden(i, hide)
            shown += not hide
        self.count_label.setText(tr("Показано {a} з {b}. Подвійний клік – відкрити.", a=shown, b=len(self.rows)))

    def _selected_ids(self):
        rows = sorted({i.row() for i in self.table.selectedIndexes()})
        return [self.table.item(r, 0).data(Qt.ItemDataRole.UserRole) for r in rows
                if not self.table.isRowHidden(r)]

    # ------------------------------------------------------------------
    def open_selected(self):
        ids = self._selected_ids()
        if not ids:
            self.banner.show_message(tr("Виділіть рядок у таблиці."), "warn")
            return
        doc = self.get_repo().get_result(ids[0])
        if doc:
            self.open_result.emit(doc)

    def delete_selected(self):
        ids = self._selected_ids()
        if not ids:
            self.banner.show_message(tr("Виділіть рядки, які треба видалити."), "warn")
            return
        if QMessageBox.question(self, tr("Видалити результати"),
                                tr("Видалити результатів зі сховища: {n}? Цю дію не можна скасувати.", n=len(ids))) \
                != QMessageBox.StandardButton.Yes:
            return
        repo = self.get_repo()
        n = sum(1 for i in ids if repo.delete_result(i))
        self.reload()
        self.banner.show_message(tr("Видалено: {n}.", n=n), "ok")

    def compare_selected(self):
        ids = self._selected_ids()
        if len(ids) < 2:
            self.banner.show_message(tr("Для порівняння виділіть щонайменше два рядки (Ctrl + клік)."), "warn")
            return
        repo = self.get_repo()
        docs = [d for d in (repo.get_result(i) for i in ids[:8]) if d]
        CompareDialog(docs, self).exec()


class CompareDialog(QDialog):
    """Порівняння 18 показників кількох збережених результатів."""

    def __init__(self, docs, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr("Порівняння результатів"))
        self.resize(1100, 680)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(18, 18, 18, 18)
        lay.addWidget(muted(tr("Середнє ± 95 % ДІ. Часи – у мілісекундах. Найкраще значення в рядку виділено жирним "
                            "(для показників, де менше – краще).")))
        t = QTableWidget(18, len(docs))
        t.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        t.setAlternatingRowColors(True)
        t.setVerticalHeaderLabels([f"{i}. {n}" for i, n in enumerate(METRIC_SHORT, 1)])
        for j, d in enumerate(docs):
            st = d.get("settings") or {}
            hi = ro_item(f"{d.get('scenario_name', '')[:40]}\nR={st.get('replications')}, seed={st.get('seed')}",
                         align_right=False)
            hi.setToolTip(d.get("scenario_name", ""))
            t.setHorizontalHeaderItem(j, hi)
        for m in range(1, 19):
            vals = []
            for j, d in enumerate(docs):
                c = next((x for x in d.get("classic", []) if x["metric"] == m), None)
                vals.append(c["mean"] if c else None)
                t.setItem(m - 1, j, ro_item(fmt_ci(c["mean"], c["ci"]) if c else "—"))
            if m not in (14, 15, 16):                 # завантаження – не «краще/гірше»
                present = [v for v in vals if v is not None]
                if present and len(set(present)) > 1:
                    best = min(present)
                    for j, v in enumerate(vals):
                        if v == best:
                            it = t.item(m - 1, j)
                            f = it.font(); f.setBold(True); it.setFont(f)
        t.resizeColumnsToContents()
        t.verticalHeader().setDefaultSectionSize(28)
        lay.addWidget(t, 1)
        row = QHBoxLayout()
        row.addStretch(1)
        row.addWidget(button(tr("Закрити"), slot=self.accept))
        lay.addLayout(row)
