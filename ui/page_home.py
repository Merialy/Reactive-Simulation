"""
Стартова сторінка: що вміє програма (зі схемою моделі), швидкі переходи, останні збережені
результати і коротко про розробника. Вікно «Про програму» – AboutDialog.
"""
from collections import OrderedDict

from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg
from matplotlib.figure import Figure
from PySide6.QtCore import QSize, Qt
from PySide6.QtWidgets import (QDialog, QGridLayout, QHBoxLayout, QLabel, QListWidget, QListWidgetItem, QVBoxLayout,
                               QWidget)

from config import Config
from ui.home_diagram import draw_model
from ui.i18n import is_en, tr
from ui.icons import icon
from ui.widgets import PALETTE, Card, button, muted


def dev(key: str) -> str:
    """Поле з Config.DEVELOPER потрібною мовою."""
    d = Config.DEVELOPER
    return d.get(key + "_en", d.get(key, "")) if is_en() else d.get(key, "")


def _icon_label(name: str, color_key: str = "accent", size: int = 22) -> QLabel:
    lab = QLabel()
    lab.setProperty("icon_name", name)
    lab.setProperty("icon_color", color_key)
    lab.setProperty("icon_size", size)
    lab.setPixmap(icon(name, PALETTE.get(color_key, "#2563EB"), size).pixmap(QSize(size, size)))
    lab.setFixedSize(size + 4, size + 4)
    return lab


class HomePage(QWidget):
    def __init__(self, go, open_history, show_about, parent=None):
        super().__init__(parent)
        self.setObjectName("Page")
        self.go = go
        self.open_history = open_history
        lay = QVBoxLayout(self)
        lay.setContentsMargins(28, 24, 28, 20)
        lay.setSpacing(16)

        # ---------------- Головний блок: опис + схема моделі ----------------
        hero = Card()
        hl = QHBoxLayout()
        hl.setSpacing(24)
        left = QVBoxLayout()
        left.setSpacing(10)
        title = QLabel(tr("Симулятор моделі реактивності"))
        title.setObjectName("PageTitle")
        title.setWordWrap(True)
        left.addWidget(title)
        left.addWidget(muted(tr("Розрахунок показників реактивності бортового обчислювача: час реакції на "
                                "переривання, черги, завантаження процесора і ймовірність порушення дедлайнів. "
                                "Модель – три черги з пріоритетами і витісненням, імітаційне моделювання з "
                                "довірчими інтервалами.")))
        left.addSpacing(6)
        for text, desc, ic, key, kind in (
                (tr("Серія сценаріїв"), tr("таблиця параметрів, 2 типи подій"), "table", "excel", "Primary"),
                (tr("Типи подій і відмови"), tr("довільні типи подій, шторм, WCET"), "layers", "scenario", ""),
                (tr("Історія результатів"), tr("збережені запуски, порівняння"), "history", "history", "")):
            row = QHBoxLayout()
            b = button(text, kind, slot=lambda _=False, k=key: self.go(k), icon=ic)
            b.setMinimumWidth(230)
            row.addWidget(b)
            row.addWidget(muted(desc, hint=True), 1)
            left.addLayout(row)
        left.addStretch(1)
        self.storage = muted("", hint=True)
        left.addWidget(self.storage)
        hl.addLayout(left, 2)
        self.fig = Figure(figsize=(6.4, 2.9), dpi=100)
        self.canvas = FigureCanvasQTAgg(self.fig)
        self.canvas.setMinimumSize(420, 220)
        self.canvas.setToolTip(tr("Схема моделі: події проходять первинну обробку (ISR), потім чекають у черзі "
                                  "свого класу; ЖРЧ обслуговуються раніше за МРЧ"))
        hl.addWidget(self.canvas, 3)
        hero.add(hl)
        lay.addWidget(hero)

        # ---------------- Нижній ряд ----------------
        grid = QGridLayout()
        grid.setSpacing(16)

        how = Card(tr("Як це працює"))
        for i, (ic, txt) in enumerate((
                ("table", tr("Задайте параметри: відкрийте Excel-файл, готовий сценарій або введіть значення вручну.")),
                ("refresh", tr("Ліворуч унизу оберіть кількість прогонів R (для звіту – 5 і більше) і seed.")),
                ("play", tr("Натисніть «Запустити» – розрахунок іде у фоні на всіх ядрах процесора.")),
                ("chart", tr("Перегляньте «Результати» і «Графіки», збережіть звіт в Excel або результат у базу.")))):
            r = QHBoxLayout()
            r.setSpacing(10)
            r.addWidget(_icon_label(ic, "accent", 18), 0, Qt.AlignmentFlag.AlignTop)
            t = muted(f"{i + 1}. {txt}")
            r.addWidget(t, 1)
            how.add(r)
        how.lay.addStretch(1)
        grid.addWidget(how, 0, 0)

        rec = Card(tr("Останні результати"), tr("Подвійний клік – відкрити в історії"))
        self.recent = QListWidget()
        self.recent.setObjectName("Recent")
        self.recent.setMinimumHeight(150)
        self.recent.itemDoubleClicked.connect(self._open_recent)
        rec.add(self.recent, 1)
        grid.addWidget(rec, 0, 1)

        about = Card(tr("Про програму"))
        head = QHBoxLayout()
        head.addWidget(_icon_label("cpu", "accent", 34))
        nm = QVBoxLayout()
        nm.setSpacing(0)
        t1 = QLabel(tr("Реактивність БО") + f"  {Config.APP_VERSION}")
        t1.setObjectName("CardTitle")
        nm.addWidget(t1)
        nm.addWidget(muted(tr("Магістерська робота, {y}", y=dev("year")), hint=True))
        head.addLayout(nm, 1)
        about.add(head)
        for ic, txt in (("user", f"{dev('name')}, {dev('role')}"),
                        ("school", dev("university")),
                        ("info", tr("Науковий керівник: {s}", s=dev("supervisor")))):
            r = QHBoxLayout()
            r.setSpacing(8)
            r.addWidget(_icon_label(ic, "muted", 16), 0, Qt.AlignmentFlag.AlignTop)
            r.addWidget(muted(txt, hint=True), 1)
            about.add(r)
        about.lay.addStretch(1)
        br = QHBoxLayout()
        br.addWidget(button(tr("Детальніше"), slot=show_about, icon="info"))
        br.addStretch(1)
        about.add(br)
        grid.addWidget(about, 0, 2)
        grid.setColumnStretch(0, 4)
        grid.setColumnStretch(1, 4)
        grid.setColumnStretch(2, 3)
        lay.addLayout(grid, 1)

        foot = muted(f"© {dev('year')} {dev('name')} · {tr('ХАІ')}", hint=True)
        foot.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(foot)
        self.redraw()

    # ------------------------------------------------------------------
    def redraw(self):
        """Перемалювати схему і значки (після зміни теми)."""
        if PALETTE:
            draw_model(self.fig, PALETTE)
            self.canvas.draw_idle()
            for lab in self.findChildren(QLabel):
                name = lab.property("icon_name")
                if name:
                    size = int(lab.property("icon_size") or 18)
                    lab.setPixmap(icon(name, PALETTE.get(lab.property("icon_color"), "#2563EB"), size)
                                  .pixmap(QSize(size, size)))

    def set_recent(self, rows):
        """Останні збережені запуски, згруповані за джерелом (файл або сценарій) і часом збереження."""
        self.recent.clear()
        groups = OrderedDict()
        for r in rows:
            src = (r.get("tags") or [r.get("scenario_name", "")])[0]
            when = (r.get("created_at") or "").replace("T", " ")[:16]
            key = (src, when)
            g = groups.setdefault(key, {"n": 0, "st": r.get("settings") or {}})
            g["n"] += 1
        for (src, when), g in list(groups.items())[:6]:
            st = g["st"]
            n = tr("{n} сценаріїв", n=g["n"]) if g["n"] > 1 else tr("1 сценарій")
            it = QListWidgetItem(f"{src}\n{when}  ·  {n}  ·  R = {st.get('replications')}, seed = {st.get('seed')}")
            it.setData(Qt.ItemDataRole.UserRole, src)
            it.setIcon(icon("results", PALETTE.get("muted", "#64748B"), 18))
            self.recent.addItem(it)
        if not groups:
            it = QListWidgetItem(tr("Ще немає збережених результатів. Після запуску натисніть «Зберегти в базу даних»."))
            it.setFlags(Qt.ItemFlag.NoItemFlags)
            self.recent.addItem(it)

    def _open_recent(self, it):
        src = it.data(Qt.ItemDataRole.UserRole)
        if src:
            self.open_history(src)


class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr("Про програму"))
        self.setMinimumWidth(560)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(26, 22, 26, 18)
        lay.setSpacing(12)
        head = QHBoxLayout()
        head.addWidget(_icon_label("cpu", "accent", 48))
        box = QVBoxLayout()
        t = QLabel(tr("Симулятор моделі реактивності"))
        t.setObjectName("PageTitle")
        box.addWidget(t)
        box.addWidget(muted(tr("Версія {v} · модель {m}", v=Config.APP_VERSION, m=Config.MODEL_VERSION)))
        head.addLayout(box, 1)
        lay.addLayout(head)
        lay.addWidget(muted(tr("Програма розраховує показники реактивності бортового обчислювача методом "
                               "імітаційного моделювання: час реакції на переривання, довжини черг, завантаження "
                               "процесора і ймовірність порушення дедлайнів, у тому числі під час відмов.")))
        g = QGridLayout()
        g.setHorizontalSpacing(14)
        g.setVerticalSpacing(8)
        rows = [("user", tr("Розробник"), f"{dev('name')}, {dev('role')}"),
                ("school", tr("Заклад"), dev("university")),
                ("info", tr("Науковий керівник"), dev("supervisor")),
                ("results", tr("Тема роботи"), dev("thesis")),
                ("mail", tr("Пошта"), dev("email")),
                ("cpu", tr("Технології"), "Python · PySide6 (Qt) · NumPy · pandas · matplotlib · MongoDB")]
        for i, (ic, k, v) in enumerate(rows):
            g.addWidget(_icon_label(ic, "muted", 16), i, 0, Qt.AlignmentFlag.AlignTop)
            kl = QLabel(k)
            kl.setObjectName("Muted")
            g.addWidget(kl, i, 1, Qt.AlignmentFlag.AlignTop)
            vl = QLabel(v)
            vl.setWordWrap(True)
            vl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            g.addWidget(vl, i, 2)
        g.setColumnStretch(2, 1)
        lay.addLayout(g)
        lay.addWidget(muted(tr("Дані розробника можна змінити у файлі config.py (розділ DEVELOPER)."), hint=True))
        row = QHBoxLayout()
        row.addStretch(1)
        row.addWidget(button(tr("Закрити"), "Primary", slot=self.accept))
        lay.addLayout(row)
