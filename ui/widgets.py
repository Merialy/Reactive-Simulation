"""
Спільні елементи інтерфейсу: картки, заголовки сторінок, «значки», поле для чисел,
фоновий виконавець (щоб вікно не «зависало» під час симуляції).
"""
import traceback

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (QFrame, QHBoxLayout, QLabel, QLineEdit, QPushButton, QSizePolicy,
                               QTableWidgetItem, QVBoxLayout, QWidget)

from ui.core import fmt, parse_number

# Поточна палітра (встановлює MainWindow при зміні теми)
PALETTE = {}


def color(name: str) -> QColor:
    return QColor(PALETTE.get(name, "#000000"))


def repolish(w: QWidget):
    """Застосувати стиль після зміни динамічної властивості (level, invalid)."""
    w.style().unpolish(w)
    w.style().polish(w)


class Card(QFrame):
    """Біла картка із заокругленими кутами; має заголовок (необов'язково) і вертикальний макет."""

    def __init__(self, title: str = "", subtitle: str = "", parent=None):
        super().__init__(parent)
        self.setObjectName("Card")
        self.lay = QVBoxLayout(self)
        self.lay.setContentsMargins(18, 16, 18, 16)
        self.lay.setSpacing(10)
        if title:
            head = QHBoxLayout()
            box = QVBoxLayout()
            box.setSpacing(2)
            t = QLabel(title)
            t.setObjectName("CardTitle")
            box.addWidget(t)
            if subtitle:
                s = QLabel(subtitle)
                s.setObjectName("Hint")
                s.setWordWrap(True)
                box.addWidget(s)
            head.addLayout(box, 1)
            self.head_actions = QHBoxLayout()
            self.head_actions.setSpacing(6)
            head.addLayout(self.head_actions)
            self.lay.addLayout(head)

    def add(self, w, stretch: int = 0):
        if isinstance(w, QWidget):
            self.lay.addWidget(w, stretch)
        else:
            self.lay.addLayout(w, stretch)
        return w


class PageHeader(QWidget):
    """Заголовок сторінки: назва, підзаголовок і кнопки праворуч."""

    def __init__(self, title: str, subtitle: str = "", parent=None):
        super().__init__(parent)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        box = QVBoxLayout()
        box.setSpacing(2)
        self.title = QLabel(title)
        self.title.setObjectName("PageTitle")
        self.title.setWordWrap(True)          # на вузькому екрані заголовок переноситься, а не обрізається
        self.subtitle = QLabel(subtitle)
        self.subtitle.setObjectName("PageSubtitle")
        self.subtitle.setWordWrap(True)
        box.addWidget(self.title)
        box.addWidget(self.subtitle)
        lay.addLayout(box, 1)
        self.actions = QHBoxLayout()
        self.actions.setSpacing(8)
        lay.addLayout(self.actions)

    def add_action(self, w):
        self.actions.addWidget(w)
        return w


class Badge(QLabel):
    """Кольоровий «значок»: level = ok / warn / err / info."""

    def __init__(self, text: str = "", level: str = "info", parent=None):
        super().__init__(text, parent)
        self.setObjectName("Badge")
        self.set_level(level)

    def set_level(self, level: str, text: str = None):
        if text is not None:
            self.setText(text)
        self.setProperty("level", level)
        repolish(self)


class Banner(QFrame):
    """Смуга з повідомленням (попередження, помилки, підказки)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("Banner")
        lay = QHBoxLayout(self)
        lay.setContentsMargins(14, 10, 14, 10)
        self.label = QLabel()
        self.label.setWordWrap(True)
        self.label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        lay.addWidget(self.label, 1)
        self.hide()

    def show_message(self, text: str, level: str = "info"):
        if not text:
            self.hide()
            return
        self.label.setText(text)
        self.setProperty("level", level)
        repolish(self)
        self.show()


def button(text: str, kind: str = "", tooltip: str = "", slot=None) -> QPushButton:
    b = QPushButton(text)
    if kind:
        b.setObjectName(kind)
    if tooltip:
        b.setToolTip(tooltip)
    if slot is not None:
        b.clicked.connect(slot)
    b.setCursor(Qt.CursorShape.PointingHandCursor)
    return b


def muted(text: str, hint: bool = False) -> QLabel:
    lab = QLabel(text)
    lab.setObjectName("Hint" if hint else "Muted")
    lab.setWordWrap(True)
    return lab


def stretch_widget() -> QWidget:
    w = QWidget()
    w.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
    return w


class NumberEdit(QLineEdit):
    """Поле для числа: приймає кому або крапку; некоректне значення підсвічується червоним."""
    valueChanged = Signal()

    def __init__(self, value=None, allow_empty: bool = False, minimum=None, parent=None):
        super().__init__(parent)
        self.allow_empty = allow_empty
        self.minimum = minimum
        self.setValue(value)
        self.textChanged.connect(self._check)

    def value(self):
        try:
            return parse_number(self.text())
        except ValueError:
            return None

    def setValue(self, v):
        self.blockSignals(True)
        self.setText("" if v is None else fmt(v, 8))
        self.blockSignals(False)
        self._mark(False)

    def is_valid(self) -> bool:
        try:
            v = parse_number(self.text())
        except ValueError:
            return False
        if v is None:
            return self.allow_empty
        return self.minimum is None or v >= self.minimum

    def _mark(self, bad: bool):
        self.setProperty("invalid", "true" if bad else "false")
        repolish(self)

    def _check(self):
        self._mark(not self.is_valid())
        self.valueChanged.emit()


def ro_item(text: str = "", align_right: bool = True) -> QTableWidgetItem:
    """Комірка таблиці лише для читання."""
    it = QTableWidgetItem(text)
    it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
    if align_right:
        it.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
    return it


class Worker(QThread):
    """
    Виконує функцію fn(progress_cb) у фоновому потоці.
    Сигнали: progress(done, total), done(результат), failed(текст помилки).
    """
    progress = Signal(int, int)
    done = Signal(object)
    failed = Signal(str)

    def __init__(self, fn, parent=None):
        super().__init__(parent)
        self.fn = fn

    def run(self):
        try:
            res = self.fn(lambda d, t: self.progress.emit(int(d), int(t)))
            self.done.emit(res)
        except Exception as e:                     # помилку показуємо у вікні, а не «падаємо»
            self.failed.emit(f"{e}\n\n{traceback.format_exc()}")
