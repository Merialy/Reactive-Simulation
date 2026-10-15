"""
Сторінка «Сценарій»: сценарій з довільною кількістю типів подій і сценаріями відмов.
Сценарій зберігається у JSON-файлі (папка scenarios) і моделюється ядром simulation/engine.py.
Підсумок (U, попередження) оновлюється після кожної зміни.
"""
import copy
import glob
import os

from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox, QFileDialog, QFormLayout, QFrame, QGridLayout,
                               QHBoxLayout, QHeaderView, QLabel, QLineEdit, QScrollArea, QSpinBox,
                               QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget)

from simulation.model import (HRT, OVERRUN, PERIODIC, POISSON, SRT, STORM, EventType, Fault, Platform,
                              Scenario)
from ui.core import comma, fmt, parse_number
from ui.widgets import Badge, Banner, Card, NumberEdit, PageHeader, button, color, muted
from ui.i18n import tr, tr_msg

SCEN_DIR = "scenarios"

# Стовпці таблиці типів подій
T_NAME, T_CLS, T_LAW, T_RATE, T_WCET, T_DL, T_PHASE = range(7)
T_HEAD = [tr("Назва"), tr("Клас"), tr("Як надходять"), tr("Інтенсивність, 1/мс"), tr("WCET, мс"), tr("Дедлайн, мс"), tr("Фаза, мс")]
T_TIPS = [tr("Назва типу подій (унікальна)"),
          tr("ЖРЧ – жорсткий реальний час (вищий пріоритет), МРЧ – м'який"),
          tr("Закон надходження: за замовчуванням ЖРЧ – періодичний, МРЧ – пуассонівський"),
          tr("Скільки подій надходить за 1 мс (для періодичних – 1/період)"),
          tr("Тривалість вторинної обробки події (без T_core)"),
          tr("Допустимий час реакції на подію"),
          tr("Момент першої періодичної події (порожньо – через один період)")]

# Стовпці таблиці відмов
F_ON, F_KIND, F_NAME, F_START, F_DUR, F_RATE, F_FWCET, F_TARGET, F_PROB, F_FACTOR = range(10)
F_HEAD = [tr("Увімк."), tr("Вид відмови"), tr("Назва"), tr("Початок, мс"), tr("Тривалість, мс"), tr("Шторм: переривань/мс"),
          tr("Шторм: обробка, мс"), tr("WCET: тип подій"), tr("WCET: імовірність"), tr("WCET: множник")]
F_TIPS = [tr("Вимкнена відмова зберігається у файлі, але не моделюється"),
          tr("Шторм переривань – хибні переривання; Перевищення WCET – обробка триває довше"),
          tr("Довільна назва для звіту"),
          tr("Момент модельного часу, з якого діє відмова"),
          tr("Скільки діє відмова (порожньо – до кінця моделювання)"),
          tr("Інтенсивність хибних переривань під час шторму"),
          tr("Вторинна обробка одного хибного переривання на рівні МРЧ (0 – лише ISR)"),
          tr("Тип подій, обробка яких затягується (або «усі типи»)"),
          tr("Частка подій з перевищенням (0…1)"),
          tr("У скільки разів обробка довша за WCET")]
KIND_LABEL = {STORM: tr("Шторм переривань"), OVERRUN: tr("Перевищення WCET")}


def _num_item(v, digits=8) -> QTableWidgetItem:
    it = QTableWidgetItem("" if v is None else fmt(v, digits))
    it.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
    return it


class ScenarioPage(QWidget):
    run_requested = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("Page")
        self.path = None
        self._loading = False
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.setInterval(250)
        self._timer.timeout.connect(self.update_summary)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 20)
        outer.setSpacing(14)
        self.header = PageHeader(tr("Кілька типів подій і відмови"),
                                 tr("Нова можливість (побажання керівника): не 2 типи подій, як в Excel, а скільки "
                                 "завгодно, і сценарії відмов. Для роботи з Excel-файлами цей розділ не потрібен."))
        self.combo = QComboBox()
        self.combo.setMinimumWidth(200)
        self.combo.setMaximumWidth(340)
        self.combo.setToolTip(tr("Готові сценарії з папки scenarios"))
        self.combo.activated.connect(self._on_combo)
        self.header.add_action(self.combo)
        self.header.add_action(button(tr("Відкрити…"), slot=self.open_dialog, icon="folder"))
        self.header.add_action(button(tr("Зберегти як…"), slot=self.save_dialog, icon="save"))
        self.header.add_action(button(tr("Новий"), slot=self.new_scenario, icon="new"))
        self.run_btn = self.header.add_action(button(tr("Запустити"), "Primary", slot=self._request_run, icon="play"))
        outer.addWidget(self.header)

        self.banner = Banner()
        outer.addWidget(self.banner)
        self.howto = Banner()
        self.howto.show_message(
            tr("<b>Як користуватися:</b> 1) угорі праворуч виберіть готовий сценарій (або «Новий»); "
            "2) за потреби змініть числа в таблиці «Типи подій» і додайте відмову кнопкою «Шторм переривань» "
            "або «Перевищення WCET»; 3) натисніть «Запустити» – відкриються «Результати», далі «Графіки». "
            "Наведіть мишу на заголовок стовпця – з'явиться пояснення."), "info")
        self.howto.label.setTextFormat(Qt.TextFormat.RichText)
        outer.addWidget(self.howto)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        body = QWidget()
        body.setObjectName("Page")
        scroll.setWidget(body)
        outer.addWidget(scroll, 1)
        lay = QVBoxLayout(body)
        lay.setContentsMargins(0, 0, 8, 0)
        lay.setSpacing(14)

        # ---------------- Загальне і платформа ----------------
        top = QHBoxLayout()
        top.setSpacing(14)
        gen = Card(tr("Загальне"))
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignLeft)
        self.name = QLineEdit()
        self.desc = QLineEdit()
        self.desc.setPlaceholderText(tr("Короткий опис: звідки взято значення"))
        self.n_spin = QSpinBox()
        self.n_spin.setRange(100, 5_000_000)
        self.n_spin.setSingleStep(1000)
        self.n_spin.setGroupSeparatorShown(True)
        self.n_spin.setToolTip(tr("Кількість подій МРЧ, що моделюються (якщо МРЧ немає – усіх подій)"))
        self.warmup = NumberEdit(0.05, minimum=0)
        self.warmup.setToolTip(tr("Частка подій, що відкидається на початку і в кінці (стаціонарність)"))
        form.addRow(tr("Назва"), self.name)
        form.addRow(tr("Опис"), self.desc)
        form.addRow(tr("N (подій МРЧ)"), self.n_spin)
        form.addRow(tr("Частка розгону"), self.warmup)
        gen.add(form)
        top.addWidget(gen, 3)

        plat = Card(tr("Платформа"), tr("Задайте частоту і такти – час буде перераховано за T = C / F. "
                                 "Або залиште такти порожніми і задайте час у мс."))
        g = QGridLayout()
        g.setHorizontalSpacing(10)
        self.p_name = QLineEdit()
        self.p_f = NumberEdit(None, allow_empty=True, minimum=0)
        self.p_cirq = NumberEdit(None, allow_empty=True, minimum=0)
        self.p_ccore = NumberEdit(None, allow_empty=True, minimum=0)
        self.p_tirq = NumberEdit(None, allow_empty=True, minimum=0)
        self.p_tcore = NumberEdit(0.0, allow_empty=True, minimum=0)
        rows = [(tr("Назва"), self.p_name, None),
                (tr("F_clk, МГц"), self.p_f, tr("Тактова частота процесора")),
                (tr("C_IRQ, тактів"), self.p_cirq, tr("Такти первинного обробника (ISR)")),
                (tr("C_core, тактів"), self.p_ccore, tr("Такти накладних витрат ОС на одну подію")),
                (tr("T_IRQ, мс"), self.p_tirq, tr("Час ISR напряму (якщо такти не задано)")),
                (tr("T_core, мс"), self.p_tcore, tr("Накладні витрати ОС напряму (якщо такти не задано)"))]
        for i, (lab, w, tip) in enumerate(rows):
            r, c = i % 3, (i // 3) * 2
            ql = QLabel(lab)
            if tip:
                ql.setToolTip(tip)
                w.setToolTip(tip)
            g.addWidget(ql, r, c)
            g.addWidget(w, r, c + 1)
        plat.add(g)
        self.p_info = muted("", hint=True)
        plat.add(self.p_info)
        top.addWidget(plat, 4)
        lay.addLayout(top)

        # ---------------- Типи подій ----------------
        tcard = Card(tr("Типи подій"), tr("ЖРЧ обробляються з вищим пріоритетом (черга QH), МРЧ – з нижчим (QS). "
                                   "Будь-яке нове переривання перериває вторинну обробку."))
        tcard.head_actions.addWidget(button(tr("Тип подій"), slot=self.add_type, icon="plus"))
        tcard.head_actions.addWidget(button(tr("Видалити"), "Danger", slot=self.remove_type, icon="trash"))
        self.types = QTableWidget(0, len(T_HEAD))
        self.types.setHorizontalHeaderLabels(T_HEAD)
        for i, tip in enumerate(T_TIPS):
            self.types.horizontalHeaderItem(i).setToolTip(tip)
        self.types.horizontalHeader().setSectionResizeMode(T_NAME, QHeaderView.ResizeMode.Stretch)
        for i in range(1, len(T_HEAD)):
            self.types.horizontalHeader().setSectionResizeMode(i, QHeaderView.ResizeMode.ResizeToContents)
        self.types.verticalHeader().setVisible(False)
        self.types.verticalHeader().setDefaultSectionSize(36)
        self.types.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.types.setAlternatingRowColors(True)
        self.types.itemChanged.connect(self._changed)
        tcard.add(self.types)
        lay.addWidget(tcard)

        # ---------------- Відмови ----------------
        fcard = Card(tr("Відмови"), tr("Що станеться, якщо щось піде не так. Шторм переривань – хибні переривання "
                                "(наприклад, від несправної лінії). Перевищення WCET – обробка триває довше, "
                                "ніж заплановано. Відмова діє лише на заданому інтервалі часу."))
        fcard.head_actions.addWidget(button(tr("Шторм переривань"), slot=lambda: self.add_fault(STORM), icon="bolt"))
        fcard.head_actions.addWidget(button(tr("Перевищення WCET"), slot=lambda: self.add_fault(OVERRUN), icon="history"))
        fcard.head_actions.addWidget(button(tr("Видалити"), "Danger", slot=self.remove_fault, icon="trash"))
        self.faults = QTableWidget(0, len(F_HEAD))
        self.faults.setHorizontalHeaderLabels(F_HEAD)
        for i, tip in enumerate(F_TIPS):
            self.faults.horizontalHeaderItem(i).setToolTip(tip)
        self.faults.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.faults.horizontalHeader().setSectionResizeMode(F_NAME, QHeaderView.ResizeMode.Interactive)
        self.faults.horizontalHeader().resizeSection(F_NAME, 200)
        self.faults.verticalHeader().setVisible(False)
        self.faults.verticalHeader().setDefaultSectionSize(36)
        self.faults.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.faults.itemChanged.connect(self._changed)
        fcard.add(self.faults)
        self.faults_empty = muted(tr("Відмов немає – система працює в нормальному режимі."), hint=True)
        fcard.add(self.faults_empty)
        lay.addWidget(fcard)

        # ---------------- Підсумок ----------------
        scard = Card(tr("Підсумок перед запуском"))
        row = QHBoxLayout()
        row.setSpacing(18)
        self.u_badge = Badge("U = –", "info")
        self.uf_badge = Badge("", "info")
        self.s1_label = muted("")
        row.addWidget(QLabel(tr("Завантаження процесора:")))
        row.addWidget(self.u_badge)
        row.addWidget(self.uf_badge)
        row.addSpacing(12)
        row.addWidget(self.s1_label, 1)
        scard.add(row)
        self.checks = muted("")
        self.checks.setTextFormat(Qt.TextFormat.RichText)
        scard.add(self.checks)
        lay.addWidget(scard)
        lay.addStretch(1)

        for w in (self.name, self.desc, self.p_name):
            w.textChanged.connect(self._changed)
        for w in (self.p_f, self.p_cirq, self.p_ccore, self.p_tirq, self.p_tcore, self.warmup):
            w.valueChanged.connect(self._changed)
        self.n_spin.valueChanged.connect(self._changed)

        self.reload_list()
        default = os.path.join(SCEN_DIR, "boryviter_adcs.json")
        if os.path.exists(default):
            self.load(default)
        else:
            self.new_scenario()

    # ==================================================================
    # Список готових сценаріїв
    # ==================================================================
    def reload_list(self):
        self.combo.blockSignals(True)
        self.combo.clear()
        self.combo.addItem(tr("— готові сценарії —"), None)
        for p in sorted(glob.glob(os.path.join(SCEN_DIR, "*.json"))):
            try:
                name = Scenario.load(p).name
            except Exception:
                name = os.path.basename(p)
            self.combo.addItem(name, p)
        self.combo.blockSignals(False)

    def _on_combo(self, i):
        p = self.combo.itemData(i)
        if p:
            self.load(p)

    # ==================================================================
    # Сценарій ↔ форма
    # ==================================================================
    def load(self, path: str):
        try:
            sc = Scenario.load(path)
        except Exception as e:
            self.banner.show_message(tr("Не вдалося відкрити сценарій: {e}", e=e), "err")
            return False
        self.path = path
        self.set_scenario(sc)
        idx = self.combo.findData(path)
        self.combo.setCurrentIndex(max(0, idx))
        self.banner.show_message(tr("Відкрито: {f}", f=os.path.basename(path)), "ok")
        return True

    def set_scenario(self, sc: Scenario):
        self._loading = True
        self.name.setText(sc.name)
        self.desc.setText(sc.description)
        self.n_spin.setValue(int(sc.N))
        self.warmup.setValue(sc.warmup_fraction)
        p = sc.platform
        self.p_name.setText(p.name)
        self.p_f.setValue(p.f_clk_mhz)
        self.p_cirq.setValue(p.c_irq)
        self.p_ccore.setValue(p.c_core)
        self.p_tirq.setValue(p.t_irq_ms)
        self.p_tcore.setValue(p.t_core_ms)
        self.types.setRowCount(0)
        for et in sc.event_types:
            self._append_type(et, sc.platform)
        self.faults.setRowCount(0)
        for f in sc.faults:
            self._append_fault(f)
        self._loading = False
        self._refresh_targets()
        self.update_summary()

    def _append_type(self, et: EventType, platform: Platform = None):
        t = self.types
        r = t.rowCount()
        t.insertRow(r)
        t.setItem(r, T_NAME, QTableWidgetItem(et.name))
        cls = QComboBox()
        cls.addItem(tr("ЖРЧ"), HRT)
        cls.addItem(tr("МРЧ"), SRT)
        cls.setToolTip(tr("ЖРЧ – жорсткий реальний час: пропуск дедлайну неприпустимий, обробляється першою.\n"
                       "МРЧ – м'який реальний час: запізнення небажане, але допустиме."))
        cls.setCurrentIndex(0 if et.cls == HRT else 1)
        cls.currentIndexChanged.connect(self._changed)
        t.setCellWidget(r, T_CLS, cls)
        law = QComboBox()
        law.addItem(tr("автоматично"), None)
        law.addItem(tr("періодично"), PERIODIC)
        law.addItem(tr("випадково"), POISSON)
        law.setToolTip(tr("Як надходять події цього типу:\n"
                       "• автоматично – ЖРЧ періодично, МРЧ випадково (так у більшості задач);\n"
                       "• періодично – рівно через 1/інтенсивність мс (датчики, таймери);\n"
                       "• випадково – потік Пуассона, інтервали випадкові (команди, повідомлення)."))
        law.setCurrentIndex({None: 0, PERIODIC: 1, POISSON: 2}.get(et.arrival, 0))
        law.currentIndexChanged.connect(self._changed)
        t.setCellWidget(r, T_LAW, law)
        t.setItem(r, T_RATE, _num_item(et.rate))
        wcet = et.wcet_ms
        if wcet is None and et.wcet_cycles is not None and platform is not None and platform.f_clk_mhz:
            wcet = et.wcet_cycles / (platform.f_clk_mhz * 1000.0)
        t.setItem(r, T_WCET, _num_item(wcet))
        t.setItem(r, T_DL, _num_item(et.deadline_ms))
        t.setItem(r, T_PHASE, _num_item(et.phase_ms))

    def _append_fault(self, f: Fault):
        t = self.faults
        r = t.rowCount()
        t.insertRow(r)
        on = QCheckBox()
        on.setChecked(f.enabled)
        on.toggled.connect(self._changed)
        box = QWidget()
        bl = QHBoxLayout(box)
        bl.setContentsMargins(10, 0, 0, 0)
        bl.addWidget(on)
        t.setCellWidget(r, F_ON, box)
        kind = QTableWidgetItem(KIND_LABEL.get(f.kind, f.kind))
        kind.setData(Qt.ItemDataRole.UserRole, f.kind)
        kind.setFlags(kind.flags() & ~Qt.ItemFlag.ItemIsEditable)
        kind.setForeground(color("hrt") if f.kind == STORM else color("warn"))
        t.setItem(r, F_KIND, kind)
        t.setItem(r, F_NAME, QTableWidgetItem(f.name))
        t.setItem(r, F_START, _num_item(f.start_ms))
        t.setItem(r, F_DUR, _num_item(f.duration_ms))
        storm = f.kind == STORM
        for col, val, active in ((F_RATE, f.rate, storm), (F_FWCET, f.wcet_ms, storm),
                                 (F_PROB, f.probability, not storm), (F_FACTOR, f.factor, not storm)):
            it = _num_item(val if active else None)
            if not active:
                it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable & ~Qt.ItemFlag.ItemIsEnabled)
                it.setText("—")
                it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            t.setItem(r, col, it)
        if storm:
            it = QTableWidgetItem("—")
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable & ~Qt.ItemFlag.ItemIsEnabled)
            it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            t.setItem(r, F_TARGET, it)
        else:
            tgt = QComboBox()
            tgt.setProperty("target", f.target)
            tgt.currentIndexChanged.connect(self._changed)
            t.setCellWidget(r, F_TARGET, tgt)
        self.faults_empty.setVisible(t.rowCount() == 0)

    @staticmethod
    def _fit(t: QTableWidget, min_rows: int = 2):
        """Висота таблиці – під усі рядки, щоб не прокручувати маленьку таблицю всередині сторінки."""
        rows = max(t.rowCount(), min_rows)
        h = t.horizontalHeader().height() + rows * t.verticalHeader().defaultSectionSize() + 22
        t.setMinimumHeight(h)
        t.setMaximumHeight(h)

    def _refresh_targets(self):
        """Оновити списки типів у комбобоксах «WCET: тип подій»."""
        self._fit(self.types)
        self._fit(self.faults, 1)
        names = [self._text(self.types, r, T_NAME) for r in range(self.types.rowCount())]
        for r in range(self.faults.rowCount()):
            cb = self.faults.cellWidget(r, F_TARGET)
            if not isinstance(cb, QComboBox):
                continue
            cur = cb.currentData() if cb.count() else cb.property("target")
            cb.blockSignals(True)
            cb.clear()
            cb.addItem(tr("усі типи"), None)
            for n in names:
                cb.addItem(n, n)
            i = cb.findData(cur) if cur else 0
            cb.setCurrentIndex(i if i >= 0 else 0)
            cb.blockSignals(False)
        self.faults_empty.setVisible(self.faults.rowCount() == 0)

    @staticmethod
    def _text(t: QTableWidget, r: int, c: int) -> str:
        it = t.item(r, c)
        return it.text().strip() if it else ""

    def _num(self, t: QTableWidget, r: int, c: int, what: str, errors: list, allow_empty=False):
        txt = self._text(t, r, c)
        if txt in ("", "—"):
            if not allow_empty:
                errors.append(tr("{w}: не заповнено", w=what))
            return None
        try:
            return parse_number(txt)
        except ValueError:
            errors.append(tr("{w}: «{t}» – не число", w=what, t=txt))
            return None

    def build_scenario(self):
        """Повертає (сценарій, список помилок форми)."""
        errors = []
        for w, lab in ((self.p_f, "F_clk"), (self.p_cirq, "C_IRQ"), (self.p_ccore, "C_core"),
                       (self.p_tirq, "T_IRQ"), (self.p_tcore, "T_core"), (self.warmup, tr("Частка розгону"))):
            if not w.is_valid():
                errors.append(tr("{w}: некоректне число", w=lab))
        plat = Platform(name=self.p_name.text().strip() or tr("БО"), f_clk_mhz=self.p_f.value() or None,
                        c_irq=self.p_cirq.value(), c_core=self.p_ccore.value(),
                        t_irq_ms=self.p_tirq.value(), t_core_ms=self.p_tcore.value() or 0.0)
        ets = []
        for r in range(self.types.rowCount()):
            name = self._text(self.types, r, T_NAME) or tr("Тип {n}", n=r + 1)
            cls = self.types.cellWidget(r, T_CLS).currentData()
            law = self.types.cellWidget(r, T_LAW).currentData()
            rate = self._num(self.types, r, T_RATE, tr("«{n}»: інтенсивність", n=name), errors)
            wcet = self._num(self.types, r, T_WCET, tr("«{n}»: WCET", n=name), errors)
            dl = self._num(self.types, r, T_DL, tr("«{n}»: дедлайн", n=name), errors)
            ph = self._num(self.types, r, T_PHASE, tr("«{n}»: фаза", n=name), errors, allow_empty=True)
            ets.append(EventType(name=name, cls=cls, rate=rate if rate is not None else 0.0,
                                 deadline_ms=dl if dl is not None else 0.0, wcet_ms=wcet, arrival=law, phase_ms=ph))
        faults = []
        for r in range(self.faults.rowCount()):
            kind = self.faults.item(r, F_KIND).data(Qt.ItemDataRole.UserRole)
            label = self._text(self.faults, r, F_NAME) or tr("Відмова {n}", n=r + 1)
            on = self.faults.cellWidget(r, F_ON).findChild(QCheckBox).isChecked()
            start = self._num(self.faults, r, F_START, tr("«{n}»: початок", n=label), errors)
            dur = self._num(self.faults, r, F_DUR, tr("«{n}»: тривалість", n=label), errors, allow_empty=True)
            f = Fault(kind=kind, start_ms=start or 0.0, duration_ms=dur, name=self._text(self.faults, r, F_NAME),
                      enabled=on)
            if kind == STORM:
                f.rate = self._num(self.faults, r, F_RATE, tr("«{n}»: інтенсивність шторму", n=label), errors) or 0.0
                f.wcet_ms = self._num(self.faults, r, F_FWCET, tr("«{n}»: обробка", n=label), errors, allow_empty=True) or 0.0
            else:
                cb = self.faults.cellWidget(r, F_TARGET)
                f.target = cb.currentData() if isinstance(cb, QComboBox) else None
                f.probability = self._num(self.faults, r, F_PROB, tr("«{n}»: імовірність", n=label), errors) or 0.0
                f.factor = self._num(self.faults, r, F_FACTOR, tr("«{n}»: множник", n=label), errors) or 0.0
            faults.append(f)
        sc = Scenario(name=self.name.text().strip() or tr("Сценарій"), platform=plat, event_types=ets,
                      N=int(self.n_spin.value()), warmup_fraction=self.warmup.value() or 0.05,
                      description=self.desc.text().strip(), faults=faults)
        return sc, errors

    # ==================================================================
    # Дії
    # ==================================================================
    def _changed(self, *args):
        if self._loading:
            return
        if self.sender() is self.types:
            self._refresh_targets()
        self._timer.start()

    def add_type(self):
        n = self.types.rowCount() + 1
        self._append_type(EventType(name=tr("Тип {n}", n=n), cls=SRT, rate=0.1, deadline_ms=1.0, wcet_ms=0.1))
        self._refresh_targets()
        self.types.setCurrentCell(self.types.rowCount() - 1, T_NAME)
        self.types.editItem(self.types.item(self.types.rowCount() - 1, T_NAME))
        self._timer.start()

    def remove_type(self):
        rows = sorted({i.row() for i in self.types.selectedIndexes()}, reverse=True)
        if not rows and self.types.currentRow() >= 0:
            rows = [self.types.currentRow()]
        for r in rows:
            self.types.removeRow(r)
        self._refresh_targets()
        self._timer.start()

    def add_fault(self, kind: str):
        sc, _ = self.build_scenario()
        start = 0.0
        if kind == STORM:
            f = Fault(kind=STORM, start_ms=start, duration_ms=1000.0, rate=5.0, wcet_ms=0.0, name=tr("Шторм переривань"))
        else:
            tgt = sc.event_types[0].name if sc.event_types else None
            f = Fault(kind=OVERRUN, start_ms=start, duration_ms=None, target=tgt, probability=0.2, factor=2.0,
                      name=tr("Перевищення WCET"))
        self._append_fault(f)
        self._refresh_targets()
        self._timer.start()

    def remove_fault(self):
        rows = sorted({i.row() for i in self.faults.selectedIndexes()}, reverse=True)
        if not rows and self.faults.currentRow() >= 0:
            rows = [self.faults.currentRow()]
        for r in rows:
            self.faults.removeRow(r)
        self._refresh_targets()
        self._timer.start()

    def new_scenario(self):
        self.path = None
        sc = Scenario(name=tr("Новий сценарій"), platform=Platform(name=tr("БО"), t_irq_ms=0.01),
                      event_types=[EventType(name=tr("МРЧ"), cls=SRT, rate=1.0, wcet_ms=0.5, deadline_ms=1.0),
                                   EventType(name=tr("ЖРЧ"), cls=HRT, rate=2.0, wcet_ms=0.1, deadline_ms=0.2)],
                      N=10000)
        self.set_scenario(sc)
        self.combo.setCurrentIndex(0)
        self.banner.show_message(tr("Створено новий сценарій з двома типами подій – змініть його під свою задачу."),
                                 "info")

    def open_dialog(self):
        path, _ = QFileDialog.getOpenFileName(self, tr("Відкрити сценарій"), os.path.abspath(SCEN_DIR),
                                              tr("Сценарій (*.json)"))
        if path:
            self.load(path)

    def save_dialog(self):
        sc, errors = self.build_scenario()
        if errors:
            self.banner.show_message(tr("Спочатку виправте помилки:\n• ") + "\n• ".join(errors), "err")
            return
        start = self.path or os.path.join(os.path.abspath(SCEN_DIR), tr("новий_сценарій.json"))
        path, _ = QFileDialog.getSaveFileName(self, tr("Зберегти сценарій"), start, tr("Сценарій (*.json)"))
        if not path:
            return
        if not path.lower().endswith(".json"):
            path += ".json"
        try:
            sc.save(path)
            self.path = path
            self.reload_list()
            i = self.combo.findData(path)
            if i < 0:
                i = self.combo.findData(os.path.relpath(path))
            self.combo.setCurrentIndex(max(0, i))
            self.banner.show_message(tr("Сценарій збережено: {p}", p=path), "ok")
        except Exception as e:
            self.banner.show_message(tr("Не вдалося зберегти: {e}", e=e), "err")

    def _request_run(self):
        sc, errors = self.build_scenario()
        if not errors:
            errors = sc.check()[0]
        if errors:
            self.banner.show_message(tr("Сценарій не можна запустити:\n• ") + "\n• ".join(errors), "err")
            return
        self.run_requested.emit()

    # ==================================================================
    # Підсумок
    # ==================================================================
    def update_summary(self):
        sc, form_errors = self.build_scenario()
        try:
            s1 = sc.s1()
            core = sc.platform.core_ms()
            self.p_info.setText(tr("T_IRQ = {a} мс,  T_core = {b} мс", a=fmt(s1), b=fmt(core)))
            self.s1_label.setText(tr("Первинна обробка T_IRQ = {a} мс", a=fmt(s1)))
        except Exception:
            self.p_info.setText(tr("Задайте T_IRQ (у мс) або F_clk і C_IRQ."))
            self.s1_label.setText("")
        errors, warns = ([], [])
        if not form_errors:
            errors, warns = sc.check()
        errors = form_errors + errors
        if errors:
            self.u_badge.set_level("err", "U = ?")
            self.uf_badge.hide()
        else:
            u = sc.utilization()
            lvl = "err" if u >= 1 else ("warn" if u > 0.95 else "ok")
            self.u_badge.set_level(lvl, f"U = {fmt(u, 3)}")
            if sc.active_faults():
                uf = sc.utilization_during_faults()
                lvlf = "err" if uf >= 1 else ("warn" if uf > 0.95 else "ok")
                self.uf_badge.set_level(lvlf, tr("під час відмов U = {u}", u=fmt(uf, 3)))
                self.uf_badge.show()
            else:
                self.uf_badge.hide()
        lines = [f"<span style='color:{color('err').name()}'>✖ {tr_msg(e)}</span>" for e in errors] + \
                [f"<span style='color:{color('warn').name()}'>⚠ {tr_msg(w)}</span>" for w in warns]
        if not lines:
            lines = [f"<span style='color:{color('ok').name()}'>✔ " + tr("Сценарій коректний, можна запускати.") + "</span>"]
        self.checks.setText(comma("<br>".join(lines)))
        self.run_btn.setEnabled(not errors)

    def current(self):
        """Сценарій для запуску: лише з увімкненими відмовами."""
        sc, errors = self.build_scenario()
        return (sc.without_disabled_faults() if not errors else None), errors
