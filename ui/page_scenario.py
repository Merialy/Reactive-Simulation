"""
Сторінка «Сценарій»: сценарій з довільною кількістю типів подій і сценаріями відмов.
Сценарій зберігається у JSON-файлі (папка scenarios) і моделюється ядром simulation/engine.py.
«Коротко про сценарій» (завантаження, відмови, попередження) оновлюється після кожної зміни.
"""
import glob
import os

from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox, QFileDialog, QFormLayout, QFrame, QGridLayout,
                               QHBoxLayout, QHeaderView, QLabel, QLineEdit, QScrollArea, QSpinBox,
                               QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget)

from simulation.model import (BG_CONTINUOUS, BG_NONE, BG_PERIODIC, BG_POISSON, HRT, OVERRUN, PERIODIC, POISSON, SRT,
                              STORM, Background, EventType, Fault, Platform, Scenario)
from ui.core import comma, fmt, parse_number
from html import escape

from ui.explain import scenario_facts
from ui.widgets import Banner, Card, Fact, NumberEdit, PageHeader, button, color, muted, pad_spin
from ui.i18n import tr, tr_msg

SCEN_DIR = "scenarios"

# Стовпці таблиці типів подій
T_NAME, T_CLS, T_LAW, T_RATE, T_WCET, T_DL, T_PHASE, T_WH = range(8)
T_HEAD = [tr("Назва"), tr("Клас"), tr("Як надходять"), tr("Інтенсивність, 1/мс"), tr("WCET, мс"), tr("Дедлайн, мс"),
          tr("Фаза, мс"), tr("m з n")]
T_TIPS = [tr("Назва типу подій (унікальна)"),
          tr("ЖРЧ – жорсткий реальний час (вищий пріоритет), МРЧ – м'який"),
          tr("Закон надходження: за замовчуванням ЖРЧ – періодичний, МРЧ – пуассонівський"),
          tr("Скільки подій надходить за 1 мс (для періодичних – 1/період)"),
          tr("Тривалість вторинної обробки події (без T_core)"),
          tr("Допустимий час реакції на подію"),
          tr("Момент першої періодичної події (порожньо – через один період)"),
          tr("Обмеження weakly hard (m, n): з будь-яких n подій поспіль хоча б m мають бути оброблені вчасно.\n"
             "Наприклад, «4 з 5» – допускається не більше 1 пропуску дедлайну на 5 подій. Порожньо – без обмеження.")]


def parse_wh(text: str):
    """«4 з 5», «4/5», «4 of 5», «(4, 5)» → (4, 5); порожньо → None. Помилка – ValueError."""
    import re
    t = (text or "").strip()
    if not t or t == "—":
        return None
    nums = re.findall(r"\d+", t)
    if len(nums) != 2 or re.search(r"[.,]\d", t.replace(", ", " ")):
        raise ValueError(t)
    return int(nums[0]), int(nums[1])


def wh_text(m, n) -> str:
    return "" if m is None or n is None else tr("{m} з {n}", m=m, n=n)

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
        self.header = PageHeader(tr("Типи подій і відмови"),
                                 tr("Виберіть сценарій → за потреби змініть числа → «Запустити»."))
        self.combo = QComboBox()
        self.combo.setMinimumWidth(260)
        self.combo.setMaximumWidth(380)
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

        # ---------------- Пояснення простими словами ----------------
        self.story_card = Card(tr("Коротко про сценарій"))
        self.story_desc = muted("")
        self.story_card.add(self.story_desc)
        facts = QHBoxLayout()
        facts.setSpacing(12)
        self.facts = [Fact() for _ in range(5)]          # 5-та плитка – фонові задачі (якщо є)
        for f in self.facts:
            facts.addWidget(f, 1)
        self.story_card.add(facts)
        self.alerts = QLabel()
        self.alerts.setWordWrap(True)
        self.alerts.setTextFormat(Qt.TextFormat.RichText)
        self.alerts.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.story_card.add(self.alerts)
        self.hint = muted("", hint=True)
        self.story_card.add(self.hint)
        lay.addWidget(self.story_card)

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
        pad_spin(self.n_spin)
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
        # ---------------- Фонові задачі ----------------
        bcard = Card(tr("Фонові задачі"),
                     tr("Робота з найнижчим пріоритетом (запис логів, самотестування пам'яті): виконується, лише коли "
                        "немає подій, і переривається будь-якою подією. Затримує події лише через перемикання "
                        "контексту t_cs і критичні секції з вимкненими перериваннями."))
        bg = QGridLayout()
        bg.setHorizontalSpacing(10)
        self.bg_mode = QComboBox()
        for label, mode in ((tr("немає"), BG_NONE), (tr("періодично"), BG_PERIODIC), (tr("випадково"), BG_POISSON),
                            (tr("постійно (процесор ніколи не простоює)"), BG_CONTINUOUS)):
            self.bg_mode.addItem(label, mode)
        self.bg_mode.setToolTip(tr("• періодично – кожні «інтервал» мс з'являється робота тривалістю «робота» мс;\n"
                                   "• випадково – робота з'являється в середньому раз на «інтервал» мс;\n"
                                   "• постійно – фонова робота є завжди (найгірший випадок)."))
        self.bg_interval = NumberEdit(100.0, minimum=0)
        self.bg_work = NumberEdit(1.0, minimum=0)
        self.bg_deadline = NumberEdit(None, allow_empty=True, minimum=0)
        self.bg_tcs = NumberEdit(0.0, minimum=0)
        self.bg_crit = NumberEdit(0.0, minimum=0)
        self.bg_share = NumberEdit(0.0, minimum=0)
        fields = [(tr("Режим"), self.bg_mode, None),
                  (tr("Інтервал, мс"), self.bg_interval, tr("Період (періодично) або середній інтервал (випадково) між фоновими роботами")),
                  (tr("Робота, мс"), self.bg_work, tr("Скільки процесорного часу потребує одна фонова робота")),
                  (tr("Дедлайн, мс"), self.bg_deadline, tr("Допустимий час виконання однієї фонової роботи (порожньо – без дедлайну)")),
                  (tr("Перемикання t_cs, мс"), self.bg_tcs, tr("Час перемикання контексту, коли подія витісняє фонову задачу, і при поверненні до неї")),
                  (tr("Критична секція t_id, мс"), self.bg_crit, tr("Тривалість однієї ділянки фонової задачі з вимкненими перериваннями")),
                  (tr("Частка в критичних секціях"), self.bg_share, tr("Яку частку свого часу фонова задача працює з вимкненими перериваннями (0…1)"))]
        for i, (lab, w, tip) in enumerate(fields):
            r, c = i % 4, (i // 4) * 2
            ql = QLabel(lab)
            if tip:
                ql.setToolTip(tip); w.setToolTip(tip)
            bg.addWidget(ql, r, c)
            bg.addWidget(w, r, c + 1)
        bg.setColumnStretch(1, 1)
        bg.setColumnStretch(3, 1)
        bcard.add(bg)
        self.bg_mode.currentIndexChanged.connect(self._bg_mode_changed)
        for w in (self.bg_interval, self.bg_work, self.bg_deadline, self.bg_tcs, self.bg_crit, self.bg_share):
            w.valueChanged.connect(self._changed)
        lay.addWidget(bcard)

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
        self.combo.addItem(tr("— виберіть готовий сценарій —"), None)
        normal, faulty = [], []
        for p in sorted(glob.glob(os.path.join(SCEN_DIR, "*.json"))):
            try:
                sc = Scenario.load(p)
                item = (tr(sc.name), p, tr(sc.description))
                (faulty if sc.faults else normal).append(item)
            except Exception:
                normal.append((os.path.basename(p), p, ""))
        for title, items in ((tr("Без відмов"), normal), (tr("З відмовами (щоб порівняти з нормальною роботою)"), faulty)):
            if not items:
                continue
            self.combo.addItem(title, None)
            head = self.combo.count() - 1
            try:
                it = self.combo.model().item(head)
                it.setEnabled(False)
                f = it.font()
                f.setBold(True)
                it.setFont(f)
            except Exception:
                pass
            for name, p, desc in items:
                self.combo.addItem("   " + name, p)
                if desc:
                    self.combo.setItemData(self.combo.count() - 1, desc, Qt.ItemDataRole.ToolTipRole)
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
        self.banner.hide()
        return True

    def set_scenario(self, sc: Scenario):
        self._loading = True
        self.name.setText(sc.name)
        self.desc.setText(sc.description)
        self.name.setCursorPosition(0)          # показувати початок довгого тексту, а не кінець
        self.desc.setCursorPosition(0)
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
        b = sc.background or Background()
        self.bg_mode.blockSignals(True)
        self.bg_mode.setCurrentIndex(max(0, self.bg_mode.findData(b.mode)))
        self.bg_mode.blockSignals(False)
        self.bg_interval.setValue(b.interval_ms)
        self.bg_work.setValue(b.work_ms)
        self.bg_deadline.setValue(b.deadline_ms)
        self.bg_tcs.setValue(b.t_cs_ms)
        self.bg_crit.setValue(b.crit_ms)
        self.bg_share.setValue(b.crit_share)
        self._bg_mode_changed(emit=False)
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
        wh = QTableWidgetItem(wh_text(et.wh_m, et.wh_n))
        wh.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        t.setItem(r, T_WH, wh)

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
            try:
                wh = parse_wh(self._text(self.types, r, T_WH))
            except ValueError:
                errors.append(tr("«{n}»: m з n – введіть два цілі числа, наприклад «4 з 5»", n=name))
                wh = None
            ets.append(EventType(name=name, cls=cls, rate=rate if rate is not None else 0.0,
                                 deadline_ms=dl if dl is not None else 0.0, wcet_ms=wcet, arrival=law, phase_ms=ph,
                                 wh_m=wh[0] if wh else None, wh_n=wh[1] if wh else None))
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
        mode = self.bg_mode.currentData()
        background = None
        if mode != BG_NONE:
            for w, lab in ((self.bg_interval, tr("Інтервал")), (self.bg_work, tr("Робота")),
                           (self.bg_deadline, tr("Дедлайн")), (self.bg_tcs, "t_cs"), (self.bg_crit, "t_id"),
                           (self.bg_share, tr("Частка в критичних секціях"))):
                if w.isEnabled() and not w.is_valid():
                    errors.append(tr("Фонові задачі – {w}: некоректне число", w=lab))
            background = Background(mode=mode, interval_ms=self.bg_interval.value() or 0.0,
                                    work_ms=self.bg_work.value() or 0.0, deadline_ms=self.bg_deadline.value(),
                                    t_cs_ms=self.bg_tcs.value() or 0.0, crit_ms=self.bg_crit.value() or 0.0,
                                    crit_share=self.bg_share.value() or 0.0)
        sc = Scenario(name=self.name.text().strip() or tr("Сценарій"), platform=plat, event_types=ets,
                      N=int(self.n_spin.value()), warmup_fraction=self.warmup.value() or 0.05,
                      description=self.desc.text().strip(), faults=faults, background=background)
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

    def _bg_mode_changed(self, *args, emit=True):
        """Поля фонових задач доступні лише для вибраного режиму."""
        mode = self.bg_mode.currentData()
        on = mode != BG_NONE
        jobs = mode in (BG_PERIODIC, BG_POISSON)
        for w in (self.bg_interval, self.bg_work, self.bg_deadline):
            w.setEnabled(jobs)
        for w in (self.bg_tcs, self.bg_crit, self.bg_share):
            w.setEnabled(on)
        if emit:
            self._changed()

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
        desc = self.desc.text().strip()
        self.story_desc.setText(tr(desc))
        self.story_desc.setVisible(bool(desc))
        try:
            self.p_info.setText(tr("T_IRQ = {a} мс,  T_core = {b} мс", a=fmt(sc.s1()), b=fmt(sc.platform.core_ms())))
        except Exception:
            self.p_info.setText(tr("Задайте T_IRQ (у мс) або F_clk і C_IRQ."))
        if form_errors:
            info = {"facts": [], "alerts": [("err", tr_msg(e)) for e in form_errors], "hint": ""}
        else:
            try:
                info = scenario_facts(sc)
            except Exception as e:
                info = {"facts": [], "alerts": [("err", str(e))], "hint": ""}
        for i, f in enumerate(self.facts):
            if i < len(info["facts"]):
                f.set(*info["facts"][i])
                f.show()
            else:
                f.hide()
        mark = {"ok": "✔", "warn": "⚠", "err": "✖"}
        self.alerts.setText(comma("<br>".join(f"<span style='color:{color(l).name()}'>{mark[l]} {escape(t)}</span>"
                                              for l, t in info["alerts"])))
        self.hint.setText(info["hint"])
        self.hint.setVisible(bool(info["hint"]))
        self.run_btn.setEnabled(not any(l == "err" for l, _ in info["alerts"]))

    def current(self):
        """Сценарій для запуску: лише з увімкненими відмовами."""
        sc, errors = self.build_scenario()
        return (sc.without_disabled_faults() if not errors else None), errors
