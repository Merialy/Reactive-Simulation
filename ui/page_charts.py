"""
Сторінка «Графіки» (matplotlib усередині вікна Qt).

Серія з Excel: будь-який з 18 показників залежно від будь-якого параметра, з групуванням
за іншим параметром і довірчими інтервалами.
Сценарій: час реакції за типами подій; час реакції в часі (з виділеними інтервалами відмов);
ймовірність порушення дедлайнів у звичайному режимі і під час відмов.
"""
import os
from collections import defaultdict
from datetime import datetime

import numpy as np
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg, NavigationToolbar2QT
from matplotlib.figure import Figure
from PySide6.QtWidgets import (QCheckBox, QComboBox, QFileDialog, QHBoxLayout, QLabel, QVBoxLayout, QWidget)

from config import Config
from simulation.validation import utilization
from ui.core import METRIC_NAMES, METRIC_SHORT, PARAM_KEYS, PARAM_SHORT, fmt
from ui.theme import SERIES
from ui.widgets import Banner, Card, PageHeader, button, muted, PALETTE
from ui.i18n import tr, tr_msg
from ui.core import cls_name, type_label, comma

# Готові графіки для серії з Excel (як у версіях 1–3, але з довірчими інтервалами):
# (назва, параметр осі X, показники, підписи ліній, підпис осі Y, логарифмічна шкала за замовчуванням)
PRESETS = [
    (tr("Час очікування від тривалості вторинної обробки s2"), 2, [1, 2, 3],
     [tr("Первинна обробка (ISR)"), tr("Вторинна обробка, МРЧ"), tr("Вторинна обробка, ЖРЧ")], tr("Середній час очікування, мс"), True),
    (tr("Максимальний час очікування від s2"), 2, [6, 7, 8],
     [tr("Первинна обробка (ISR)"), tr("Вторинна обробка, МРЧ"), tr("Вторинна обробка, ЖРЧ")], tr("Максимальний час очікування, мс"), True),
    (tr("Довжина черг від s2"), 2, [11, 12, 13],
     [tr("Черга ISR (Q0)"), tr("Черга МРЧ (QS)"), tr("Черга ЖРЧ (QH)")], tr("Середня довжина черги"), True),
    (tr("Час перебування (реакції) від s2"), 2, [4, 5], [tr("МРЧ"), tr("ЖРЧ")], tr("Середній час перебування, мс"), True),
    (tr("Завантаження процесора від s2"), 2, [14, 15, 16],
     [tr("Первинна обробка (ISR)"), tr("Вторинна обробка, МРЧ"), tr("Вторинна обробка, ЖРЧ")], tr("Коефіцієнт завантаження"), False),
    (tr("Порушення дедлайнів МРЧ від дедлайну D1"), 4, [17], [tr("Ймовірність порушення дедлайну МРЧ")],
     tr("Ймовірність порушення дедлайну"), False),
    (tr("Порушення дедлайнів ЖРЧ від дедлайну D2"), 7, [18], [tr("Ймовірність порушення дедлайну ЖРЧ")],
     tr("Ймовірність порушення дедлайну"), False),
]
CUSTOM = "custom"

def plain_log(axis):
    """Підписи логарифмічної осі звичайними числами (0,001; 0,1; 10) замість 10⁻³."""
    from matplotlib.ticker import FuncFormatter, LogLocator, NullFormatter
    lo, hi = axis.get_view_interval()
    wide = lo > 0 and hi / lo > 300           # широкий діапазон – лише 0,001; 0,01; 0,1; 1…
    axis.set_major_locator(LogLocator(base=10, subs=(1.0,) if wide else (1.0, 2.0, 5.0)))
    axis.set_major_formatter(FuncFormatter(
        lambda v, _: comma(f"{v:.10f}".rstrip("0").rstrip(".") if v < 1 else f"{v:g}")))
    axis.set_minor_formatter(NullFormatter())


SC_CHARTS = [tr("Час реакції за типами подій"), tr("Час реакції в часі (один прогін)"), tr("Ймовірність порушення дедлайнів")]


class ChartsPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("Page")
        self.result = None
        self._busy = False

        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 20)
        outer.setSpacing(14)
        self.header = PageHeader(tr("Графіки"), tr("Спочатку запустіть серію або сценарій."))
        self.header.add_action(button(tr("Зберегти PNG…"), slot=self.save_png))
        outer.addWidget(self.header)
        self.banner = Banner()
        outer.addWidget(self.banner)

        card = Card()
        # Елементи керування – у два рядки, щоб поміщались і на невеликому екрані
        row1, ctl = QHBoxLayout(), QHBoxLayout()
        row1.setSpacing(10)
        ctl.setSpacing(10)
        self.lab1, self.cb1 = QLabel(tr("Показник")), QComboBox()
        self.lab2, self.cb2 = QLabel(tr("Вісь X")), QComboBox()
        self.lab3, self.cb3 = QLabel(tr("Показник")), QComboBox()     # лише для «Довільного графіка»
        self.cb1.setMinimumWidth(300)
        self.cb1.setMaxVisibleItems(20)
        row1.addWidget(self.lab1)
        row1.addWidget(self.cb1, 1)
        self.cb3.setMinimumWidth(260)
        for lab, cb in ((self.lab3, self.cb3), (self.lab2, self.cb2)):
            ctl.addWidget(lab)
            ctl.addWidget(cb)
        self.cb1.currentIndexChanged.connect(self._on_chart_type)
        for cb in (self.cb2, self.cb3):
            cb.currentIndexChanged.connect(self.draw)
        self.ci_box = QCheckBox(tr("Довірчі інтервали"))
        self.ci_box.setChecked(True)
        self.ci_box.toggled.connect(self.draw)
        self.log_box = QCheckBox(tr("Логарифмічна шкала"))
        self.log_box.toggled.connect(self.draw)
        ctl.addSpacing(12)
        ctl.addWidget(self.ci_box)
        ctl.addWidget(self.log_box)
        ctl.addStretch(1)
        card.add(row1)
        card.add(ctl)

        self.fig = Figure(figsize=(9, 5.2), dpi=100)
        self.canvas = FigureCanvasQTAgg(self.fig)
        self.canvas.setMinimumHeight(420)
        self.toolbar = self._make_toolbar()
        self.card = card
        card.add(self.toolbar)
        card.add(self.canvas, 1)
        self.note = muted("", hint=True)
        card.add(self.note)
        outer.addWidget(card, 1)
        self.set_result(None)

    # ------------------------------------------------------------------
    def set_result(self, result):
        self.result = result
        self._busy = True
        for cb in (self.cb1, self.cb2, self.cb3):
            cb.clear()
        kind = result["kind"] if result else None
        if kind == "excel":
            self.header.subtitle.setText(tr("Серія «{s}»", s=result["source"]))
            self.lab1.setText(tr("Графік"))
            varying = self._varying_params()
            for i, (title, xk, *_rest) in enumerate(PRESETS):
                if xk in varying:                      # графік має сенс, лише якщо параметр осі змінюється
                    self.cb1.addItem(title, i)
            self.cb1.addItem(tr("Довільний графік: будь-який показник від будь-якого параметра"), CUSTOM)
            for i, n in enumerate(METRIC_SHORT, 1):
                self.cb3.addItem(f"{i}. {n}", i)
            for k in sorted(varying):
                self.cb2.addItem(PARAM_SHORT[k], k)
            self.cb2.addItem(tr("№ сценарію"), -1)
            x = varying[0] if varying else -1
            self.cb2.setCurrentIndex(self.cb2.findData(x))
            self.cb3.setCurrentIndex(16 if x == 4 else 17 if x == 7 else 3)
            self.cb1.setCurrentIndex(0)
        elif kind == "scenario":
            sc = result["scenario"]
            self.header.subtitle.setText(tr("Сценарій «{s}»", s=sc.name))
            for n in SC_CHARTS:
                self.cb1.addItem(n)
            if sc.active_faults():
                self.cb1.setCurrentIndex(1)
            self.lab1.setText(tr("Графік"))
            self.log_box.setChecked(True)          # часи реакції різних типів відрізняються на порядки
            self._show_controls(True, False, False)
        else:
            self.header.subtitle.setText(tr("Спочатку запустіть серію або сценарій."))
            self._show_controls(False, False, False)
        self._busy = False
        if kind == "excel":
            self._on_chart_type()
        else:
            self.draw()

    def _on_chart_type(self, *_):
        """Зміна типу графіка: показати потрібні поля і встановити шкалу за замовчуванням."""
        if self.result is None or self.result["kind"] != "excel":
            self.draw()
            return
        d = self.cb1.currentData()
        custom = d == CUSTOM
        self._show_controls(True, custom, custom)
        busy, self._busy = self._busy, True
        self.log_box.setChecked(False if custom or d is None else PRESETS[d][5])
        self._busy = busy
        self.draw()

    def _show_controls(self, a, b, c):
        for w in (self.lab1, self.cb1):
            w.setVisible(a)
        for w in (self.lab2, self.cb2):
            w.setVisible(b)
        for w in (self.lab3, self.cb3):
            w.setVisible(c)

    def _varying_params(self):
        """Параметри, що змінюються між сценаріями серії (спочатку ті, що мають більше значень)."""
        runs = [p for p in self.result["run_list"] if p]
        counts = []
        for k, key in enumerate(PARAM_KEYS):
            n = len({round(p[key], 12) for p in runs})
            if n > 1:
                counts.append((n, k))
        counts.sort(key=lambda t: (-t[0], t[1]))
        return [k for _, k in counts]

    # ------------------------------------------------------------------
    def _style_axes(self, ax):
        fg, mut, grid = PALETTE.get("text", "#222"), PALETTE.get("muted", "#666"), PALETTE.get("grid", "#ddd")
        ax.set_facecolor(PALETTE.get("surface", "#fff"))
        for s in ("top", "right"):
            ax.spines[s].set_visible(False)
        for s in ("left", "bottom"):
            ax.spines[s].set_color(mut)
        ax.tick_params(colors=mut, labelsize=9)
        ax.xaxis.label.set_color(fg)
        ax.yaxis.label.set_color(fg)
        ax.title.set_color(fg)
        ax.grid(True, color=grid, linewidth=0.8, alpha=0.9)
        ax.set_axisbelow(True)

    def draw(self):
        if self._busy:
            return
        self.fig.clear()
        self.fig.set_facecolor(PALETTE.get("surface", "#fff"))
        self.note.setText("")
        r = self.result
        ax = self.fig.add_subplot(111)
        self._style_axes(ax)
        try:
            if r is None:
                ax.text(0.5, 0.5, tr("Немає даних для графіка"), ha="center", va="center",
                        color=PALETTE.get("muted", "#666"), transform=ax.transAxes, fontsize=12)
                ax.set_xticks([]); ax.set_yticks([])
            elif r["kind"] == "excel":
                self._draw_excel(ax)
            else:
                i = self.cb1.currentIndex()
                [self._draw_sc_types, self._draw_sc_time, self._draw_sc_miss][max(0, i)](ax)
            self.fig.tight_layout()
            self.banner.hide()
        except Exception as e:
            self.banner.show_message(tr("Не вдалося побудувати графік: {e}", e=e), "err")
        self.canvas.draw_idle()

    def _legend(self, ax, **kw):
        leg = ax.legend(frameon=False, fontsize=9, **kw)
        if leg:
            for t in leg.get_texts():
                t.set_color(PALETTE.get("text", "#222"))

    # ---------------- Серія з Excel ----------------
    def _draw_excel(self, ax):
        d = self.cb1.currentData()
        if d == CUSTOM or d is None:
            m = self.cb3.currentData() or 4
            xk = self.cb2.currentData()
            self._plot_lines(ax, xk, [m], [METRIC_SHORT[m - 1]],
                             METRIC_SHORT[m - 1] + (tr(", мс") if m <= 10 else ""), METRIC_NAMES[m - 1])
        else:
            title, xk, metrics, labels, ylabel, _ = PRESETS[d]
            self._plot_lines(ax, xk, metrics, labels, ylabel, title)

    @staticmethod
    def _irrelevant(metrics):
        """Параметри, що НЕ впливають на показники: дедлайни змінюють лише ймовірність їх порушення."""
        if all(m <= 16 for m in metrics):
            return {4, 7}
        if metrics == [17]:
            return {7}
        if metrics == [18]:
            return {4}
        return set()

    def _plot_lines(self, ax, xk, metrics, labels, ylabel, title):
        """
        Лінії показників залежно від параметра xk.
        Сценарії, що відрізняються лише параметрами, які на показник не впливають (дедлайни),
        мають однакові результати і зливаються в одну точку. Якщо змінюються ще й інші параметри
        (наприклад, s1), для кожного їх значення будується окрема лінія – різні сценарії не змішуються.
        """
        from matplotlib.ticker import FuncFormatter, NullFormatter
        r = self.result
        skip = self._irrelevant(metrics)
        others = [] if xk == -1 else sorted(k for k in self._varying_params() if k != xk and k not in skip)
        data = defaultdict(lambda: defaultdict(list))      # data[key][x] = [(mean{}, ci{}), ...]
        for c, (p, res) in enumerate(zip(r["run_list"], r["batch"])):
            if p is None or res is None:
                continue
            x = c + 1 if xk == -1 else p[PARAM_KEYS[xk]]
            key = tuple(round(p[PARAM_KEYS[k]], 12) for k in others)
            data[key][x].append((res, p))
        if not data:
            raise ValueError(tr("немає змодельованих сценаріїв"))
        log = self.log_box.isChecked()
        keys = sorted(data)
        # сценарії, де параметр осі X не змінювався (одна точка), не утворюють лінії – не показуємо їх
        single = [k for k in keys if len(data[k]) < 2]
        if xk != -1 and len(single) < len(keys):
            keys = [k for k in keys if k not in single]
        else:
            single = []
        multi_key = len(keys) > 1
        styles = ["-", "--", ":", "-."]
        markers = ["o", "s", "^", "D", "v", "P"]
        # ---- 1. Зібрати лінії ----
        lines = []
        for ki, key in enumerate(keys):
            klabel = " · ".join(f"{PARAM_SHORT[k].split(',')[0]} = {fmt(v)}" for k, v in zip(others, key))
            for mi, (m, mlab) in enumerate(zip(metrics, labels)):
                xs = sorted(data[key])
                ys = [float(np.mean([rp[0]["mean"][m] for rp in data[key][x]])) for x in xs]
                es = [float(np.mean([rp[0]["ci"][m] for rp in data[key][x]])) for x in xs]
                ps = [data[key][x][0][1] for x in xs]
                if log:                                    # нуль на логарифмічній шкалі не показується
                    keep = [i for i, y in enumerate(ys) if y > 0]
                    xs, ys, es, ps = ([v[i] for i in keep] for v in (xs, ys, es, ps))
                if not xs:
                    continue
                if len(metrics) > 1:
                    col, ls, mk = SERIES[mi % len(SERIES)], styles[ki % len(styles)], markers[ki % len(markers)]
                    label = mlab + (f" ({klabel})" if multi_key else "")
                else:
                    col, ls, mk = SERIES[ki % len(SERIES)], styles[(ki // len(SERIES)) % len(styles)], "o"
                    label = klabel if multi_key else mlab
                lines.append(dict(xs=xs, ys=ys, es=es, ps=ps, col=col, ls=ls, mk=mk, label=label, m=m))
        # ---- 2. Лінії, що збігаються, малюємо одну, з об'єднаним підписом (інакше одна «зникає») ----
        merged_lines, same_notes, near_notes = [], [], []
        for ln in lines:
            twin = next((o for o in merged_lines if o["xs"] == ln["xs"] and
                         np.allclose(o["ys"], ln["ys"], rtol=1e-3, atol=1e-12)), None)
            if twin is not None:
                twin["label"] += " = " + ln["label"]
                same_notes.append(twin["label"])
                continue
            near = next((o for o in merged_lines if o["xs"] == ln["xs"] and
                         np.allclose(o["ys"], ln["ys"], rtol=0.08, atol=1e-12)), None)
            if near is not None:                    # майже збігаються – друга лінія штрихом, щоб було видно обидві
                ln["ls"], ln["mk"] = (0, (4, 3)), "D"
                near_notes.append(tr("«{a}» ({x}) і «{b}» ({y})", a=near["label"], x=fmt(np.mean(near["ys"]), 3),
                                     b=ln["label"], y=fmt(np.mean(ln["ys"]), 3)))
            merged_lines.append(ln)
        # ---- 3. Малювання ----
        all_y, all_hi, n_lines = [], [], 0
        overload = unreachable = 0
        for ln in merged_lines:
            xs, ys, es = ln["xs"], ln["ys"], ln["es"]
            all_y += ys
            all_hi += [y + e for y, e in zip(ys, es)]
            kw = dict(color=ln["col"], ls=ln["ls"], marker=ln["mk"], ms=6, lw=2, label=ln["label"])
            if self.ci_box.isChecked():
                lo = [min(e, 0.9 * y) if log else min(e, y) if y >= 0 else e for y, e in zip(ys, es)]
                ax.errorbar(xs, ys, yerr=[lo, es], capsize=3, **kw)
            else:
                ax.plot(xs, ys, **kw)
            n_lines += 1
            # перевантаження (U ≥ 1): значення не стаціонарні – позначаємо червоним хрестиком
            bad = [(x, y) for x, y, p in zip(xs, ys, ln["ps"]) if utilization(p) >= 1]
            if bad:
                overload += len(bad)
                ax.scatter([b[0] for b in bad], [b[1] for b in bad], marker="x", s=90, linewidths=2.5,
                           color=PALETTE.get("err", "#c00"), zorder=6)
            if ln["m"] == 17:
                unreachable += sum(1 for p in ln["ps"] if p["D1"] < p["s1"] + p["s2"])
            if ln["m"] == 18:
                unreachable += sum(1 for p in ln["ps"] if p["D2"] < p["s1"] + p["s2b"])
        if not all_y:
            raise ValueError(tr("немає додатних значень для логарифмічної шкали"))
        y_max, y_min = max(all_y), min(all_y)
        flat = (y_max - y_min) <= 0.02 * max(abs(y_max), 1e-12)
        notes = [tr("Точки – середнє за R прогонів; вертикальні риски – 95 % довірчий інтервал.")]
        if overload:
            notes.append(tr("✕ – перевантаження (U ≥ 1): черга росте без меж, ці значення залежать від N і не є "
                            "стаціонарними – для висновків їх не використовують."))
        if unreachable:
            notes.append(tr("Точки з імовірністю 1 – дедлайн менший за саму обробку (s1 + s2), тобто недосяжний."))
        for t in same_notes:
            notes.append(tr("Лінії «{t}» збігаються, тому показано одну.", t=t))
        for t in near_notes:
            notes.append(tr("Лінії {t} майже збігаються – друга показана штрихом.", t=t))
        if log and y_min > 0 and y_max / y_min >= 5:
            ax.set_yscale("log")
            plain_log(ax.yaxis)
        else:
            ax.ticklabel_format(axis="y", style="plain", useOffset=False)
            if flat and y_min >= 0:
                ax.set_ylim(0, (max(all_hi if self.ci_box.isChecked() else all_y) or 1) * 1.25)
        if xk == 2:
            ax.set_xlabel(tr("Тривалість вторинної обробки МРЧ s2, мс"))
        elif xk == 4:
            ax.set_xlabel(tr("Дедлайн МРЧ D1, мс"))
        else:
            ax.set_xlabel(tr("№ сценарію") if xk == -1 else PARAM_SHORT[xk])
        ax.set_ylabel(ylabel)
        ax.set_title(title, fontsize=11)
        if n_lines > 1:
            if n_lines <= 3:
                self._legend(ax, loc="best")
            elif n_lines <= 9:                       # під графіком, щоб не закривати лінії
                self._legend(ax, loc="upper center", bbox_to_anchor=(0.5, -0.13), ncol=3)
            else:
                self._legend(ax, loc="center left", bbox_to_anchor=(1.01, 0.5))
        merged = [k for k in skip if k in self._varying_params() and k != xk]
        if merged:
            names = tr(" і ").join(PARAM_SHORT[k].split(',')[0] for k in sorted(merged))
            one = len(merged) == 1
            notes.append(tr("{p} на ці показники не впливає, тому сценарії, що відрізняються лише цим параметром, "
                            "дають однакові результати і показані однією точкою." if one else
                            "{p} на ці показники не впливають, тому сценарії, що відрізняються лише ними, "
                            "дають однакові результати і показані однією точкою.", p=names))
        if single:
            notes.append(tr("Не показано сценаріїв: {n} – у них {p} не змінювався (одна точка не утворює лінії).",
                            n=len(single), p=PARAM_SHORT[xk].split(',')[0]))
        if multi_key:
            notes.append(tr("Окремі лінії – для різних значень ") +
                         ", ".join(PARAM_SHORT[k].split(',')[0] for k in others) + ".")
        if flat and len(metrics) == 1 and xk != -1:
            notes.append(tr("Показник практично не залежить від {p} – лінія горизонтальна, це очікуваний результат.",
                            p=PARAM_SHORT[xk].split(',')[0]))
        if r["settings"]["replications"] < 5 and self.ci_box.isChecked():
            notes.append(tr("R = {r}: інтервали дуже широкі, для звіту беріть R ≥ 5.", r=r["settings"]["replications"]))
        self.note.setText(" ".join(notes))

    # ---------------- Сценарій ----------------
    def _draw_sc_types(self, ax):
        sc, res = self.result["scenario"], self.result["res"]
        names = [et.name for et in sc.event_types]
        y = np.arange(len(names))
        keys = [("sojourn_mean", tr("середній")), ("sojourn_p99", tr("99-й перцентиль")), ("sojourn_max", tr("максимальний"))]
        h = 0.26
        for i, (k, lab) in enumerate(keys):
            vals = [res["per_type"][n][k]["mean"] for n in names]
            ax.barh(y + (i - 1) * h, vals, height=h, color=SERIES[i], label=lab)
        dl = [et.deadline_ms for et in sc.event_types]
        ax.scatter(dl, y, marker="|", s=900, color=PALETTE.get("err", "#c00"), linewidths=2.5, zorder=5,
                   label=tr("дедлайн"))
        ax.set_yticks(y)
        ax.set_yticklabels([type_label(et.cls, et.name) for et in sc.event_types])
        ax.invert_yaxis()
        ax.set_xlabel(tr("Час реакції, мс"))
        ax.set_title(tr("Час реакції за типами подій (середнє за R прогонів)"), fontsize=11)
        allv = [res["per_type"][n][k]["mean"] for n in names for k, _ in keys] + dl
        pos = [v for v in allv if v > 0]
        if self.log_box.isChecked() and pos and max(pos) / min(pos) >= 10:
            ax.set_xscale("log")
            plain_log(ax.xaxis)
        else:
            ax.set_xlim(left=0)
        self._legend(ax, loc="upper center", bbox_to_anchor=(0.5, -0.13), ncol=4)
        self.note.setText(tr("Червона риска – дедлайн типу. Якщо стовпчик «максимальний» заходить за риску, "
                          "дедлайн порушувався."))

    def _draw_sc_time(self, ax):
        sc = self.result["scenario"]
        ev = self.result.get("events")
        if not ev or len(ev["arrival"]) == 0:
            raise ValueError(tr("немає даних про окремі події"))
        t, typ, soj = ev["arrival"], ev["type"], ev["sojourn"]
        t_s = t / 1000.0                                   # секунди – зручніше читати
        bins = np.linspace(t_s.min(), t_s.max(), 250)
        centers = (bins[:-1] + bins[1:]) / 2
        idx = np.clip(np.digitize(t_s, bins) - 1, 0, len(centers) - 1)
        for f in sc.active_faults():
            a = f.start_ms / 1000.0
            b = min(f.end_ms() / 1000.0, t_s.max())
            if b > a:
                ax.axvspan(a, b, color=PALETTE.get("err", "#c00"), alpha=0.10, lw=0)
                ax.text(a, 1.0, " " + tr_msg(f.label()), transform=ax.get_xaxis_transform(), va="top", ha="left",
                        fontsize=8, color=PALETTE.get("err", "#c00"))
        for k, et in enumerate(sc.event_types):
            m = typ == k
            if not m.any():
                continue
            mx = np.full(len(centers), -np.inf)
            np.maximum.at(mx, idx[m], soj[m])               # максимум у кожному інтервалі
            mx[np.isinf(mx)] = np.nan
            ls = "-" if et.cls == "HRT" else "--"
            ax.plot(centers, mx, ls, lw=1.6, color=SERIES[k % len(SERIES)],
                    label=type_label(et.cls, et.name))
        if np.nanmax(soj) / max(np.nanmin(soj[soj > 0]), 1e-12) >= 10:
            ax.set_yscale("log")
            plain_log(ax.yaxis)
        else:
            ax.set_ylim(bottom=0)
            ax.ticklabel_format(axis="y", style="plain", useOffset=False)
        ax.set_xlabel(tr("Модельний час, с"))
        ax.set_ylabel(tr("Макс. час реакції в інтервалі, мс"))
        ax.set_title(tr("Час реакції в часі (перший прогін, seed = %s)") % self.result["settings"]["seed"], fontsize=11)
        self._legend(ax, loc="upper center", bbox_to_anchor=(0.5, -0.13), ncol=3)
        note = (tr("Кожна точка – найбільший час реакції подій цього типу за короткий інтервал часу. "
                "Суцільні лінії – ЖРЧ, штрихові – МРЧ."))
        if sc.active_faults():
            note += tr(" Червона смуга – інтервал дії відмови.")
        if sc.utilization() >= 1:
            note += " " + tr("Увага: U = {u} ≥ 1 – процесор перевантажений, тому час реакції МРЧ постійно зростає. "
                            "Після останньої події МРЧ модель лише дообробляє чергу.", u=fmt(sc.utilization(), 3))
        self.note.setText(note)

    def _draw_sc_miss(self, ax):
        sc, res = self.result["scenario"], self.result["res"]
        names = [et.name for et in sc.event_types]
        x = np.arange(len(names))
        has_f = bool(sc.active_faults())
        w = 0.38 if has_f else 0.6
        a = [res["per_type"][n]["p_miss"]["mean"] for n in names]
        ax.bar(x - (w / 2 if has_f else 0), a, width=w, color=SERIES[0], label=tr("за весь час моделювання"))
        if has_f:
            b = [res["per_type"][n]["p_miss_fault"]["mean"] for n in names]
            ax.bar(x + w / 2, b, width=w, color=SERIES[4], label=tr("під час відмови"))
        ax.set_xticks(x)
        ax.set_xticklabels([type_label(et.cls, et.name, "\n") for et in sc.event_types],
                           fontsize=8)
        ax.set_ylabel(tr("Ймовірність порушення дедлайну"))
        ax.set_ylim(0, max(0.05, max(a + (b if has_f else [])) * 1.15))
        ax.set_title(tr("Ймовірність порушення дедлайнів"), fontsize=11)
        self._legend(ax)
        self.note.setText(tr("Для ЖРЧ будь-яке значення більше нуля означає порушення жорсткої вимоги."))

    def _make_toolbar(self):
        """
        Власна панель інструментів графіка (замість стандартної matplotlib): українські підписи,
        кольори теми програми. Самі дії виконує прихована стандартна панель matplotlib.
        """
        self._nav = NavigationToolbar2QT(self.canvas, self)
        self._nav.hide()
        bar = QWidget()
        row = QHBoxLayout(bar)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(6)
        self._b_pan = button(tr("✥ Перемістити"), "Tool", tr("Перетягуйте графік мишею (ще раз – вимкнути)"),
                             lambda: self._mode("pan"))
        self._b_zoom = button(tr("⊕ Збільшити"), "Tool", tr("Виділіть мишею прямокутник, щоб його збільшити "
                                                             "(ще раз – вимкнути)"), lambda: self._mode("zoom"))
        for b in (self._b_pan, self._b_zoom):
            b.setCheckable(True)
        for w in (button(tr("⌂ Скинути"), "Tool", tr("Повернути початковий вигляд графіка"), self._nav.home),
                  button("←", "Tool", tr("Попередній вигляд"), self._nav.back),
                  button("→", "Tool", tr("Наступний вигляд"), self._nav.forward),
                  self._b_pan, self._b_zoom):
            row.addWidget(w)
        row.addStretch(1)
        self._coords = muted("", hint=True)
        row.addWidget(self._coords)
        self.canvas.mpl_connect("motion_notify_event", self._on_move)
        return bar

    def _mode(self, which):
        (self._nav.pan if which == "pan" else self._nav.zoom)()
        mode = str(getattr(self._nav, "mode", "")).lower()
        self._b_pan.setChecked("pan" in mode)
        self._b_zoom.setChecked("zoom" in mode)

    def _on_move(self, ev):
        if ev.inaxes is None or ev.xdata is None:
            self._coords.setText("")
        else:
            self._coords.setText(f"x = {fmt(ev.xdata, 4)},  y = {fmt(ev.ydata, 4)}")

    def rebuild_toolbar(self):
        """Залишено для сумісності: власна панель не залежить від теми."""
        return

    # ------------------------------------------------------------------
    def save_png(self):
        if self.result is None:
            return
        os.makedirs(Config.PLOTS_FOLDER, exist_ok=True)
        start = os.path.join(os.path.abspath(Config.PLOTS_FOLDER),
                             tr("графік") + f"_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
        path, _ = QFileDialog.getSaveFileName(self, tr("Зберегти графік"), start, "PNG (*.png)")
        if not path:
            return
        try:
            self.fig.savefig(path, dpi=200, facecolor=self.fig.get_facecolor())
            self.banner.show_message(tr("Графік збережено: {p}", p=path), "ok")
        except Exception as e:
            self.banner.show_message(tr("Не вдалося зберегти графік: {e}", e=e), "err")
