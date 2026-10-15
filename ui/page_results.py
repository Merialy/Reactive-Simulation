"""
Сторінка «Результати»: підсумки останнього запуску (серії з Excel або сценарію з типами подій)
або результату, відкритого з історії бази даних.
"""
import os
from datetime import datetime

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QFileDialog, QHBoxLayout, QHeaderView, QLabel,
                               QTableWidget, QVBoxLayout, QWidget, QScrollArea, QFrame)

from config import Config
from ui.core import (FAULT_ROWS, comma, METRIC_SHORT, PER_TYPE_ROWS, export_batch_excel, export_scenario_excel, fmt,
                     fmt_ci)
from ui.widgets import Banner, Card, PageHeader, button, color, muted, ro_item
from ui.i18n import tr, tr_msg
from ui.core import METRIC_NAMES, cls_name, type_label


class Tile(QFrame):
    """Плитка з великим числом і підписом."""

    def __init__(self, caption: str, parent=None):
        super().__init__(parent)
        self.setObjectName("Card")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 12, 16, 12)
        lay.setSpacing(2)
        self.value = QLabel("–")
        self.value.setObjectName("BigNumber")
        self.caption = muted(caption, hint=True)
        lay.addWidget(self.value)
        lay.addWidget(self.caption)

    def set(self, value: str, level: str = "", caption: str = None):
        self.value.setText(value)
        c = {"ok": color("ok"), "warn": color("warn"), "err": color("err")}.get(level)
        self.value.setStyleSheet(f"color: {c.name()};" if c else "")
        if caption is not None:
            self.caption.setText(caption)


class ResultsPage(QWidget):
    open_charts = Signal()

    def __init__(self, get_repo, parent=None):
        super().__init__(parent)
        self.setObjectName("Page")
        self.get_repo = get_repo          # функція, що повертає сховище (MongoDB або JSON)
        self.result = None

        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 20)
        outer.setSpacing(14)
        self.header = PageHeader(tr("Результати"), tr("Ще немає результатів. Запустіть серію з Excel або сценарій."))
        self.save_btn = self.header.add_action(button(tr("Зберегти в базу даних"), slot=self.save_to_db, icon="database",
                                                      tooltip=tr("Записати сценарії і результати в MongoDB. "
                                                              "Повторне збереження не створює дублікатів.")))
        self.export_btn = self.header.add_action(button(tr("Експорт в Excel…"), slot=self.export_excel, icon="export"))
        self.charts_btn = self.header.add_action(button(tr("Графіки"), "Primary", slot=lambda: self.open_charts.emit(), icon="chart"))
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
        self.body = QVBoxLayout(body)
        self.body.setContentsMargins(0, 0, 8, 0)
        self.body.setSpacing(14)

        self.tiles_row = QHBoxLayout()
        self.tiles_row.setSpacing(14)
        self.tiles = [Tile("") for _ in range(4)]
        for t in self.tiles:
            self.tiles_row.addWidget(t)
        self.body.addLayout(self.tiles_row)

        self.main_card = Card(tr("Показники"))
        self.show_ci = QCheckBox(tr("Показувати ± 95 % довірчий інтервал"))
        self.show_ci.setChecked(True)
        self.show_ci.toggled.connect(self.render)
        self.main_card.head_actions.addWidget(self.show_ci)
        self.table = self._make_table()
        self.main_card.add(self.table)
        self.legend = muted("", hint=True)
        self.main_card.add(self.legend)
        self.body.addWidget(self.main_card)

        self.fault_card = Card(tr("Під час дії відмов"),
                               tr("Показники лише для подій, що надійшли, поки діяла відмова "
                               "(без відкидання перших і останніх 5 % – відмова є перехідним процесом)."))
        self.fault_table = self._make_table()
        self.fault_card.add(self.fault_table)
        self.body.addWidget(self.fault_card)

        self.warn_card = Card(tr("Попередження і помилки"))
        self.warn_text = QLabel()
        self.warn_text.setWordWrap(True)
        self.warn_text.setTextFormat(Qt.TextFormat.RichText)
        self.warn_text.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.warn_card.add(self.warn_text)
        self.body.addWidget(self.warn_card)
        self.body.addStretch(1)
        self.render()

    @staticmethod
    def _make_table() -> QTableWidget:
        t = QTableWidget()
        t.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        t.setAlternatingRowColors(True)
        t.horizontalHeader().setDefaultSectionSize(130)
        t.horizontalHeader().setMinimumSectionSize(90)
        t.verticalHeader().setDefaultSectionSize(30)
        t.verticalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        return t

    @staticmethod
    def _fit_height(t: QTableWidget, max_rows: int = 30):
        rows = min(t.rowCount(), max_rows)
        h = t.horizontalHeader().height() + sum(t.rowHeight(r) for r in range(rows)) + 6
        if t.horizontalScrollBar().isVisible() or t.columnCount() > 6:
            h += 12
        t.setMinimumHeight(h)
        t.setMaximumHeight(h)

    # ------------------------------------------------------------------
    def set_result(self, result: dict):
        self.result = result
        self.banner.hide()
        self.render()

    def render(self):
        r = self.result
        has = r is not None
        self.save_btn.setEnabled(has and r["kind"] in ("excel", "scenario"))
        self.export_btn.setEnabled(has and r["kind"] in ("excel", "scenario"))
        self.charts_btn.setEnabled(has and r["kind"] in ("excel", "scenario"))
        for w in (self.main_card, self.fault_card, self.warn_card):
            w.setVisible(has)
        for t in self.tiles:
            t.setVisible(has)
        if not has:
            self.header.subtitle.setText(tr("Ще немає результатів. Запустіть серію з Excel або сценарій."))
            return
        {"excel": self._render_excel, "scenario": self._render_scenario, "doc": self._render_doc}[r["kind"]]()
        self._fit_height(self.table, 40)
        self._fit_height(self.fault_table, 20)

    def _cell(self, mean, ci):
        return fmt_ci(mean, ci) if self.show_ci.isChecked() else fmt(mean)

    # ---------------- Серія з Excel ----------------
    def _render_excel(self):
        r = self.result
        st = r["settings"]
        batch, warns, errs = r["batch"], r["warnings"], r["errors"]
        n_done = sum(1 for b in batch if b is not None)
        self.header.title.setText(tr("Результати серії"))
        self.header.subtitle.setText(f"{r['source']}  ·  R = {st['replications']}, seed = {st['seed']}  ·  "
                                     f"{r['when']}  ·  " + tr("розрахунок {t} с", t=fmt(r["elapsed"], 2)))
        self.tiles[0].set(str(n_done), "", tr("сценаріїв змодельовано"))
        self.tiles[1].set(str(len(warns)), "warn" if warns else "ok", tr("з попередженнями"))
        self.tiles[2].set(str(len(errs)), "err" if errs else "ok", tr("з помилками (не моделювались)"))
        worst = max((b["mean"][18] for b in batch if b), default=0.0)
        self.tiles[3].set(fmt(worst, 3), "err" if worst > 0 else "ok", tr("макс. P порушення дедлайну ЖРЧ"))

        t = self.table
        t.clear()
        cols = [c for c in range(len(batch)) if batch[c] is not None or c in errs]
        t.setRowCount(18)
        t.setColumnCount(len(cols))
        t.setVerticalHeaderLabels([f"{i}. {n}" for i, n in enumerate(METRIC_SHORT, 1)])
        for i in range(18):
            t.verticalHeaderItem(i).setToolTip(METRIC_NAMES[i])
        for j, c in enumerate(cols):
            head = tr("Сценарій {n}", n=c + 1)
            tip = ""
            if c in errs:
                head = "✖ " + head
                tip = tr("Помилка: ") + tr_msg(errs[c])
            elif c in warns:
                head = "⚠ " + head
                tip = "\n".join(tr_msg(x) for x in warns[c])
            hi = ro_item(head, align_right=False)
            hi.setToolTip(tip or tr("Без зауважень"))
            t.setHorizontalHeaderItem(j, hi)
            res = batch[c]
            for m in range(1, 19):
                if res is None:
                    it = ro_item("—")
                    it.setForeground(color("muted"))
                else:
                    it = ro_item(self._cell(res["mean"][m], res["ci"][m]))
                    v = res["mean"][m]
                    if m == 18 and v > 0:
                        it.setBackground(color("err_bg")); it.setForeground(color("err"))
                    elif m == 17 and v > 0:
                        it.setBackground(color("warn_bg")); it.setForeground(color("warn"))
                t.setItem(m - 1, j, it)
        t.resizeColumnsToContents()
        self.legend.setText(tr("Часи – у мілісекундах. Жовтим виділено порушення м'яких дедлайнів (МРЧ), червоним – "
                            "жорстких (ЖРЧ). ⚠ – сценарій з попередженням, ✖ – з помилкою (наведіть на заголовок)."))
        self.main_card.setVisible(True)
        self.fault_card.setVisible(False)
        lines = [f"<b style='color:{color('err').name()}'>" + tr("Сценарій {n}", n=c + 1) + f":</b> {tr_msg(e)}" for c, e in sorted(errs.items())]
        lines += [f"<b style='color:{color('warn').name()}'>" + tr("Сценарій {n}", n=c + 1) + f":</b> {tr_msg(w)}"
                  for c, ws in sorted(warns.items()) for w in ws]
        self.warn_text.setText(comma("<br>".join(lines)))
        self.warn_card.setVisible(bool(lines))

    # ---------------- Сценарій з типами подій ----------------
    def _render_scenario(self):
        r = self.result
        sc, res, st = r["scenario"], r["res"], r["settings"]
        self.header.title.setText(tr("Результати сценарію"))
        self.header.subtitle.setText(f"{sc.name}  ·  R = {st['replications']}, seed = {st['seed']}  ·  "
                                     f"{r['when']}  ·  " + tr("розрахунок {t} с", t=fmt(r["elapsed"], 2)))
        u = sc.utilization()
        self.tiles[0].set(fmt(u, 3), "err" if u >= 1 else ("warn" if u > 0.95 else "ok"), tr("завантаження U"))
        hrt = [et.name for et in sc.event_types if et.cls == "HRT"]
        srt = [et.name for et in sc.event_types if et.cls != "HRT"]
        pm_h = max((res["per_type"][n]["p_miss"]["mean"] for n in hrt), default=0.0)
        pm_s = max((res["per_type"][n]["p_miss"]["mean"] for n in srt), default=0.0)
        if hrt:
            self.tiles[1].set(fmt(pm_h, 3), "err" if pm_h > 0 else "ok", tr("макс. P порушення дедлайну ЖРЧ"))
        else:
            self.tiles[1].set("—", "", tr("подій ЖРЧ у сценарії немає"))
        if srt:
            self.tiles[2].set(fmt(pm_s, 3), "warn" if pm_s > 0 else "ok", tr("макс. P порушення дедлайну МРЧ"))
        else:
            self.tiles[2].set("—", "", tr("подій МРЧ у сценарії немає"))
        if sc.active_faults():
            uf = sc.utilization_during_faults()
            fs = res.get("faults", {})
            extra = []
            if fs.get("storm_events", {}).get("mean"):
                extra.append(tr("хибних переривань ≈ {n}", n=f"{round(fs['storm_events']['mean']):,}".replace(",", " ")))
            if fs.get("overrun_events", {}).get("mean"):
                extra.append(tr("подій з перевищенням WCET ≈ {n}", n=f"{round(fs['overrun_events']['mean']):,}".replace(",", " ")))
            self.tiles[3].set(fmt(uf, 3), "err" if uf >= 1 else ("warn" if uf > 0.95 else "ok"),
                              tr("U під час відмов") + (" · " + "; ".join(extra) if extra else ""))
        else:
            self.tiles[3].set(tr("немає"), "", tr("відмов у сценарії"))

        self._fill_types(self.table, sc, res, PER_TYPE_ROWS)
        self.legend.setText(tr("Часи – у мілісекундах; час реакції = від надходження переривання до завершення обробки. "
                            "Червоним виділено порушення дедлайнів ЖРЧ, жовтим – МРЧ."))
        if sc.active_faults():
            self.fault_card.setVisible(True)
            self.fault_card.findChild(QLabel, "CardTitle").setText(
                tr("Під час дії відмов: ") + "; ".join(tr_msg(f.label()) for f in sc.active_faults()))
            self._fill_types(self.fault_table, sc, res, FAULT_ROWS)
        else:
            self.fault_card.setVisible(False)
        _, warns = sc.check()
        self.warn_text.setText(comma("<br>".join(f"⚠ {tr_msg(w)}" for w in warns)))
        self.warn_card.setVisible(bool(warns))

    def _fill_types(self, t: QTableWidget, sc, res, rows):
        t.clear()
        names = [et.name for et in sc.event_types]
        t.setRowCount(len(rows))
        t.setColumnCount(len(names))
        t.setVerticalHeaderLabels([lab for _, lab in rows])
        for j, et in enumerate(sc.event_types):
            hi = ro_item(type_label(et.cls, et.name), align_right=False)
            hi.setForeground(color("hrt") if et.cls == "HRT" else color("srt"))
            hi.setToolTip(tr("Дедлайн {d} мс, WCET + T_core = {w} мс, інтенсивність {r} 1/мс",
                             d=fmt(et.deadline_ms), w=fmt(et.service_ms(sc.platform)), r=fmt(et.rate)))
            t.setHorizontalHeaderItem(j, hi)
            for i, (key, _) in enumerate(rows):
                v = res["per_type"][et.name].get(key)
                if v is None:
                    t.setItem(i, j, ro_item("—"))
                    continue
                txt = fmt(v["mean"], 6) if key.startswith("count") else self._cell(v["mean"], v["ci"])
                it = ro_item(txt)
                if key.startswith("p_miss") and v["mean"] > 0:
                    lvl = "err" if et.cls == "HRT" else "warn"
                    it.setBackground(color(lvl + "_bg")); it.setForeground(color(lvl))
                if key.startswith("sojourn_max") and v["mean"] > et.deadline_ms:
                    it.setForeground(color("err" if et.cls == "HRT" else "warn"))
                t.setItem(i, j, it)
        t.resizeColumnsToContents()

    # ---------------- Документ з бази даних ----------------
    def _render_doc(self):
        d = self.result["doc"]
        st = d.get("settings") or {}
        self.header.title.setText(tr("Результат з бази даних"))
        self.header.subtitle.setText(f"{d.get('scenario_name', '')}  ·  R = {st.get('replications')}, "
                                     f"seed = {st.get('seed')}  ·  " +
                                     tr("збережено {d}  ·  теги: {t}", d=d.get("created_at", ""),
                                        t=", ".join(d.get("tags") or []) or "—"))
        u = d.get("utilization", 0.0)
        self.tiles[0].set(fmt(u, 3), "err" if u >= 1 else ("warn" if u > 0.95 else "ok"), tr("завантаження U"))
        cl = {c["metric"]: c for c in d.get("classic", [])}
        p17, p18 = cl.get(17, {}).get("mean", 0.0), cl.get(18, {}).get("mean", 0.0)
        self.tiles[1].set(fmt(p18, 3), "err" if p18 > 0 else "ok", tr("P порушення дедлайну ЖРЧ"))
        self.tiles[2].set(fmt(p17, 3), "warn" if p17 > 0 else "ok", tr("P порушення дедлайну МРЧ"))
        self.tiles[3].set(d.get("app_version", "–"), "", tr("версія програми"))
        t = self.table
        t.clear()
        per_type = d.get("per_type") or []
        if per_type:
            keys = [k for k, _ in PER_TYPE_ROWS]
            t.setRowCount(len(keys))
            t.setColumnCount(len(per_type))
            t.setVerticalHeaderLabels([lab for _, lab in PER_TYPE_ROWS])
            for j, pt in enumerate(per_type):
                hi = ro_item(type_label(pt.get('class'), pt['name']), align_right=False)
                hi.setForeground(color("hrt") if pt.get("class") == "HRT" else color("srt"))
                t.setHorizontalHeaderItem(j, hi)
                for i, k in enumerate(keys):
                    v = pt["metrics"].get(k)
                    t.setItem(i, j, ro_item("—" if not v else (fmt(v["mean"], 6) if k == "count"
                                                                 else self._cell(v["mean"], v["ci"]))))
        else:
            t.setRowCount(18)
            t.setColumnCount(1)
            t.setVerticalHeaderLabels([f"{i}. {n}" for i, n in enumerate(METRIC_SHORT, 1)])
            t.setHorizontalHeaderItem(0, ro_item(tr("Значення"), align_right=False))
            for m in range(1, 19):
                c = cl.get(m)
                t.setItem(m - 1, 0, ro_item(self._cell(c["mean"], c["ci"]) if c else "—"))
        t.resizeColumnsToContents()
        self.legend.setText(tr("Відкрито зі сховища. Щоб перерахувати, відкрийте сценарій і запустіть знову."))
        self.fault_card.setVisible(False)
        self.warn_card.setVisible(False)

    # ------------------------------------------------------------------
    def save_to_db(self):
        r = self.result
        if not r or r["kind"] not in ("excel", "scenario"):
            return
        try:
            from simulation.model import Scenario
            repo = self.get_repo()
            settings = dict(r["settings"])
            new = same = 0
            if r["kind"] == "excel":
                source = r["source"]
                settings["source_file"] = source
                for c, (p, res) in enumerate(zip(r["run_list"], r["batch"]), 1):
                    if p is None or res is None:
                        continue
                    sc = Scenario.from_legacy(p, name=f"{source} – " + tr("сценарій {n}", n=c))
                    sid = repo.save_scenario(sc)
                    _, is_new = repo.save_result_checked(sc, {"mean": res["mean"], "ci": res["ci"]}, settings,
                                                         scenario_id=sid, tags=[source])
                    new += is_new
                    same += not is_new
            else:
                sc = r["scenario"]
                sid = repo.save_scenario(sc)
                _, is_new = repo.save_result_checked(sc, r["res"], settings, scenario_id=sid, tags=[sc.name])
                new, same = int(is_new), int(not is_new)
            msg = tr("Збережено в {d}: нових результатів – {n}", d=tr_msg(repo.describe()), n=new)
            if same:
                msg += tr("; вже були в базі (ті самі параметри, R і seed), не дублювались – {n}", n=same)
            self.banner.show_message(msg + ".", "ok")
        except Exception as e:
            self.banner.show_message(tr("Не вдалося зберегти в базу даних: {e}", e=e), "err")

    def export_excel(self):
        r = self.result
        if not r or r["kind"] not in ("excel", "scenario"):
            return
        os.makedirs(Config.EXPORT_FOLDER, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        base = os.path.splitext(r["source"])[0] if r["kind"] == "excel" else tr("сценарій")
        start = os.path.join(os.path.abspath(Config.EXPORT_FOLDER), tr("Результати") + f"_{base}_{stamp}.xlsx")
        path, _ = QFileDialog.getSaveFileName(self, tr("Експорт результатів"), start, "Excel (*.xlsx)")
        if not path:
            return
        try:
            if r["kind"] == "excel":
                export_batch_excel(path, r["table"], r["batch"], r["settings"], r["warnings"], r["errors"])
            else:
                export_scenario_excel(path, r["scenario"], r["res"], r["settings"])
            self.banner.show_message(tr("Звіт збережено: {p}", p=path), "ok")
        except Exception as e:
            self.banner.show_message(tr("Не вдалося експортувати: {e}", e=e), "err")
