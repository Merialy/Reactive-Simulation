"""
Головне вікно програми (версія 4.0, інтерфейс на Qt / PySide6).

Ліворуч – бічна панель з розділами і налаштуваннями запуску (R, seed, тема).
Праворуч – сторінки: Початок, Серія сценаріїв, Типи подій і відмови, Результати, Графіки, Історія.
Симуляція виконується у фоновому потоці (Worker), тому вікно не «зависає».
"""
import copy
import os
import time
from datetime import datetime

from PySide6.QtCore import QSettings, QSize, Qt, QTimer
from PySide6.QtWidgets import (QApplication, QFrame, QGridLayout, QHBoxLayout, QLabel, QListWidget, QListWidgetItem, QMainWindow,
                               QMessageBox, QProgressBar, QSpinBox, QStackedWidget, QVBoxLayout, QWidget)

from config import Config
from simulation.replications import run_batch, run_scenario_replications
from ui import theme
from ui import widgets as W
from ui.core import fmt, prepare_batch
from ui.errors import log_error
from ui.page_charts import ChartsPage
from ui.page_excel import ExcelPage
from ui.page_history import HistoryPage
from ui.page_results import ResultsPage
from ui.page_scenario import ScenarioPage
from ui.widgets import Card, PageHeader, Worker, button, muted, refresh_icons, set_icon, pad_spin
from ui.icons import icon
from ui.page_home import AboutDialog, HomePage
from ui.i18n import is_en, tr, tr_msg

NAV = [("home", tr("Початок"), "home"), ("excel", tr("Серія сценаріїв"), "table"),
       ("scenario", tr("Типи подій і відмови"), "layers"), ("results", tr("Результати"), "results"),
       ("charts", tr("Графіки"), "chart"), ("history", tr("Історія"), "history")]


class MainWindow(QMainWindow):
    def __init__(self, repo_factory=None):
        super().__init__()
        self.setWindowTitle(tr("Симулятор моделі реактивності {v}", v=Config.APP_VERSION))
        self.resize(1360, 860)
        self.setMinimumSize(960, 600)
        self.settings = QSettings("KhAI", "ReactiveSimulation")
        self._repo = None
        self._repo_factory = repo_factory
        self._worker = None
        self._started = 0.0
        # палітру задаємо ДО створення сторінок: від неї залежать, наприклад, кольори значків графіків
        self.dark = self.settings.value("dark", "false") == "true"
        self._set_palette()

        root = QWidget()
        h = QHBoxLayout(root)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(0)
        self.setCentralWidget(root)

        # ---------------- Бічна панель ----------------
        side = QFrame()
        side.setObjectName("Sidebar")
        side.setFixedWidth(232)
        sl = QVBoxLayout(side)
        sl.setContentsMargins(8, 22, 8, 16)
        sl.setSpacing(4)
        t = QLabel(tr("  Реактивність БО"))
        t.setObjectName("AppTitle")
        st = QLabel("   " + tr("версія {v}", v=Config.APP_VERSION))
        st.setObjectName("AppSubtitle")
        sl.addWidget(t)
        sl.addWidget(st)
        sl.addSpacing(18)
        self.nav = QListWidget()
        self.nav.setObjectName("Nav")
        self.nav.setIconSize(QSize(18, 18))
        for key, label, ic in NAV:
            it = QListWidgetItem("  " + label)
            it.setData(Qt.ItemDataRole.UserRole, key)
            self.nav.addItem(it)
        self.nav.currentRowChanged.connect(self._on_nav)
        sl.addWidget(self.nav, 1)

        cap = QLabel(tr("  НАЛАШТУВАННЯ ЗАПУСКУ"))
        cap.setObjectName("SideCaption")
        sl.addWidget(cap)
        form = QGridLayout()
        form.setContentsMargins(10, 6, 10, 6)
        form.setHorizontalSpacing(8)
        self.r_spin = QSpinBox()
        pad_spin(self.r_spin)
        self.r_spin.setRange(1, 200)
        self.r_spin.setValue(int(self.settings.value("replications", Config.DEFAULT_REPLICATIONS)))
        self.r_spin.setToolTip(tr("Кількість незалежних прогонів R. Більше – вужчий довірчий інтервал, довше рахує."))
        self.seed_spin = QSpinBox()
        pad_spin(self.seed_spin)
        self.seed_spin.setRange(0, 2_000_000_000)
        self.seed_spin.setValue(int(self.settings.value("seed", Config.DEFAULT_SEED)))
        self.seed_spin.setToolTip(tr("Початкове значення генератора. Той самий seed – ті самі результати."))
        for i, (lab, w) in enumerate(((tr("Прогонів R"), self.r_spin), ("Seed", self.seed_spin))):
            ql = QLabel(lab)
            ql.setObjectName("SideLabel")
            form.addWidget(ql, i, 0)
            form.addWidget(w, i, 1)
        sl.addLayout(form)
        sl.addSpacing(6)
        self.theme_btn = button("", "SideButton", slot=self.toggle_theme)
        sl.addWidget(self.theme_btn)
        self.lang_btn = button("English" if not is_en() else "Українська", "SideButton",
                               slot=self.toggle_language, icon="globe")
        self.lang_btn.setToolTip(tr("Змінити мову інтерфейсу (потрібен перезапуск програми)"))
        sl.addWidget(self.lang_btn)
        self.about_btn = button(tr("Про програму"), "SideButton", slot=self.show_about, icon="info")
        sl.addWidget(self.about_btn)
        h.addWidget(side)

        # ---------------- Сторінки ----------------
        self.stack = QStackedWidget()
        h.addWidget(self.stack, 1)
        self.home = HomePage(self.go, self.open_history, self.show_about)
        self.excel = ExcelPage()
        self.scenario = ScenarioPage()
        self.results = ResultsPage(self.repo)
        self.charts = ChartsPage()
        self.history = HistoryPage(self.repo)
        self.pages = {"home": self.home, "excel": self.excel, "scenario": self.scenario,
                      "results": self.results, "charts": self.charts, "history": self.history}
        for key, *_ in NAV:
            self.stack.addWidget(self.pages[key])

        self.excel.run_requested.connect(self.run_excel)
        self.scenario.run_requested.connect(self.run_scenario)
        self.results.open_charts.connect(lambda: self.go("charts"))
        self.history.open_result.connect(self.show_doc)

        # ---------------- Рядок стану ----------------
        sb = self.statusBar()
        self.status = QLabel(tr("Готово"))
        self.progress = QProgressBar()
        self.progress.setFixedWidth(260)
        self.progress.hide()
        self.storage_label = QLabel(tr("Сховище: перевірка…"))
        sb.addWidget(self.status, 1)
        sb.addPermanentWidget(self.progress)
        sb.addPermanentWidget(self.storage_label)

        self.apply_theme()
        self.nav.setCurrentRow(0)
        QTimer.singleShot(50, self._init_storage)
        QTimer.singleShot(0, self.charts.rebuild_toolbar)    # значки графіка – у кольорах уже застосованої теми

    # ==================================================================
    # Тема, навігація
    # ==================================================================
    def _set_palette(self):
        pal = theme.DARK if self.dark else theme.LIGHT
        W.PALETTE.clear()
        W.PALETTE.update(pal)
        app = QApplication.instance()
        if app is not None:
            try:                                     # Qt ≥ 6.8: не брати кольори з теми Windows
                app.styleHints().setColorScheme(Qt.ColorScheme.Dark if self.dark else Qt.ColorScheme.Light)
            except AttributeError:
                pass
            app.setPalette(theme.palette(pal))
            app.setStyleSheet(theme.qss(pal))       # на рівні програми – діє і на діалогові вікна

    def apply_theme(self):
        self._set_palette()
        self.theme_btn.setText(tr("Світла тема") if self.dark else tr("Темна тема"))
        set_icon(self.theme_btn, "sun" if self.dark else "moon")
        refresh_icons(self)
        for i, (_, _, ic) in enumerate(NAV):
            self.nav.item(i).setIcon(icon(ic, W.PALETTE["side_text"], 18))
        from ui.icons import app_icon
        self.setWindowIcon(app_icon())
        self.home.redraw()
        # перемалювати сторінки, де кольори задаються в коді
        self.excel.refresh()
        self.scenario.update_summary()
        self.results.render()
        self.charts.draw()

    def toggle_theme(self):
        self.dark = not self.dark
        self.settings.setValue("dark", "true" if self.dark else "false")
        self.apply_theme()
        self.charts.rebuild_toolbar()

    def toggle_language(self):
        """Зміна мови: зберегти вибір і перезапустити програму (підписи створюються під час запуску)."""
        new = "uk" if is_en() else "en"
        self.settings.setValue("lang", new)
        self.settings.sync()
        # повідомлення двома мовами: його читає людина, яка ще бачить стару мову, а перемикає на нову
        ask = ("Мову буде змінено на українську. Перезапустити програму зараз?\n"
               "The language will be changed to Ukrainian. Restart the program now?" if new == "uk" else
               "The language will be changed to English. Restart the program now?\n"
               "Мову буде змінено на англійську. Перезапустити програму зараз?")
        if QMessageBox.question(self, "🌐", ask) == QMessageBox.StandardButton.Yes:
            if self._worker is not None and self._worker.isRunning():
                self._worker.terminate()
            import sys
            from PySide6.QtCore import QProcess
            QProcess.startDetached(sys.executable, [os.path.abspath(sys.argv[0])] + sys.argv[1:], os.getcwd())
            QApplication.instance().quit()

    def show_about(self):
        AboutDialog(self).exec()

    def open_history(self, query: str = ""):
        self.go("history")
        self.history.search.setText(query)

    def refresh_recent(self):
        try:
            self.home.set_recent(self.repo().list_results(limit=300))
        except Exception:
            self.home.set_recent([])

    def _on_nav(self, row):
        key = NAV[row][0]
        self.stack.setCurrentWidget(self.pages[key])
        if key == "history":
            self.history.reload()
        elif key == "home" and self._repo is not None:
            self.refresh_recent()

    def go(self, key, open_file=False):
        self.nav.setCurrentRow([k for k, *_ in NAV].index(key))
        if open_file and key == "excel":
            self.excel.open_file()

    # ==================================================================
    # Сховище
    # ==================================================================
    def repo(self):
        if self._repo is None:
            self._repo = self._make_repo()
        return self._repo

    def _make_repo(self):
        if self._repo_factory:
            return self._repo_factory()
        from storage import get_repository
        return get_repository()

    def _init_storage(self):
        def job(_):
            from storage.repository import last_storage_warning
            r = self._make_repo()
            return r, last_storage_warning()

        def done(res):
            self._repo, warn = res
            mongo = self._repo.kind == "MongoDB"
            dot = W.PALETTE["ok"] if mongo else W.PALETTE["warn"]
            self.storage_label.setText(f"<span style='color:{dot}'>●</span> " + ("MongoDB" if mongo else tr("JSON-файли")))
            self.storage_label.setToolTip(tr_msg(self._repo.describe()) + ("\n" + warn if warn else ""))
            self.home.storage.setText(f"<span style='color:{dot}'>●</span> " +
                                      tr("Сховище результатів: {d}", d=tr_msg(self._repo.describe())) +
                                      ("" if mongo else tr(" – MongoDB не запущена, результати зберігаються у файлах.")))
            self.refresh_recent()

        self._storage_worker = Worker(job, self)
        self._storage_worker.done.connect(done)
        self._storage_worker.failed.connect(lambda e: self.storage_label.setText(tr("Сховище: помилка")))
        self._storage_worker.start()

    # ==================================================================
    # Запуск симуляції
    # ==================================================================
    def _run_settings(self):
        r, s = int(self.r_spin.value()), int(self.seed_spin.value())
        self.settings.setValue("replications", r)
        self.settings.setValue("seed", s)
        return {"replications": r, "seed": s}

    def _busy(self, on: bool, text: str = ""):
        self.excel.run_btn.setEnabled(not on)
        self.scenario.run_btn.setEnabled(not on)
        self.progress.setVisible(on)
        self.progress.setValue(0)
        if text:
            self.status.setText(text)

    def _start(self, fn, on_done, text):
        if self._worker is not None and self._worker.isRunning():
            return
        self._busy(True, text)
        self._started = time.time()
        self._worker = Worker(fn, self)
        self._worker.progress.connect(self._on_progress)
        self._worker.done.connect(on_done)
        self._worker.failed.connect(self._on_failed)
        self._worker.start()

    def _on_progress(self, done, total):
        self.progress.setMaximum(max(1, total))
        self.progress.setValue(done)
        el = time.time() - self._started
        eta = el / done * (total - done) if done else 0
        self.status.setText(tr("Симуляція: прогін {a} з {b}  ·  {t} с", a=done, b=total, t=f"{el:.0f}") +
                            (tr(", залишилось ≈ {t} с", t=f"{eta:.0f}") if done and done < total else ""))

    def _on_failed(self, text):
        self._busy(False, tr("Помилка під час симуляції"))
        QMessageBox.critical(self, tr("Помилка симуляції"), text.split("\n\n")[0])

    # ---------------- Серія з Excel ----------------
    def run_excel(self):
        table = copy.deepcopy(self.excel.data)
        run_list, warns, errs = prepare_batch(table)
        if not any(p is not None for p in run_list):
            self.excel.banner.show_message(tr("Немає жодного сценарію, який можна змоделювати. ")
                                           + (tr("Помилки: ") + "; ".join(tr("сценарій {n}: {e}", n=c + 1, e=tr_msg(e)) for c, e in errs.items())
                                              if errs else tr("Заповніть таблицю або відкрийте Excel-файл.")), "err")
            return
        st = self._run_settings()
        self._pending = dict(kind="excel", table=table, run_list=run_list, warnings=warns, errors=errs,
                             settings=st, source=table.source_name())
        n = sum(p is not None for p in run_list)
        self._start(lambda cb: run_batch(run_list, st["replications"], st["seed"], Config.WORKERS, cb),
                    self._excel_done, tr("Симуляція серії: {n} сценаріїв × {r} прогонів…", n=n, r=st["replications"]))

    def _excel_done(self, batch):
        r = self._pending
        r.update(batch=batch, elapsed=time.time() - self._started, when=datetime.now().strftime("%d.%m.%Y %H:%M"))
        self._finish(r, tr("Готово: {n} сценаріїв за {t} с", n=sum(b is not None for b in batch), t=fmt(r["elapsed"], 2)))

    # ---------------- Сценарій ----------------
    def run_scenario(self):
        sc, errors = self.scenario.current()
        if errors or sc is None:
            return
        st = self._run_settings()
        self._pending = dict(kind="scenario", scenario=sc, settings=st)

        def job(cb):
            from simulation.engine import EventSimulation
            res = run_scenario_replications(sc, st["replications"], st["seed"], Config.WORKERS, cb)
            events = EventSimulation(seed=st["seed"]).run(sc, keep_events=True).get("events")
            return res, events

        self._start(job, self._scenario_done, tr("Симуляція сценарію «{s}»: {r} прогонів…", s=sc.name, r=st["replications"]))

    def _scenario_done(self, out):
        res, events = out
        r = self._pending
        r.update(res=res, events=events, elapsed=time.time() - self._started,
                 when=datetime.now().strftime("%d.%m.%Y %H:%M"))
        self._finish(r, tr("Готово: сценарій «{s}» за {t} с", s=r["scenario"].name, t=fmt(r["elapsed"], 2)))

    def _finish(self, r, text):
        self._busy(False, text)
        self.go("results")
        for fn in (self.results.set_result, self.charts.set_result):
            try:
                fn(r)
            except Exception:
                import traceback
                log_error(traceback.format_exc())
                self.status.setText(tr("Результат пораховано, але його не вдалося повністю показати – "
                                    "подробиці у logs/errors.log"))

    # ---------------- Документ з історії ----------------
    def show_doc(self, doc):
        self.results.set_result({"kind": "doc", "doc": doc})
        self.go("results")

    def closeEvent(self, e):
        if self._worker is not None and self._worker.isRunning():
            if QMessageBox.question(self, tr("Симуляція ще триває"),
                                    tr("Симуляція ще не завершилась. Закрити програму?")) \
                    != QMessageBox.StandardButton.Yes:
                e.ignore()
                return
            self._worker.terminate()
        e.accept()
