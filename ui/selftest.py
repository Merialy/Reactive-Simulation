"""
Самоперевірка нового інтерфейсу.

    python -m ui.selftest

Програма сама відкриває вікно, проходить усі сторінки (Excel → серія, сценарій з відмовою,
результати, графіки, історія, порівняння, темна тема), робить знімки екрана і записує
журнал у папку docs/screens. Ваша база MongoDB не використовується: результати
зберігаються в тимчасову папку.
"""
import glob
import multiprocessing
import os
import shutil
import sys
import tempfile
import time
import traceback

OUT = os.path.join("docs", "screens_en" if "--en" in sys.argv else "screens")
LOG = []


def log(msg):
    line = f"[{time.strftime('%H:%M:%S')}] {msg}"
    print(line, flush=True)
    LOG.append(line)


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    os.chdir(root)
    sys.path.insert(0, root)
    os.makedirs(OUT, exist_ok=True)
    for f in glob.glob(os.path.join(OUT, "*.png")):
        os.remove(f)

    from ui import i18n
    i18n.set_lang("en" if "--en" in sys.argv else "uk")
    from PySide6.QtWidgets import QApplication, QDialog
    from PySide6.QtCore import QTimer
    import PySide6
    from config import Config
    from storage.repository import JsonRepository
    from ui.main_window import MainWindow

    log(f"Python {sys.version.split()[0]}, PySide6 {PySide6.__version__}, програма {Config.APP_VERSION}")
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    tmp = tempfile.mkdtemp(prefix="rs_selftest_")
    w = MainWindow(repo_factory=lambda: JsonRepository(tmp))
    w.resize(1440, 900)
    w.show()
    errors = 0

    def pump(sec=0.3):
        end = time.time() + sec
        while time.time() < end:
            app.processEvents()
            time.sleep(0.02)

    def shot(name, widget=None):
        pump(0.4)
        (widget or w).grab().save(os.path.join(OUT, name + ".png"))
        log(f"  знімок {name}.png")

    def wait_worker(limit=600):
        t0 = time.time()
        while w._worker is not None and w._worker.isRunning():
            pump(0.2)
            if time.time() - t0 > limit:
                raise TimeoutError("симуляція триває надто довго")
        pump(0.5)

    def step(title, fn):
        nonlocal errors
        log(title)
        try:
            fn()
        except Exception:
            errors += 1
            log("  ПОМИЛКА:\n" + traceback.format_exc())

    slot_errors = []

    def hook(t, e, tb):                       # помилки всередині обробників Qt теж рахуємо
        slot_errors.append("".join(traceback.format_exception(t, e, tb)))
        log("  ПОМИЛКА В ОБРОБНИКУ:\n" + slot_errors[-1])
    sys.excepthook = hook

    def page_is(key):
        assert w.stack.currentWidget() is w.pages[key], \
            f"очікувалась сторінка «{key}», відкрита інша"

    w.r_spin.setValue(2)
    w.seed_spin.setValue(42)

    step("1. Стартова сторінка", lambda: shot("01_home"))

    def excel():
        files = sorted(glob.glob(os.path.join("Excel file", "*TestFile_4*.xlsx"))) or \
            sorted(glob.glob(os.path.join("Excel file", "*.xlsx"))) or sorted(glob.glob("*.xlsx"))
        w.go("excel")
        if files:
            log(f"  файл: {files[0]}")
            assert w.excel.open_file(files[0]), "файл не відкрився"
        else:
            log("  Excel-файлів не знайдено – заповнюю таблицю вручну")
            for c, s2 in enumerate((0.3, 0.5, 0.7)):
                for r, v in enumerate((1, 0.01, s2, 5000, 1, 2, 0.1, 0.2)):
                    w.excel.data.set(r, c, v)
            w.excel.refresh()
        shot("02_excel")
    step("2. Серія сценаріїв (таблиця параметрів)", excel)

    def run_excel():
        t0 = time.time()
        w.run_excel()
        wait_worker()
        log(f"  серія порахована за {time.time() - t0:.1f} с")
        assert w.results.result and w.results.result["kind"] == "excel", "немає результату серії"
        page_is("results")
        shot("03_results_excel")
        assert w.results.extra_card.isVisible() and w.results.theory_card.isVisible(), \
            "немає карток «Додаткові показники» / «Порівняння з теорією»"
        shot("03b_results_extra", w.results.extra_card)
        shot("03c_results_theory", w.results.theory_card)
        w.go("charts")
        shot("04_charts_excel")
        w.charts.cb1.setCurrentIndex(max(0, w.charts.cb1.findData(5)))   # порушення дедлайнів від D1
        shot("05_charts_excel_pmiss")
        w.results.save_to_db()
        log("  " + w.results.banner.label.text())
        w.results.save_to_db()
        log("  повторно: " + w.results.banner.label.text())
    step("3. Запуск серії, результати, графік, збереження", run_excel)

    def scenario():
        w.go("scenario")
        p = os.path.join("scenarios", "boryviter_fault_storm.json")
        assert w.scenario.load(p), "сценарій не відкрився"
        shot("06_scenario")
        sc, errs = w.scenario.current()
        assert not errs, errs
        log(f"  сценарій «{sc.name}», типів: {len(sc.event_types)}, відмов: {len(sc.active_faults())}")
    step("4. Сторінка сценарію", scenario)

    def run_sc():
        t0 = time.time()
        w.run_scenario()
        wait_worker()
        log(f"  сценарій порахований за {time.time() - t0:.1f} с")
        assert w.results.result and w.results.result["kind"] == "scenario", "немає результату сценарію"
        page_is("results")
        shot("07_results_scenario")
        assert w.results.system_card.isVisible(), "немає картки «Системні показники»"
        shot("07b_results_types", w.results.main_card)
        shot("07c_results_system", w.results.system_card)
        w.go("charts")
        for i, name in enumerate(("08_chart_types", "09_chart_time", "10_chart_miss")):
            w.charts.cb1.setCurrentIndex(i)
            shot(name)
        w.results.save_to_db()
        log("  " + w.results.banner.label.text())
    step("5. Запуск сценарію з відмовою, графіки", run_sc)

    def background():
        w.go("scenario")
        p = os.path.join("scenarios", "boryviter_background.json")
        assert w.scenario.load(p), "сценарій з фоновими задачами не відкрився"
        shot("07d_scenario_background")
        w.run_scenario()
        wait_worker()
        page_is("results")
        rows = [w.results.system_table.verticalHeaderItem(i).text() for i in range(w.results.system_table.rowCount())]
        assert any("Фонов" in r or "Background" in r for r in rows), "немає показників фонових задач"
        shot("07e_results_background", w.results.system_card)
        shot("07f_results_background_types", w.results.main_card)
        log("  " + w.results.concl.text()[:200].replace("<br>", " | "))
    step("5б. Фонові задачі (запис логів на SD-карту)", background)

    def history():
        w.go("history")
        shot("11_history")
        n = w.history.table.rowCount()
        log(f"  у сховищі результатів: {n}")
        assert n >= 2, "очікувалось щонайменше 2 результати"
        w.history.table.selectAll()
        ids = w.history._selected_ids()
        from ui.page_history import CompareDialog
        docs = [w.history.get_repo().get_result(i) for i in ids[:3]]
        dlg = CompareDialog(docs, w)
        dlg.show()
        shot("12_compare", dlg)
        dlg.close()
        w.history.table.selectRow(0)
        w.history.open_selected()
        shot("13_results_from_db")
    step("6. Історія, порівняння, відкриття з бази", history)

    def new_sc():
        w.go("scenario")
        w.scenario.new_scenario()
        w.scenario.add_type()
        w.scenario.add_fault("overrun")
        w.scenario.update_summary()
        shot("14_scenario_new")
    step("7. Новий сценарій (додати тип і відмову)", new_sc)

    def dark():
        w.toggle_theme()
        w.go("home")
        shot("15_dark_home")
        w.go("results")
        shot("16_dark_results")
        w.go("charts")
        shot("17_dark_charts")
        w.toggle_theme()
    step("8. Темна тема", dark)

    def about_clear():
        from ui.page_home import AboutDialog
        w.go("home")
        w.refresh_recent()
        shot("18_home_after_runs")
        dlg = AboutDialog(w)
        dlg.show()
        shot("19_about", dlg)
        dlg.close()
        from PySide6.QtWidgets import QMessageBox
        ask = QMessageBox.question
        QMessageBox.question = staticmethod(lambda *a, **k: QMessageBox.StandardButton.Yes)
        try:
            w.go("excel")
            w.excel.clear_table()
            assert w.excel.data.n_cols == 3 and w.excel._filled_count() == 0, "таблиця не очистилась"
            shot("20_series_cleared")
        finally:
            QMessageBox.question = ask
    step("9. Стартовий екран з результатами, «Про програму», очищення таблиці", about_clear)

    errors += len(slot_errors)
    log(f"ГОТОВО. Помилок: {errors}. Знімки і журнал: {os.path.abspath(OUT)}")
    with open(os.path.join(OUT, "selftest.log"), "w", encoding="utf-8") as f:
        f.write("\n".join(LOG) + "\n")
    shutil.rmtree(tmp, ignore_errors=True)
    QTimer.singleShot(200, app.quit)
    app.exec()
    return 1 if errors else 0


if __name__ == "__main__":
    multiprocessing.freeze_support()
    sys.exit(main())
