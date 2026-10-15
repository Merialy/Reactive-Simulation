"""
Симулятор моделі реактивності – новий інтерфейс (версія 4.0, Qt / PySide6).

    python app.py

Старий інтерфейс (Tkinter) поки доступний: python main.py
"""
import multiprocessing
import os
import sys


def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))     # щоб працювали відносні шляхи (scenarios, data)
    sys.path.insert(0, os.getcwd())
    from PySide6.QtCore import QLibraryInfo, QSettings, QTranslator
    from PySide6.QtWidgets import QApplication
    from ui import i18n
    # мову треба задати ДО імпорту сторінок: частина підписів – константи модулів
    lang = str(QSettings("KhAI", "ReactiveSimulation").value("lang", "uk"))
    i18n.set_lang(lang)
    from ui.main_window import MainWindow
    from ui.errors import install_excepthook

    app = QApplication(sys.argv)
    app.setApplicationName("Симулятор моделі реактивності")
    app.setStyle("Fusion")
    # стандартні кнопки Qt («Так/Ні», діалоги файлів) – тією ж мовою
    tr_qt = QTranslator(app)
    if tr_qt.load(f"qtbase_{'en' if i18n.is_en() else 'uk'}",
                  QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath)):
        app.installTranslator(tr_qt)
    install_excepthook()
    w = MainWindow()
    # вікно не більше за екран: на ноутбуці – на весь екран, на великому моніторі – по центру
    geo = app.primaryScreen().availableGeometry()
    if geo.width() < 1500 or geo.height() < 940:
        w.showMaximized()
    else:
        w.resize(1400, 900)
        w.move(geo.center() - w.rect().center())
        w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    multiprocessing.freeze_support()      # потрібно для паралельних прогонів у Windows
    main()
