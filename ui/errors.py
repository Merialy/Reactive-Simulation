"""
Перехоплення неочікуваних помилок: замість «тихого» закриття програми показується вікно,
а повний текст помилки записується у файл logs/errors.log (його зручно надіслати для аналізу).
"""
import os
import sys
import traceback
from datetime import datetime
from ui.i18n import tr

LOG = os.path.join("logs", "errors.log")


def log_error(text: str):
    try:
        os.makedirs(os.path.dirname(LOG), exist_ok=True)
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(f"\n===== {datetime.now():%Y-%m-%d %H:%M:%S} =====\n{text}\n")
    except OSError:
        pass


def install_excepthook():
    def hook(exc_type, exc, tb):
        text = "".join(traceback.format_exception(exc_type, exc, tb))
        log_error(text)
        sys.__stderr__.write(text)
        try:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.critical(None, tr("Неочікувана помилка"),
                                 f"{exc_type.__name__}: {exc}\n\n" + tr("Повний текст записано у файл {p}", p=os.path.abspath(LOG)))
        except Exception:
            pass
    sys.excepthook = hook
