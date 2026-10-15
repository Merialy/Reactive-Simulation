"""
Оформлення нового інтерфейсу: кольори світлої і темної теми та таблиця стилів Qt (QSS).
Усі кольори – тут, щоб змінювати вигляд програми в одному місці.
"""

LIGHT = dict(
    bg="#F3F5F9", surface="#FFFFFF", surface2="#F8FAFC", border="#DCE2EA", text="#1E293B", muted="#64748B",
    accent="#2563EB", accent_hover="#1D4ED8", accent_text="#FFFFFF", accent_soft="#DBEAFE",
    side_bg="#0F172A", side_text="#CBD5E1", side_muted="#64748B", side_sel="#1E3A8A", side_sel_text="#FFFFFF",
    ok="#15803D", ok_bg="#DCFCE7", warn="#B45309", warn_bg="#FEF3C7", err="#B91C1C", err_bg="#FEE2E2",
    hrt="#C2410C", srt="#1D4ED8", grid="#E2E8F0", sel="#BFDBFE", sel_text="#0F172A",
)

DARK = dict(
    bg="#0B1220", surface="#131C2E", surface2="#182338", border="#26334A", text="#E2E8F0", muted="#94A3B8",
    accent="#3B82F6", accent_hover="#60A5FA", accent_text="#FFFFFF", accent_soft="#1E3A5F",
    side_bg="#060B16", side_text="#CBD5E1", side_muted="#64748B", side_sel="#1D4ED8", side_sel_text="#FFFFFF",
    ok="#4ADE80", ok_bg="#14351F", warn="#FBBF24", warn_bg="#3A2A0A", err="#F87171", err_bg="#3F1515",
    hrt="#FB923C", srt="#60A5FA", grid="#26334A", sel="#1E3A8A", sel_text="#FFFFFF",
)

# Кольори графіків (однакові для обох тем; підібрані для розрізнення)
SERIES = ["#2563EB", "#EA580C", "#16A34A", "#9333EA", "#DC2626", "#0891B2", "#CA8A04", "#DB2777", "#4B5563"]


def palette(c: dict):
    """
    Палітра Qt для всієї програми. Задається явно, щоб вигляд не залежав від теми Windows
    (інакше в темному режимі Windows частина елементів стає чорною).
    """
    from PySide6.QtGui import QColor, QPalette
    p = QPalette()
    R = QPalette.ColorRole
    for role, key in ((R.Window, "bg"), (R.WindowText, "text"), (R.Base, "surface"), (R.AlternateBase, "surface2"),
                      (R.Text, "text"), (R.Button, "surface"), (R.ButtonText, "text"), (R.ToolTipBase, "surface"),
                      (R.ToolTipText, "text"), (R.Highlight, "accent"), (R.HighlightedText, "accent_text"),
                      (R.PlaceholderText, "muted"), (R.Link, "accent"), (R.BrightText, "accent_text"),
                      (R.Light, "surface"), (R.Midlight, "surface2"), (R.Mid, "border"), (R.Dark, "border"),
                      (R.Shadow, "border")):
        p.setColor(role, QColor(c[key]))
    for role in (R.Text, R.WindowText, R.ButtonText):
        p.setColor(QPalette.ColorGroup.Disabled, role, QColor(c["muted"]))
    return p


_ICON_SVG = {
    "down": "<path d='M2.5 4 L6 7.5 L9.5 4' fill='none' stroke='{c}' stroke-width='1.7' stroke-linecap='round' stroke-linejoin='round'/>",
    "up": "<path d='M2.5 8 L6 4.5 L9.5 8' fill='none' stroke='{c}' stroke-width='1.7' stroke-linecap='round' stroke-linejoin='round'/>",
    "check": "<path d='M2.5 6.2 L5 8.6 L9.5 3.6' fill='none' stroke='{c}' stroke-width='1.9' stroke-linecap='round' stroke-linejoin='round'/>",
}


def icon(name: str, col: str) -> str:
    """Маленький SVG-значок (стрілка, галочка) потрібного кольору; повертає шлях для таблиці стилів."""
    import os
    import tempfile
    folder = os.path.join(tempfile.gettempdir(), "reactive_sim_icons")
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, f"{name}_{col.lstrip('#')}.svg")
    if not os.path.exists(path):
        with open(path, "w", encoding="utf-8") as f:
            f.write("<svg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'>"
                    + _ICON_SVG[name].format(c=col) + "</svg>")
    return path.replace("\\", "/")


def qss(c: dict) -> str:
    down, up = icon("down", c["muted"]), icon("up", c["muted"])
    down_s, up_s = icon("down", c["side_text"]), icon("up", c["side_text"])
    check = icon("check", c["accent_text"])
    return f"""
* {{ font-family: "Segoe UI", "Inter", "Helvetica Neue", Arial; font-size: 10pt; }}
QMainWindow, QWidget#Page, QScrollArea, QScrollArea > QWidget > QWidget {{ background: {c['bg']}; color: {c['text']}; }}
QWidget {{ color: {c['text']}; }}
QToolTip {{ background: {c['surface']}; color: {c['text']}; border: 1px solid {c['border']}; padding: 6px; }}

/* ---- Бічна панель ---- */
QFrame#Sidebar {{ background: {c['side_bg']}; border: none; }}
QLabel#AppTitle {{ color: #FFFFFF; font-size: 13pt; font-weight: 600; }}
QLabel#AppSubtitle {{ color: {c['side_muted']}; font-size: 9pt; }}
QLabel#SideCaption {{ color: {c['side_muted']}; font-size: 8pt; font-weight: 600; letter-spacing: 1px; }}
QLabel#SideLabel {{ color: {c['side_text']}; }}
QListWidget#Nav {{ background: transparent; border: none; outline: none; color: {c['side_text']}; }}
QListWidget#Nav::item {{ padding: 10px 14px; border-radius: 8px; margin: 2px 8px; }}
QListWidget#Nav::item:hover {{ background: rgba(255,255,255,0.06); }}
QListWidget#Nav::item:selected {{ background: {c['side_sel']}; color: {c['side_sel_text']}; }}
QFrame#Sidebar QSpinBox {{ background: rgba(255,255,255,0.06); color: #FFFFFF; border: 1px solid rgba(255,255,255,0.12);
    border-radius: 6px; padding: 4px 6px; }}
QPushButton#SideButton {{ background: rgba(255,255,255,0.06); color: {c['side_text']}; border: 1px solid rgba(255,255,255,0.10);
    border-radius: 6px; padding: 6px 10px; }}
QPushButton#SideButton:hover {{ background: rgba(255,255,255,0.12); }}

/* ---- Заголовки сторінок ---- */
QLabel#PageTitle {{ font-size: 17pt; font-weight: 600; color: {c['text']}; }}
QLabel#PageSubtitle {{ color: {c['muted']}; font-size: 10pt; }}
QLabel#CardTitle {{ font-size: 11pt; font-weight: 600; color: {c['text']}; }}
QLabel#Muted {{ color: {c['muted']}; }}
QLabel#Hint {{ color: {c['muted']}; font-size: 9pt; }}
QLabel#BigNumber {{ font-size: 20pt; font-weight: 600; }}
QFrame#Fact {{ background: {c['surface2']}; border: 1px solid {c['border']}; border-radius: 10px; }}
QLabel#FactValue {{ font-size: 15pt; font-weight: 600; color: {c['text']}; }}
QLabel#FactValue[level="ok"] {{ color: {c['ok']}; }}
QLabel#FactValue[level="warn"] {{ color: {c['warn']}; }}
QLabel#FactValue[level="err"] {{ color: {c['err']}; }}

/* ---- Картки ---- */
QFrame#Card {{ background: {c['surface']}; border: 1px solid {c['border']}; border-radius: 12px; }}
QFrame#Banner {{ border-radius: 10px; border: 1px solid {c['border']}; background: {c['surface2']}; }}
QFrame#Banner[level="warn"] {{ background: {c['warn_bg']}; border-color: {c['warn']}; }}
QFrame#Banner[level="err"] {{ background: {c['err_bg']}; border-color: {c['err']}; }}
QFrame#Banner[level="ok"] {{ background: {c['ok_bg']}; border-color: {c['ok']}; }}
QLabel#Badge {{ border-radius: 10px; padding: 2px 10px; font-weight: 600; }}
QLabel#Badge[level="ok"] {{ background: {c['ok_bg']}; color: {c['ok']}; }}
QLabel#Badge[level="warn"] {{ background: {c['warn_bg']}; color: {c['warn']}; }}
QLabel#Badge[level="err"] {{ background: {c['err_bg']}; color: {c['err']}; }}
QLabel#Badge[level="info"] {{ background: {c['accent_soft']}; color: {c['accent']}; }}

/* ---- Кнопки ---- */
QPushButton {{ background: {c['surface']}; color: {c['text']}; border: 1px solid {c['border']}; border-radius: 8px;
    padding: 7px 14px; }}
QPushButton:hover {{ border-color: {c['accent']}; color: {c['accent']}; }}
QPushButton:disabled {{ color: {c['muted']}; background: {c['surface2']}; border-color: {c['border']}; }}
QPushButton#Primary {{ background: {c['accent']}; color: {c['accent_text']}; border: none; font-weight: 600; padding: 8px 18px; }}
QPushButton#Primary:hover {{ background: {c['accent_hover']}; color: {c['accent_text']}; }}
QPushButton#Primary:disabled {{ background: {c['border']}; color: {c['muted']}; }}
QPushButton#Danger:hover {{ border-color: {c['err']}; color: {c['err']}; }}
QPushButton#Link {{ border: none; background: transparent; color: {c['accent']}; padding: 2px 4px; text-align: left; }}
QPushButton#Link:hover {{ text-decoration: underline; }}
QPushButton#Tool {{ padding: 4px 10px; border-radius: 6px; }}
QPushButton#Tool:checked {{ background: {c['accent_soft']}; border-color: {c['accent']}; color: {c['accent']}; }}

/* ---- Поля ---- */
QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox, QPlainTextEdit, QTextEdit {{ background: {c['surface']}; color: {c['text']};
    border: 1px solid {c['border']}; border-radius: 6px; padding: 5px 8px; selection-background-color: {c['sel']};
    selection-color: {c['sel_text']}; }}
QLineEdit:focus, QSpinBox:focus, QComboBox:focus, QPlainTextEdit:focus {{ border-color: {c['accent']}; }}
QLineEdit[invalid="true"] {{ border-color: {c['err']}; background: {c['err_bg']}; }}
QComboBox QAbstractItemView {{ background: {c['surface']}; color: {c['text']}; selection-background-color: {c['sel']};
    selection-color: {c['sel_text']}; border: 1px solid {c['border']}; }}
QCheckBox {{ spacing: 6px; }}
QCheckBox::indicator {{ width: 16px; height: 16px; border: 1px solid {c['border']}; border-radius: 4px;
    background: {c['surface']}; }}
QCheckBox::indicator:hover {{ border-color: {c['accent']}; }}
QCheckBox::indicator:checked {{ background: {c['accent']}; border-color: {c['accent']}; }}
QCheckBox::indicator:checked {{ image: url({check}); }}
QAbstractSpinBox {{ padding-top: 3px; padding-bottom: 3px; padding-right: 20px; }}
QAbstractSpinBox::up-button {{ subcontrol-origin: border; subcontrol-position: top right; width: 18px;
    border: none; background: transparent; }}
QAbstractSpinBox::down-button {{ subcontrol-origin: border; subcontrol-position: bottom right; width: 18px;
    border: none; background: transparent; }}
QAbstractSpinBox::up-arrow {{ image: url({up}); width: 10px; height: 10px; }}
QAbstractSpinBox::down-arrow {{ image: url({down}); width: 10px; height: 10px; }}
QFrame#Sidebar QAbstractSpinBox::up-arrow {{ image: url({up_s}); }}
QFrame#Sidebar QAbstractSpinBox::down-arrow {{ image: url({down_s}); }}
QComboBox {{ padding-right: 24px; }}
QComboBox::drop-down {{ subcontrol-origin: padding; subcontrol-position: center right; width: 22px; border: none; }}
QComboBox::down-arrow {{ image: url({down}); width: 10px; height: 10px; }}

/* ---- Таблиці ---- */
QTableWidget, QTableView {{ background: {c['surface']}; alternate-background-color: {c['surface2']}; color: {c['text']};
    gridline-color: {c['grid']}; border: 1px solid {c['border']}; border-radius: 8px;
    selection-background-color: {c['sel']}; selection-color: {c['sel_text']}; }}
QHeaderView {{ background: {c['surface2']}; border: none; }}
QHeaderView::section {{ background: {c['surface2']}; color: {c['muted']}; border: none; border-bottom: 1px solid {c['border']};
    border-right: 1px solid {c['grid']}; padding: 6px 8px; font-weight: 600; }}
QTableCornerButton::section {{ background: {c['surface2']}; border: none; }}

/* ---- Інше ---- */
QStatusBar {{ background: {c['surface']}; border-top: 1px solid {c['border']}; color: {c['muted']}; }}
QStatusBar QLabel {{ color: {c['muted']}; padding: 0 6px; }}
QProgressBar {{ background: {c['surface2']}; border: 1px solid {c['border']}; border-radius: 6px; height: 14px;
    text-align: center; color: {c['text']}; font-size: 8pt; }}
QProgressBar::chunk {{ background: {c['accent']}; border-radius: 5px; }}
QScrollBar:vertical {{ background: transparent; width: 10px; margin: 0; }}
QScrollBar::handle:vertical {{ background: {c['border']}; border-radius: 5px; min-height: 30px; }}
QScrollBar:horizontal {{ background: transparent; height: 10px; margin: 0; }}
QScrollBar::handle:horizontal {{ background: {c['border']}; border-radius: 5px; min-width: 30px; }}
QScrollBar::add-line, QScrollBar::sub-line {{ width: 0; height: 0; }}
QTabWidget::pane {{ border: none; }}
QTabBar::tab {{ background: transparent; color: {c['muted']}; padding: 8px 14px; border-bottom: 2px solid transparent; }}
QTabBar::tab:selected {{ color: {c['accent']}; border-bottom: 2px solid {c['accent']}; font-weight: 600; }}
QMessageBox, QDialog {{ background: {c['bg']}; }}
QToolBar {{ background: transparent; border: none; spacing: 4px; }}
QToolButton {{ background: transparent; border: 1px solid transparent; border-radius: 6px; padding: 3px; }}
QToolButton:hover {{ background: {c['accent_soft']}; border-color: {c['border']}; }}
QToolButton:checked {{ background: {c['accent_soft']}; border-color: {c['accent']}; }}
"""
