"""
Значки інтерфейсу: прості лінійні піктограми 24×24 у форматі SVG (намальовані для цієї програми).

    icon("play", "#FFFFFF")       → QIcon потрібного кольору
    svg("chart", "#2563EB")       → текст SVG (наприклад, для збереження у файл)

Колір підставляється під час створення, тому значки перефарбовуються разом з темою
(див. MainWindow.apply_theme → widgets.refresh_icons).
"""

# {c} – колір лінії/заливки. Використано лише прості елементи (path з M/L/H/V/C/Z, rect, circle, line, polyline).
ICONS = {
    "home": "<path d='M3 11 L12 4 L21 11'/><path d='M5 10 V20 H19 V10'/><rect x='10' y='14' width='4' height='6'/>",
    "table": "<rect x='3' y='4' width='18' height='16' rx='2'/><line x1='3' y1='10' x2='21' y2='10'/>"
             "<line x1='3' y1='15' x2='21' y2='15'/><line x1='9' y1='4' x2='9' y2='20'/>"
             "<line x1='15' y1='4' x2='15' y2='20'/>",
    "layers": "<path d='M12 3 L21 8 L12 13 L3 8 Z'/><polyline points='3,12.5 12,17.5 21,12.5'/>"
              "<polyline points='3,16.5 12,21.5 21,16.5'/>",
    "results": "<rect x='4' y='3' width='16' height='18' rx='2'/><line x1='8' y1='8' x2='16' y2='8'/>"
               "<line x1='8' y1='12' x2='16' y2='12'/><line x1='8' y1='16' x2='13' y2='16'/>",
    "chart": "<polyline points='3,3 3,21 21,21'/><polyline points='7,16 11,11 14,14 20,7'/>",
    "history": "<circle cx='12' cy='12' r='8.5'/><polyline points='12,7.5 12,12 15,14'/>",
    "folder": "<path d='M3 7 V18 C3 19 4 20 5 20 H19 C20 20 21 19 21 18 V9 C21 8 20 7 19 7 H12 L10 5 H5 C4 5 3 6 3 7 Z'/>",
    "save": "<path d='M5 3 H16 L21 8 V19 C21 20 20 21 19 21 H5 C4 21 3 20 3 19 V5 C3 4 4 3 5 3 Z'/>"
            "<rect x='7' y='3' width='8' height='5'/><rect x='7' y='13' width='10' height='8'/>",
    "play": "<path d='M7 4 L20 12 L7 20 Z' fill='{c}'/>",
    "plus": "<line x1='12' y1='5' x2='12' y2='19'/><line x1='5' y1='12' x2='19' y2='12'/>",
    "trash": "<polyline points='4,7 20,7'/><path d='M6 7 L7 20 H17 L18 7'/><path d='M9 7 V4 H15 V7'/>"
             "<line x1='10' y1='11' x2='10' y2='16.5'/><line x1='14' y1='11' x2='14' y2='16.5'/>",
    "copy": "<rect x='8' y='8' width='12' height='12' rx='2'/><path d='M4 16 V6 C4 5 5 4 6 4 H16'/>",
    "eraser": "<path d='M8 20 L3.5 15.5 L13.5 5.5 L20 12 L12 20 Z'/><line x1='8' y1='20' x2='21' y2='20'/>"
              "<line x1='8.5' y1='10.5' x2='15' y2='17'/>",
    "database": "<path d='M4 6 C4 3 20 3 20 6 C20 9 4 9 4 6 Z'/><path d='M4 6 V18 C4 21 20 21 20 18 V6'/>"
                "<path d='M4 12 C4 15 20 15 20 12'/>",
    "export": "<line x1='12' y1='3' x2='12' y2='15'/><polyline points='7,10 12,15 17,10'/>"
              "<polyline points='4,17 4,20 20,20 20,17'/>",
    "image": "<rect x='3' y='4' width='18' height='16' rx='2'/><circle cx='9' cy='10' r='2'/>"
             "<polyline points='21,16 16,11 6,20'/>",
    "info": "<circle cx='12' cy='12' r='9'/><line x1='12' y1='11' x2='12' y2='16.5'/>"
            "<circle cx='12' cy='7.8' r='0.9' fill='{c}'/>",
    "globe": "<circle cx='12' cy='12' r='9'/><line x1='3' y1='12' x2='21' y2='12'/>"
             "<path d='M12 3 C8 7 8 17 12 21'/><path d='M12 3 C16 7 16 17 12 21'/>",
    "sun": "<circle cx='12' cy='12' r='4'/><line x1='12' y1='2' x2='12' y2='4.5'/><line x1='12' y1='19.5' x2='12' y2='22'/>"
           "<line x1='2' y1='12' x2='4.5' y2='12'/><line x1='19.5' y1='12' x2='22' y2='12'/>"
           "<line x1='4.9' y1='4.9' x2='6.7' y2='6.7'/><line x1='17.3' y1='17.3' x2='19.1' y2='19.1'/>"
           "<line x1='4.9' y1='19.1' x2='6.7' y2='17.3'/><line x1='17.3' y1='6.7' x2='19.1' y2='4.9'/>",
    "moon": "<path d='M20 14.5 C18 19 11.5 21 7.5 17 C3.5 13 5 6 10 4 C8.5 9 13 15 20 14.5 Z'/>",
    "cpu": "<rect x='6' y='6' width='12' height='12' rx='2'/><rect x='9.5' y='9.5' width='5' height='5'/>"
           "<line x1='9' y1='2.5' x2='9' y2='6'/><line x1='15' y1='2.5' x2='15' y2='6'/>"
           "<line x1='9' y1='18' x2='9' y2='21.5'/><line x1='15' y1='18' x2='15' y2='21.5'/>"
           "<line x1='2.5' y1='9' x2='6' y2='9'/><line x1='2.5' y1='15' x2='6' y2='15'/>"
           "<line x1='18' y1='9' x2='21.5' y2='9'/><line x1='18' y1='15' x2='21.5' y2='15'/>",
    "bolt": "<path d='M13 2 L4 14 H12 L11 22 L20 10 H12 Z'/>",
    "user": "<circle cx='12' cy='8' r='4'/><path d='M4 21 C4 16 8 14 12 14 C16 14 20 16 20 21'/>",
    "school": "<path d='M2 9 L12 4 L22 9 L12 14 Z'/><path d='M6 11 V16 C6 18.5 18 18.5 18 16 V11'/>"
              "<line x1='22' y1='9' x2='22' y2='15'/>",
    "mail": "<rect x='3' y='5' width='18' height='14' rx='2'/><polyline points='3,7 12,13 21,7'/>",
    "search": "<circle cx='11' cy='11' r='7'/><line x1='16' y1='16' x2='21' y2='21'/>",
    "compare": "<rect x='3' y='4' width='7' height='16' rx='1.5'/><rect x='14' y='4' width='7' height='16' rx='1.5'/>"
               "<line x1='5.5' y1='9' x2='7.5' y2='9'/><line x1='16.5' y1='13' x2='18.5' y2='13'/>",
    "refresh": "<path d='M4.5 11 C5 7 8 4 12 4 C15 4 17.5 5.5 19 8'/><polyline points='19.5,3.5 19.5,8.5 14.5,8.5'/>"
               "<path d='M19.5 13 C19 17 16 20 12 20 C9 20 6.5 18.5 5 16'/><polyline points='4.5,20.5 4.5,15.5 9.5,15.5'/>",
    "arrow": "<line x1='5' y1='12' x2='19' y2='12'/><polyline points='13,6 19,12 13,18'/>",
    "new": "<path d='M6 3 H14 L19 8 V21 H6 Z'/><polyline points='14,3 14,8 19,8'/>"
           "<line x1='12.5' y1='11.5' x2='12.5' y2='17.5'/><line x1='9.5' y1='14.5' x2='15.5' y2='14.5'/>",
}


def svg(name: str, color: str, size: int = 24, stroke: float = 1.8) -> str:
    body = ICONS[name].replace("{c}", color)
    return (f"<svg xmlns='http://www.w3.org/2000/svg' width='{size}' height='{size}' viewBox='0 0 24 24' "
            f"fill='none' stroke='{color}' stroke-width='{stroke}' stroke-linecap='round' "
            f"stroke-linejoin='round'>{body}</svg>")


_cache = {}


def icon(name: str, color: str, size: int = 20):
    """QIcon з SVG (з урахуванням HiDPI). Повертає порожній значок, якщо Qt SVG недоступний."""
    key = (name, color, size)
    if key in _cache:
        return _cache[key]
    from PySide6.QtCore import QByteArray, QRectF, Qt
    from PySide6.QtGui import QIcon, QPainter, QPixmap
    try:
        from PySide6.QtSvg import QSvgRenderer
    except ImportError:
        return QIcon()
    ic = QIcon()
    for scale in (1, 2):                              # звичайний екран і HiDPI
        px = QPixmap(size * scale, size * scale)
        px.fill(Qt.GlobalColor.transparent)
        r = QSvgRenderer(QByteArray(svg(name, color).encode("utf-8")))
        p = QPainter(px)
        r.render(p, QRectF(0, 0, size * scale, size * scale))
        p.end()
        px.setDevicePixelRatio(scale)
        ic.addPixmap(px)
    _cache[key] = ic
    return ic


APP_COLOR = "#2563EB"


def app_svg() -> str:
    """Значок програми: синій квадрат із заокругленими кутами і білим процесором."""
    cpu = ICONS["cpu"]
    return ("<svg xmlns='http://www.w3.org/2000/svg' width='256' height='256' viewBox='0 0 24 24'>"
            f"<rect x='0.5' y='0.5' width='23' height='23' rx='5.5' fill='{APP_COLOR}'/>"
            "<g transform='translate(3.6 3.6) scale(0.7)' fill='none' stroke='#FFFFFF' stroke-width='2.1' "
            f"stroke-linecap='round' stroke-linejoin='round'>{cpu}</g></svg>")


def app_icon():
    """QIcon програми у кількох розмірах (вікно, панель завдань Windows, Alt+Tab)."""
    from PySide6.QtCore import QByteArray, QRectF, Qt
    from PySide6.QtGui import QIcon, QPainter, QPixmap
    try:
        from PySide6.QtSvg import QSvgRenderer
    except ImportError:
        return QIcon()
    ic = QIcon()
    data = QByteArray(app_svg().encode("utf-8"))
    for size in (16, 20, 24, 32, 40, 48, 64, 128, 256):
        px = QPixmap(size, size)
        px.fill(Qt.GlobalColor.transparent)
        p = QPainter(px)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        QSvgRenderer(data).render(p, QRectF(0, 0, size, size))
        p.end()
        ic.addPixmap(px)
    return ic
