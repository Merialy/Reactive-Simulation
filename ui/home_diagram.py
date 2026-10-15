"""
Схема моделі для стартового екрана: переривання → первинний обробник (Q0) → черги ЖРЧ (QH) і МРЧ (QS) → процесор.
Малюється засобами matplotlib, кольори – з поточної теми.
"""
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch, Rectangle

from ui.i18n import tr


def draw_model(fig, P: dict):
    fig.clear()
    fig.set_facecolor(P["surface"])
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 44)
    ax.axis("off")
    text, muted, border = P["text"], P["muted"], P["border"]
    hrt, srt, acc = P["hrt"], P["srt"], P["accent"]

    def box(x, y, w, h, title, sub, color):
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.4,rounding_size=1.6",
                                    fc=P["surface2"], ec=color, lw=1.6))
        ax.text(x + w / 2, y + h * 0.70, title, ha="center", va="center", fontsize=10, color=text, weight="bold")
        ax.text(x + w / 2, y + h * 0.30, sub, ha="center", va="center", fontsize=7, color=muted, linespacing=1.2)

    def arrow(x1, y1, x2, y2, color, lw=1.6, rad=0.0):
        ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>", mutation_scale=11, color=color, lw=lw,
                                     connectionstyle=f"arc3,rad={rad}"))

    def queue(x, y, n, color):
        for i in range(n):
            ax.add_patch(Rectangle((x + i * 3.1, y), 2.5, 3.6, fc=color, ec="none", alpha=0.85 - i * 0.12))

    # джерела подій
    ax.text(1, 37.5, tr("Переривання"), fontsize=9, color=text, weight="bold")
    for i, (lab, col) in enumerate(((tr("ЖРЧ · періодично"), hrt), (tr("МРЧ · випадково"), srt))):
        yy = 30 - i * 8
        ax.text(1, yy + 1.6, lab, fontsize=7.5, color=col)
        for k in range(4):
            ax.add_patch(Rectangle((1 + k * (2.6 if col == hrt else 1.7 + k * 0.9), yy - 1.6), 0.9, 2.2, fc=col, ec="none"))
        arrow(15, yy, 23, 21, col, rad=0.15 if i == 0 else -0.15)

    box(23, 15, 17, 12, tr("ISR"), tr("первинна\nобробка, Q0"), acc)
    arrow(40.5, 24, 51, 32, hrt)
    arrow(40.5, 18, 51, 10, srt)
    queue(51.5, 30.2, 4, hrt)
    queue(51.5, 8.2, 4, srt)
    ax.text(51.5, 36.2, tr("QH · ЖРЧ – вищий пріоритет"), fontsize=7.5, color=hrt)
    ax.text(51.5, 4.6, tr("QS · МРЧ – нижчий пріоритет"), fontsize=7.5, color=srt)
    arrow(65, 32, 77, 24.5, hrt)
    arrow(65, 10, 77, 17.5, srt)
    box(77, 14, 21, 14, tr("Процесор"), tr("вторинна\nобробка (WCET)"), muted)
    ax.text(87.5, 9.5, tr("нова подія перериває\nвторинну обробку"), ha="center", va="top", fontsize=7, color=muted,
            linespacing=1.3)
