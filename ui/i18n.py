"""
Мова інтерфейсу (українська / англійська).

    tr("Текст українською")                 → переклад або той самий текст
    tr("Сценаріїв: {n}", n=5)               → шаблон з підстановкою

Мова задається ДО імпорту інших модулів інтерфейсу (app.py), бо частина підписів –
константи модулів. Перемикання мови вимагає перезапуску програми.
"""
LANG = "uk"
EN = {}
NAMES = {}


def set_lang(lang: str):
    global LANG
    LANG = "en" if lang == "en" else "uk"
    if LANG == "en" and not EN:
        from ui.i18n_en import EN as _EN, NAMES as _NAMES
        EN.update(_EN)
        NAMES.update(_NAMES)


def tr(text: str, **kw) -> str:
    t = EN.get(text, text) if LANG == "en" else text
    return t.format(**kw) if kw else t


def is_en() -> bool:
    return LANG == "en"


def tr_name(text: str) -> str:
    """Назва з готового сценарію (тип подій, відмова, сценарій) – англійською, якщо вона відома."""
    if LANG != "en" or not text:
        return text
    return NAMES.get(text) or EN.get(text) or text


# ---------------------------------------------------------------------------
# Переклад повідомлень, які формує модель (simulation/validation.py, simulation/model.py,
# storage/repository.py). Вони лишаються українськими в ядрі, а в інтерфейсі перекладаються.
# ---------------------------------------------------------------------------
import re as _re

_MSG = [
    (r"Завантаження U = (\S+) ≥ 1: стаціонарного режиму не існує, черга МРЧ зростає необмежено, показники МРЧ залежать від N\.",
     r"Utilization U = \1 ≥ 1: no steady state exists, the SRT queue grows without bound, SRT metrics depend on N."),
    (r"Завантаження U = (\S+) близьке до 1: потрібні більші N і кількість прогонів\.",
     r"Utilization U = \1 is close to 1: larger N and more replications are needed."),
    (r"D1 = (\S+) < s1 \+ s2 = (\S+): дедлайн МРЧ недосяжний, P_miss = 1\.",
     r"D1 = \1 < s1 + s2 = \2: the SRT deadline is unreachable, P_miss = 1."),
    (r"D2 = (\S+) < s1 \+ s2b = (\S+): дедлайн ЖРЧ недосяжний\.",
     r"D2 = \1 < s1 + s2b = \2: the HRT deadline is unreachable."),
    (r"s1 \+ s2b ≥ 1/λ2: подія ЖРЧ не встигає обробитися до надходження наступної\.",
     r"s1 + s2b ≥ 1/λ2: an HRT event cannot be processed before the next one arrives."),
    (r"N = (\S+) < 1000: оцінки показників можуть бути нестійкими\.",
     r"N = \1 < 1000: metric estimates may be unstable."),
    (r"λ1 має бути більшою за 0\.", r"λ1 must be greater than 0."),
    (r"λ2 не може бути від'ємною\.", r"λ2 cannot be negative."),
    (r"^(s1|s2|s2b|D1|D2|λ1|λ2) має бути більшою за 0\.", r"\1 must be greater than 0."),
    (r"N має бути не меншою за 1\.", r"N must be at least 1."),
    (r"Сценарій не містить жодного типу подій\.", r"The scenario contains no event types."),
    (r"Тривалість первинної обробки має бути більшою за 0\.", r"Primary processing time must be greater than 0."),
    (r"Тактова частота має бути більшою за 0\.", r"Clock frequency must be greater than 0."),
    (r"Назви типів подій мають бути унікальними\.", r"Event type names must be unique."),
    (r"Не задано тривалість первинної обробки \(t_irq_ms або c_irq \+ f_clk_mhz\)",
     r"Primary processing time is not set (t_irq_ms or c_irq + f_clk_mhz)"),
    (r"Тип «(.+?)»: не задано WCET \(wcet_ms або wcet_cycles\)", r"Type “\1”: WCET is not set (wcet_ms or wcet_cycles)"),
    (r"Тип «(.+?)»: клас має бути HRT або SRT\.", r"Type “\1”: class must be HRT or SRT."),
    (r"Тип «(.+?)»: закон надходження має бути periodic або poisson\.", r"Type “\1”: arrival law must be periodic or poisson."),
    (r"Тип «(.+?)»: інтенсивність має бути більшою за 0\.", r"Type “\1”: rate must be greater than 0."),
    (r"Тип «(.+?)»: WCET має бути більшим за 0\.", r"Type “\1”: WCET must be greater than 0."),
    (r"Тип «(.+?)»: дедлайн (\S+) мс < s1 \+ T = (\S+) мс – недосяжний\.",
     r"Type “\1”: deadline \2 ms < s1 + T = \3 ms – unreachable."),
    (r"Тип «(.+?)»: обробка \((\S+) мс\) не встигає до наступної події \(період (\S+) мс\)\.",
     r"Type “\1”: processing (\2 ms) does not finish before the next event (period \3 ms)."),
    (r"Тип «(.+?)»: для обмеження \(m, n\) потрібно задати обидва числа\.", r"Type “\1”: both numbers of the (m, n) constraint are required."),
    (r"Тип «(.+?)»: обмеження \(m, n\) має задовольняти 1 ≤ m ≤ n ≤ 10000\.", r"Type “\1”: the (m, n) constraint must satisfy 1 ≤ m ≤ n ≤ 10000."),
    (r"Відмова «(.+?)»: тип має бути storm або overrun\.", r"Fault “\1”: kind must be storm or overrun."),
    (r"Фонові задачі: режим має бути none, periodic, poisson або continuous\.", r"Background tasks: mode must be none, periodic, poisson or continuous."),
    (r"Фонові задачі: інтервал і обсяг роботи мають бути більшими за 0\.", r"Background tasks: interval and work must be greater than 0."),
    (r"Фонові задачі: t_cs ≥ 0, критична секція ≥ 0, частка в критичних секціях від 0 до 1\.", r"Background tasks: t_cs ≥ 0, critical section ≥ 0, share in critical sections from 0 to 1."),
    (r"Фонові задачі: дедлайн має бути більшим за 0\.", r"Background tasks: deadline must be greater than 0."),
    (r"Завантаження разом з фоновими задачами U \+ U_BG = (\S+) ≥ 1: фонова робота не встигає виконуватися, її черга зростає\.",
     r"Utilization with background tasks U + U_BG = \1 ≥ 1: background work cannot keep up, its queue grows."),
    (r"Відмова «(.+?)»: некоректний інтервал дії\.", r"Fault “\1”: invalid time interval."),
    (r"Відмова «(.+?)»: інтенсивність шторму має бути більшою за 0\.", r"Fault “\1”: storm rate must be greater than 0."),
    (r"Відмова «(.+?)»: потрібно factor ≥ 1 і 0 < probability ≤ 1\.", r"Fault “\1”: factor ≥ 1 and 0 < probability ≤ 1 are required."),
    (r"Відмова «(.+?)»: тип подій «(.+?)» не знайдено\.", r"Fault “\1”: event type “\2” not found."),
    (r"Під час відмов завантаження сягає U = (\S+) ≥ 1: процесор перевантажений, черги зростають, поки діє відмова\.",
     r"During faults utilization reaches U = \1 ≥ 1: the processor is overloaded, queues grow while the fault lasts."),
    (r"MongoDB \((.+), база (.+)\)", r"MongoDB (\1, database \2)"),
    (r"JSON-файли \((.+)\)", r"JSON files (\1)"),
    (r"Шторм переривань (\S+) 1/мс \((.+)\)", r"Interrupt storm \1 1/ms (\2)"),
    (r"Перевищення WCET ×(\S+), p = (\S+), (.+) \((.+)\)", r"WCET overrun ×\1, p = \2, \3 (\4)"),
    (r"від (\S+) мс", r"from \1 ms"),
    (r"(\S+)–(\S+) мс", r"\1–\2 ms"),
    (r"усі типи", r"all types"),
    (r"λ1 має бути додатною", r"λ1 must be positive"),
    (r"Кількість прогонів R має бути не меншою за 1\.", r"The number of replications R must be at least 1."),
    (r"Seed має бути невід'ємним цілим числом \(0, 1, 2, …\)\.", r"Seed must be a non-negative integer (0, 1, 2, …)."),
    (r"MongoDB недоступна \((.+)\): використовуються JSON-файли\.", r"MongoDB is unavailable (\1): JSON files are used."),
]
_MSG = [(_re.compile(a), b) for a, b in _MSG]


def tr_msg(text: str) -> str:
    """Переклад повідомлення ядра моделі (якщо інтерфейс англійською)."""
    if LANG != "en" or not text:
        return text
    out = text
    for pat, rep in _MSG:
        out = pat.sub(rep, out)
    return out
