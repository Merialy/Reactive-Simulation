"""
Логіка нового інтерфейсу, що НЕ залежить від Qt (її можна перевіряти тестами):
  * читання і запис Excel-файлу параметрів (8 рядків × стовпці-сценарії);
  * перевірка параметрів перед запуском;
  * форматування чисел;
  * експорт результатів в Excel.
"""
import math
import os
import re
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import pandas as pd

from config import Config
from simulation.validation import check_params, utilization
from ui.i18n import is_en, tr, tr_msg

PARAM_KEYS = Config.PARAM_NAMES                       # lambda1, s1, s2, N, D1, lambda2, s2b, D2
PARAM_SHORT = [tr("λ1, мс⁻¹"), tr("s1, мс"), tr("s2, мс"), "N", tr("D1, мс"), tr("λ2, мс⁻¹"), tr("s2b, мс"), tr("D2, мс")]
PARAM_HINT = [tr(Config.PARAM_BOUNDS[i][2]) for i in range(8)]
METRIC_NAMES = [tr(n) for n in Config.METRIC_NAMES]       # повні назви 18 показників
PARAMETER_NAMES = [tr(n) for n in Config.PARAMETER_NAMES]


# ---------------------------------------------------------------------------
# Форматування
# ---------------------------------------------------------------------------
def fmt(value, sig: int = 4) -> str:
    """Число з sig значущими цифрами, без «e» для звичайних величин; кома як роздільник."""
    if value is None:
        return ""
    try:
        v = float(value)
    except (TypeError, ValueError):
        return str(value)
    if math.isnan(v):
        return ""
    if v == 0:
        return "0"
    if v.is_integer() and abs(v) < 1e9:
        return f"{int(v)}"
    a = abs(v)
    if 1e-7 <= a < 1e7:
        digits = max(0, sig - 1 - int(math.floor(math.log10(a))))
        s = f"{v:.{digits}f}"
        if "." in s:
            s = s.rstrip("0").rstrip(".")
        return s if is_en() else s.replace(".", ",")
    s = f"{v:.{sig - 1}e}"
    return s if is_en() else s.replace(".", ",")


def fmt_ci(mean, ci, sig: int = 4) -> str:
    if mean is None:
        return ""
    if not ci:
        return fmt(mean, sig)
    return f"{fmt(mean, sig)} ± {fmt(ci, 2)}"


def comma(text: str) -> str:
    """Десяткова кома в текстах повідомлень українською: «U = 1.104» → «U = 1,104»."""
    return text if is_en() else re.sub(r"(\d)\.(\d)", r"\1,\2", text)


def cls_name(cls) -> str:
    """Коротка назва класу подій: ЖРЧ / МРЧ (англійською HRT / SRT)."""
    return tr("ЖРЧ") if cls == "HRT" else tr("МРЧ")


def type_label(cls, name: str, sep: str = " · ") -> str:
    """
    Підпис типу подій «клас · назва». Типи зі стовпців Excel називаються просто «МРЧ»/«ЖРЧ» –
    тоді показуємо лише клас (і перекладаємо його), щоб не було «МРЧ · МРЧ» чи «SRT · МРЧ».
    """
    c = cls_name(cls)
    n = tr(name) if name in ("МРЧ", "ЖРЧ") else name
    return c if n == c else f"{c}{sep}{n}"


def parse_number(text) -> Optional[float]:
    """'0,5' / '0.5' / 0.5 → 0.5; порожнє → None; нечислове → ValueError."""
    if text is None:
        return None
    if isinstance(text, (int, float)):
        return None if (isinstance(text, float) and math.isnan(text)) else float(text)
    s = str(text).strip().replace(",", ".").replace(" ", "").replace(" ", "")
    if s == "":
        return None
    return float(s)


# ---------------------------------------------------------------------------
# Excel-файл параметрів
# ---------------------------------------------------------------------------
class ParamTable:
    """
    Таблиця параметрів: values[рядок 0..7][стовпець] – число або None.
    Стовпець = сценарій. Формат файлу такий самий, як у версіях 1–3:
    перший стовпець – підписи, далі – по одному сценарію в стовпці.
    """

    def __init__(self, columns: int = 3):
        self.values: List[List[Optional[float]]] = [[None] * columns for _ in range(8)]
        self.labels: List[str] = list(PARAMETER_NAMES)
        self.path: Optional[str] = None
        self.row_format = False                 # файл був у «рядковому» форматі

    # ---------- розміри ----------
    @property
    def n_cols(self) -> int:
        return len(self.values[0]) if self.values else 0

    def add_column(self, copy_from: Optional[int] = None):
        for r in range(8):
            v = self.values[r][copy_from] if copy_from is not None and 0 <= copy_from < self.n_cols else None
            self.values[r].append(v)

    def remove_columns(self, cols: List[int]):
        keep = [c for c in range(self.n_cols) if c not in set(cols)]
        self.values = [[row[c] for c in keep] for row in self.values]
        if self.n_cols == 0:
            self.add_column()

    # ---------- читання / запис ----------
    @staticmethod
    def load(path: str) -> Tuple["ParamTable", List[str]]:
        """
        Повертає (таблиця, список проблем у файлі).

        Підтримуються два формати:
          * основний – параметри в РЯДКАХ (8 рядків), сценарії в стовпцях, перший стовпець – підписи;
          * «рядковий» – перший рядок містить назви параметрів (lambda1, s1, s2, N, D1, lambda2, s2b, D2),
            кожен наступний рядок – окремий сценарій. Порядок стовпців може бути довільним.
        """
        df = pd.read_excel(path, header=None).astype(object)
        problems = []
        t = ParamTable(columns=0)
        t.path = path
        header_map = _header_columns(df)
        if header_map is not None:                       # «рядковий» формат
            t.row_format = True
            cols = []
            for r in range(1, df.shape[0]):
                col = []
                for k in range(8):
                    raw = df.iloc[r, header_map[k]]
                    try:
                        col.append(parse_number(raw))
                    except ValueError:
                        problems.append(tr("Рядок {r}, «{k}»: «{raw}» – не число, комірку очищено.", r=r + 1, k=PARAM_KEYS[k], raw=raw))
                        col.append(None)
                if any(v is not None for v in col):
                    cols.append(col)
            problems.insert(0, tr("Файл у «рядковому» форматі (сценарії – у рядках). Його перетворено: "
                               "тепер кожен сценарій – окремий стовпець таблиці."))
        else:
            if df.shape[0] < 8 or df.shape[1] < 2:
                raise ValueError(tr("Файл має неправильний формат: потрібно 8 рядків параметрів і хоча б один "
                                 "стовпець-сценарій (або перший рядок з назвами параметрів lambda1, s1, s2, N, "
                                 "D1, lambda2, s2b, D2)."))
            t.labels = [str(df.iloc[r, 0]) if pd.notna(df.iloc[r, 0]) else PARAMETER_NAMES[r]
                        for r in range(8)]
            cols = []
            for c in range(1, df.shape[1]):
                col = []
                for r in range(8):
                    raw = df.iloc[r, c]
                    try:
                        col.append(parse_number(raw))
                    except ValueError:
                        problems.append(tr("Стовпець {c} (сценарій {n}), рядок {r}: «{raw}» – не число, комірку очищено.",
                                           c=c, n=len(cols) + 1, r=r + 1, raw=raw))
                        col.append(None)
                if any(v is not None for v in col):
                    cols.append(col)
        if not cols:
            raise ValueError(tr("У файлі немає жодного сценарію з даними."))
        t.values = [[col[r] for col in cols] for r in range(8)]
        return t, problems

    def save(self, path: str):
        rows = []
        for r in range(8):
            rows.append([self.labels[r]] + [("" if v is None else v) for v in self.values[r]])
        pd.DataFrame(rows).to_excel(path, header=False, index=False)
        self.path = path

    # ---------- значення ----------
    def set(self, row: int, col: int, value: Optional[float]):
        self.values[row][col] = value

    def bounds_problem(self, row: int, value: Optional[float]) -> Optional[str]:
        if value is None:
            return None
        lo, hi, name = Config.PARAM_BOUNDS[row]
        if not (lo <= value <= hi):
            return tr("{name}: допустимо від {lo} до {hi}", name=tr(name), lo=fmt(lo), hi=fmt(hi))
        if row == 3 and not float(value).is_integer():
            return tr("N має бути цілим числом")
        return None

    def column_params(self, col: int) -> Optional[Dict]:
        """Словник параметрів стовпця або None, якщо стовпець порожній/неповний."""
        vals = [self.values[r][col] for r in range(8)]
        if all(v is None for v in vals):
            return None
        if any(v is None for v in vals):
            return None
        p = dict(zip(PARAM_KEYS, vals))
        p["N"] = int(p["N"])
        return p

    def column_u(self, col: int) -> Optional[float]:
        p = self.column_params(col)
        if not p:
            return None
        try:
            return utilization(p)
        except (TypeError, ZeroDivisionError):
            return None

    def source_name(self) -> str:
        return os.path.basename(self.path) if self.path else tr("таблиця")


_ALIASES = {
    "lambda1": 0, "λ1": 0, "l1": 0, "s1": 1, "s2": 2, "n": 3, "d1": 4, "deadline1": 4,
    "lambda2": 5, "λ2": 5, "l2": 5, "s2b": 6, "d2": 7, "deadline2": 7,
}


def _header_columns(df) -> Optional[List[int]]:
    """Якщо перший рядок містить назви всіх 8 параметрів – номери їхніх стовпців, інакше None."""
    found = {}
    for c in range(df.shape[1]):
        v = df.iloc[0, c]
        if isinstance(v, str):
            key = v.strip().lower().replace(" ", "").split(",")[0]
            if key in _ALIASES and _ALIASES[key] not in found:
                found[_ALIASES[key]] = c
    return [found[k] for k in range(8)] if len(found) == 8 else None


def prepare_batch(table: ParamTable) -> Tuple[List[Optional[Dict]], Dict[int, List[str]], Dict[int, str]]:
    """
    Перевіряє всі стовпці перед запуском.
    Повертає (run_list, warnings{стовпець: [..]}, errors{стовпець: текст}).
    Стовпець з помилкою отримує None у run_list і не моделюється.
    """
    run_list, warns, errs = [], {}, {}
    for c in range(table.n_cols):
        vals = [table.values[r][c] for r in range(8)]
        if all(v is None for v in vals):
            run_list.append(None)
            continue
        missing = [PARAM_SHORT[r] for r in range(8) if vals[r] is None]
        bound = [table.bounds_problem(r, vals[r]) for r in range(8)]
        bound = [b for b in bound if b]
        if missing or bound:
            errs[c] = "; ".join(([tr("не заповнено: {m}", m=", ".join(missing))] if missing else []) + bound)
            run_list.append(None)
            continue
        p = table.column_params(c)
        e, w = check_params(p)
        if e:
            errs[c] = "; ".join(e)
            run_list.append(None)
            continue
        if w:
            warns[c] = w
        run_list.append(p)
    return run_list, warns, errs


# ---------------------------------------------------------------------------
# Експорт в Excel
# ---------------------------------------------------------------------------
def _meta_rows(settings: Dict, source: str, extra: Optional[List[Tuple[str, str]]] = None):
    rows = [(tr("Програма"), tr("Симулятор моделі реактивності {v}", v=Config.APP_VERSION)),
            (tr("Дата"), datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            (tr("Джерело"), source),
            (tr("Кількість прогонів R"), settings.get("replications")),
            ("Seed", settings.get("seed")),
            (tr("Довірчий інтервал"), tr("95 %, розподіл Стьюдента"))]
    return rows + list(extra or [])


def export_batch_excel(path: str, table: ParamTable, batch: List[Optional[Dict]], settings: Dict,
                       warnings: Dict[int, List[str]], errors: Dict[int, str]):
    """Звіт для серії сценаріїв з Excel: параметри, середні, ДІ, попередження, метадані."""
    names = [tr("Сценарій {n}", n=c + 1) for c in range(table.n_cols)]
    params = pd.DataFrame([[table.values[r][c] for c in range(table.n_cols)] for r in range(8)],
                          index=PARAMETER_NAMES, columns=names)
    mean = pd.DataFrame(index=[f"{i}. {n}" for i, n in enumerate(METRIC_NAMES, 1)], columns=names)
    ci = mean.copy()
    for c, res in enumerate(batch):
        if res is None:
            continue
        for m in range(1, 19):
            mean.iloc[m - 1, c] = res["mean"][m]
            ci.iloc[m - 1, c] = res["ci"][m]
    notes = [(names[c], tr("ПОМИЛКА"), t) for c, t in errors.items()] + \
            [(names[c], tr("Попередження"), w) for c, ws in warnings.items() for w in ws]
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        params.to_excel(xw, sheet_name=tr("Параметри"))
        mean.to_excel(xw, sheet_name=tr("Результати"))
        ci.to_excel(xw, sheet_name=tr("Довірчі_інтервали_95"))
        pd.DataFrame(notes, columns=[tr("Сценарій"), tr("Тип"), tr("Текст")]).to_excel(xw, sheet_name=tr("Попередження"), index=False)
        pd.DataFrame(_meta_rows(settings, table.source_name()), columns=[tr("Параметр"), tr("Значення")]) \
            .to_excel(xw, sheet_name=tr("Метадані"), index=False)
        _autowidth(xw)


PER_TYPE_ROWS = [
    ("count", tr("Кількість оброблених подій")),
    ("wait_primary_mean", tr("Сер. очікування первинної обробки (ISR), мс")),
    ("wait_secondary_mean", tr("Сер. очікування вторинної обробки, мс")),
    ("sojourn_mean", tr("Сер. час реакції, мс")),
    ("sojourn_p95", tr("Час реакції, 95-й перцентиль, мс")),
    ("sojourn_p99", tr("Час реакції, 99-й перцентиль, мс")),
    ("sojourn_max", tr("Макс. час реакції, мс")),
    ("p_miss", tr("Ймовірність порушення дедлайну")),
    ("utilization_secondary", tr("Завантаження вторинним обробником")),
]
FAULT_ROWS = [
    ("count_fault", tr("Кількість подій під час відмови")),
    ("sojourn_mean_fault", tr("Сер. час реакції під час відмови, мс")),
    ("sojourn_max_fault", tr("Макс. час реакції під час відмови, мс")),
    ("p_miss_fault", tr("Ймовірність порушення дедлайну під час відмови")),
]


def export_scenario_excel(path: str, scenario, res: Dict, settings: Dict):
    """Звіт для сценарію з кількома типами подій (і відмовами)."""
    names = [et.name for et in scenario.event_types]
    types = pd.DataFrame([{tr("Тип"): type_label(et.cls, et.name), tr("Клас"): cls_name(et.cls),
                           tr("Інтенсивність, 1/мс"): et.rate, tr("WCET + T_core, мс"): et.service_ms(scenario.platform),
                           tr("Дедлайн, мс"): et.deadline_ms} for et in scenario.event_types])
    rows = PER_TYPE_ROWS + (FAULT_ROWS if scenario.active_faults() else [])
    mean = pd.DataFrame(index=[r[1] for r in rows], columns=names)
    ci = mean.copy()
    for j, n in enumerate(names):
        for i, (key, _) in enumerate(rows):
            v = res["per_type"][n].get(key)
            if v:
                mean.iloc[i, j] = v["mean"]
                ci.iloc[i, j] = v["ci"]
    extra = [(tr("Сценарій"), scenario.name), (tr("U (без відмов)"), round(scenario.utilization(), 4))]
    if scenario.active_faults():
        extra.append((tr("U під час відмов"), round(scenario.utilization_during_faults(), 4)))
        extra += [(tr("Відмова"), tr_msg(f.label())) for f in scenario.active_faults()]
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        types.to_excel(xw, sheet_name=tr("Типи_подій"), index=False)
        mean.to_excel(xw, sheet_name=tr("Результати"))
        ci.to_excel(xw, sheet_name=tr("Довірчі_інтервали_95"))
        pd.DataFrame(_meta_rows(settings, scenario.name, extra), columns=[tr("Параметр"), tr("Значення")]) \
            .to_excel(xw, sheet_name=tr("Метадані"), index=False)
        _autowidth(xw)


def _autowidth(xw):
    for ws in xw.book.worksheets:
        for col in ws.columns:
            width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
            ws.column_dimensions[col[0].column_letter].width = min(60, max(10, width + 2))


# Короткі назви 18 показників для таблиць і графіків (повні – Config.METRIC_NAMES)
METRIC_SHORT = [
    tr("Сер. очікування ISR"), tr("Сер. очікування вторинної, МРЧ"), tr("Сер. очікування вторинної, ЖРЧ"),
    tr("Сер. час перебування, МРЧ"), tr("Сер. час перебування, ЖРЧ"), tr("Макс. очікування ISR"),
    tr("Макс. очікування вторинної, МРЧ"), tr("Макс. очікування вторинної, ЖРЧ"), tr("Макс. час перебування, МРЧ"),
    tr("Макс. час перебування, ЖРЧ"), tr("Сер. довжина черги Q0"), tr("Сер. довжина черги QS (МРЧ)"),
    tr("Сер. довжина черги QH (ЖРЧ)"), tr("Завантаження ISR"), tr("Завантаження вторинної, МРЧ"),
    tr("Завантаження вторинної, ЖРЧ"), tr("P порушення дедлайну, МРЧ"), tr("P порушення дедлайну, ЖРЧ"),
]
