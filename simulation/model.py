"""
Модель сценарію моделювання (версія 2.0).

Сценарій описує:
  * апаратну платформу БО (тактова частота, такти первинного обробника та ядра ОС);
  * довільний список типів подій двох класів – ЖРЧ (HRT) і МРЧ (SRT);
  * налаштування моделювання (N, дедлайни задаються для кожного типу окремо).

Часові параметри можна задавати або безпосередньо в мілісекундах, або в тактах
процесора – тоді вони перераховуються за формулою (2.12): T = C / F_clk.

Сценарій зберігається у форматі JSON (а на етапі 3 – у MongoDB).
"""
from dataclasses import dataclass, field, asdict
from typing import List, Optional, Dict, Tuple
import json

HRT, SRT = "HRT", "SRT"
PERIODIC, POISSON = "periodic", "poisson"
STORM, OVERRUN = "storm", "overrun"
FAULT_KINDS = (STORM, OVERRUN)
BG_NONE, BG_PERIODIC, BG_POISSON, BG_CONTINUOUS = "none", "periodic", "poisson", "continuous"
BG_MODES = (BG_NONE, BG_PERIODIC, BG_POISSON, BG_CONTINUOUS)


def cycles_to_ms(cycles: float, f_clk_mhz: float) -> float:
    """Формула (2.12): T = C / F_clk; F_clk у МГц, результат у мс."""
    return cycles / (f_clk_mhz * 1000.0)


@dataclass
class Platform:
    """Апаратна платформа БО."""
    name: str = "БО"
    f_clk_mhz: Optional[float] = None      # тактова частота, МГц
    c_irq: Optional[float] = None          # такти первинного обробника (ISR)
    c_core: Optional[float] = None         # такти накладних витрат ОС на подію
    t_irq_ms: Optional[float] = None       # або час ISR у мс напряму
    t_core_ms: float = 0.0                 # або накладні витрати ОС у мс напряму

    def irq_ms(self) -> float:
        if self.c_irq is not None and self.f_clk_mhz:
            return cycles_to_ms(self.c_irq, self.f_clk_mhz)
        if self.t_irq_ms is None:
            raise ValueError("Не задано тривалість первинної обробки (t_irq_ms або c_irq + f_clk_mhz)")
        return self.t_irq_ms

    def core_ms(self) -> float:
        if self.c_core is not None and self.f_clk_mhz:
            return cycles_to_ms(self.c_core, self.f_clk_mhz)
        return self.t_core_ms or 0.0


@dataclass
class EventType:
    """Тип зовнішньої події (кортеж (2.15))."""
    name: str
    cls: str                                # "HRT" або "SRT"
    rate: float                             # інтенсивність (SRT) або частота f_max (HRT), подій/мс
    deadline_ms: float
    wcet_ms: Optional[float] = None         # час вторинної обробки, мс
    wcet_cycles: Optional[float] = None     # або в тактах
    arrival: Optional[str] = None           # "periodic" / "poisson"; за замовчуванням HRT – periodic, SRT – poisson
    phase_ms: Optional[float] = None        # момент першої періодичної події (за замовчуванням – період)
    # Обмеження weakly hard real-time (m, n) (п. 1.3, показник 14; п. 1.4.4): з будь-яких n послідовних
    # подій цього типу хоча б m мають бути оброблені до дедлайну. None – обмеження не задано.
    wh_m: Optional[int] = None
    wh_n: Optional[int] = None

    def weakly_hard(self) -> Optional[Tuple[int, int]]:
        if self.wh_m is None or self.wh_n is None:
            return None
        return int(self.wh_m), int(self.wh_n)

    def arrival_law(self) -> str:
        return self.arrival or (PERIODIC if self.cls == HRT else POISSON)

    def service_ms(self, platform: Platform) -> float:
        """Тривалість вторинної обробки разом з накладними витратами ОС (T_Core + T_i)."""
        if self.wcet_cycles is not None and platform.f_clk_mhz:
            t = cycles_to_ms(self.wcet_cycles, platform.f_clk_mhz)
        elif self.wcet_ms is not None:
            t = self.wcet_ms
        else:
            raise ValueError(f"Тип «{self.name}»: не задано WCET (wcet_ms або wcet_cycles)")
        return t + platform.core_ms()


@dataclass
class Fault:
    """
    Сценарій відмови (версія 3.2).

    kind = "storm"   – шторм переривань: на інтервалі [start_ms; start_ms + duration_ms]
                       з'являється додаткове джерело хибних переривань (пуассонівський потік
                       інтенсивністю rate, подій/мс). Кожне таке переривання займає первинний
                       обробник на T_IRQ і, якщо wcet_ms > 0, ще й вторинну обробку класу МРЧ.
                       Хибні переривання не входять у статистику, але забирають час процесора.
    kind = "overrun" – перевищення WCET: подія типу target (або будь-якого типу, якщо target
                       не задано), що надійшла на інтервалі відмови, з імовірністю probability
                       обробляється у factor разів довше за WCET.
    duration_ms = None означає «до кінця моделювання».
    """
    kind: str
    start_ms: float = 0.0
    duration_ms: Optional[float] = None
    rate: float = 0.0                       # storm: інтенсивність хибних переривань, подій/мс
    wcet_ms: float = 0.0                    # storm: вторинна обробка хибного переривання, мс
    target: Optional[str] = None            # overrun: назва типу подій (None – усі типи)
    probability: float = 1.0                # overrun: імовірність перевищення для однієї події
    factor: float = 1.0                     # overrun: у скільки разів зростає час обробки
    name: str = ""
    enabled: bool = True                    # вимкнена відмова зберігається у файлі, але не діє

    def end_ms(self) -> float:
        return float("inf") if self.duration_ms is None else self.start_ms + self.duration_ms

    def active(self, t: float) -> bool:
        return self.start_ms <= t <= self.end_ms()

    def label(self) -> str:
        if self.name:
            return self.name
        when = f"{self.start_ms:g}–{self.end_ms():g} мс" if self.duration_ms is not None else f"від {self.start_ms:g} мс"
        if self.kind == STORM:
            return f"Шторм переривань {self.rate:g} 1/мс ({when})"
        return f"Перевищення WCET ×{self.factor:g}, p = {self.probability:g}, {self.target or 'усі типи'} ({when})"


@dataclass
class Background:
    """
    Фонові задачі (BG, версія 4.6; розділ 1, п. 1.2.2 і 1.4.1): робота з найнижчим пріоритетом – запис логів,
    самотестування пам'яті тощо. Процесор виконує її, лише коли черги Q0, QH і QS порожні; будь-яке
    переривання її витісняє, після чого вона продовжується з місця зупинки (черга QBG).

    mode = "none"       – фонових задач немає;
           "periodic"   – кожні interval_ms з'являється робота тривалістю work_ms;
           "poisson"    – робота з'являється випадково, в середньому раз на interval_ms;
           "continuous" – фонова робота є завжди (процесор ніколи не простоює).
    deadline_ms     – допустимий час виконання однієї фонової роботи (None – без дедлайну).
    t_cs_ms         – час перемикання контексту (показник 11): коли подія витісняє фонову задачу, ISR
                      триває на t_cs довше; повернення до фонової задачі теж коштує t_cs.
    crit_ms         – тривалість однієї критичної секції фонової задачі з вимкненими перериваннями (t_id,
                      показник 10);
    crit_share      – частка часу фонової задачі всередині критичних секцій (0…1). Подія, що надійшла в
                      критичну секцію, чекає її завершення – це час блокування (показник 6).
    """
    mode: str = BG_NONE
    interval_ms: float = 100.0
    work_ms: float = 1.0
    deadline_ms: Optional[float] = None
    t_cs_ms: float = 0.0
    crit_ms: float = 0.0
    crit_share: float = 0.0

    def active(self) -> bool:
        return self.mode in (BG_PERIODIC, BG_POISSON, BG_CONTINUOUS)

    def load(self) -> float:
        """Частка процесора, потрібна фоновим задачам (для «постійно» – невизначена, 0)."""
        if self.mode in (BG_PERIODIC, BG_POISSON) and self.interval_ms > 0:
            return self.work_ms / self.interval_ms
        return 0.0


@dataclass
class Scenario:
    name: str
    platform: Platform
    event_types: List[EventType]
    N: int = 10000                          # кількість подій МРЧ (або всіх подій, якщо МРЧ немає)
    warmup_fraction: float = 0.05           # частка подій, що відкидається на початку і в кінці
    description: str = ""
    faults: List[Fault] = field(default_factory=list)   # сценарії відмов (версія 3.2)
    background: Optional[Background] = None              # фонові задачі (версія 4.6)

    def has_background(self) -> bool:
        return self.background is not None and self.background.active()

    # ---------- допоміжні величини ----------
    def s1(self) -> float:
        return self.platform.irq_ms()

    def utilization(self) -> float:
        """Формула (2.16): U = Σ λ_k (T_IRQ + T_Core + T_k)."""
        s1 = self.s1()
        return sum(et.rate * (s1 + et.service_ms(self.platform)) for et in self.event_types)

    def active_faults(self) -> List["Fault"]:
        return [f for f in self.faults if f.enabled]

    def without_disabled_faults(self) -> "Scenario":
        """Копія сценарію лише з увімкненими відмовами (саме вона моделюється і зберігається в БД)."""
        import copy
        sc = copy.deepcopy(self)
        sc.faults = sc.active_faults()
        return sc

    def utilization_during_faults(self) -> float:
        """Завантаження на інтервалі, коли діють УСІ відмови одночасно (оцінка зверху)."""
        s1 = self.s1()
        u = self.utilization()
        for f in self.active_faults():
            if f.kind == STORM:
                u += f.rate * (s1 + f.wcet_ms)
            elif f.kind == OVERRUN:
                for et in self.event_types:
                    if f.target in (None, et.name):
                        u += et.rate * f.probability * (f.factor - 1) * et.service_ms(self.platform)
        return u

    # ---------- перевірка (таблиця 2.3) ----------
    def check(self) -> Tuple[List[str], List[str]]:
        errors, warnings = [], []
        if not self.event_types:
            errors.append("Сценарій не містить жодного типу подій.")
        try:
            s1 = self.s1()
            if s1 <= 0:
                errors.append("Тривалість первинної обробки має бути більшою за 0.")
        except ValueError as e:
            errors.append(str(e)); s1 = None
        if self.platform.f_clk_mhz is not None and self.platform.f_clk_mhz <= 0:
            errors.append("Тактова частота має бути більшою за 0.")
        names = [et.name for et in self.event_types]
        if len(set(names)) != len(names):
            errors.append("Назви типів подій мають бути унікальними.")
        for et in self.event_types:
            if et.cls not in (HRT, SRT):
                errors.append(f"Тип «{et.name}»: клас має бути HRT або SRT.")
            if et.arrival_law() not in (PERIODIC, POISSON):
                errors.append(f"Тип «{et.name}»: закон надходження має бути periodic або poisson.")
            if et.rate <= 0:
                errors.append(f"Тип «{et.name}»: інтенсивність має бути більшою за 0.")
            if (et.wh_m is None) != (et.wh_n is None):
                errors.append(f"Тип «{et.name}»: для обмеження (m, n) потрібно задати обидва числа.")
            elif et.wh_n is not None and not (1 <= et.wh_m <= et.wh_n <= 10000):
                errors.append(f"Тип «{et.name}»: обмеження (m, n) має задовольняти 1 ≤ m ≤ n ≤ 10000.")
            try:
                if et.service_ms(self.platform) <= 0:
                    errors.append(f"Тип «{et.name}»: WCET має бути більшим за 0.")
            except ValueError as e:
                errors.append(str(e))
        if self.N < 1:
            errors.append("N має бути не меншою за 1.")
        for f in self.faults:
            if f.kind not in FAULT_KINDS:
                errors.append(f"Відмова «{f.label()}»: тип має бути storm або overrun.")
                continue
            if f.start_ms < 0 or (f.duration_ms is not None and f.duration_ms <= 0):
                errors.append(f"Відмова «{f.label()}»: некоректний інтервал дії.")
            if f.kind == STORM and (f.rate <= 0 or f.wcet_ms < 0):
                errors.append(f"Відмова «{f.label()}»: інтенсивність шторму має бути більшою за 0.")
            if f.kind == OVERRUN:
                if f.factor < 1 or not (0 < f.probability <= 1):
                    errors.append(f"Відмова «{f.label()}»: потрібно factor ≥ 1 і 0 < probability ≤ 1.")
                if f.target is not None and f.target not in names:
                    errors.append(f"Відмова «{f.label()}»: тип подій «{f.target}» не знайдено.")
        bg = self.background
        if bg is not None and bg.mode not in BG_MODES:
            errors.append("Фонові задачі: режим має бути none, periodic, poisson або continuous.")
        elif self.has_background():
            if bg.mode != BG_CONTINUOUS and (bg.interval_ms <= 0 or bg.work_ms <= 0):
                errors.append("Фонові задачі: інтервал і обсяг роботи мають бути більшими за 0.")
            if bg.t_cs_ms < 0 or bg.crit_ms < 0 or not (0 <= bg.crit_share <= 1):
                errors.append("Фонові задачі: t_cs ≥ 0, критична секція ≥ 0, частка в критичних секціях від 0 до 1.")
            if bg.deadline_ms is not None and bg.deadline_ms <= 0:
                errors.append("Фонові задачі: дедлайн має бути більшим за 0.")
        if errors:
            return errors, warnings

        u = self.utilization()
        if u >= 1:
            warnings.append(f"Завантаження U = {u:.3f} ≥ 1: стаціонарного режиму не існує, "
                            "черга МРЧ зростає необмежено, показники МРЧ залежать від N.")
        elif u > 0.95:
            warnings.append(f"Завантаження U = {u:.3f} близьке до 1: потрібні більші N і кількість прогонів.")
        for et in self.event_types:
            need = s1 + et.service_ms(self.platform)
            if et.deadline_ms < need:
                warnings.append(f"Тип «{et.name}»: дедлайн {et.deadline_ms:g} мс < s1 + T = {need:.4g} мс – недосяжний.")
            if et.arrival_law() == PERIODIC and need >= 1.0 / et.rate:
                warnings.append(f"Тип «{et.name}»: обробка ({need:.4g} мс) не встигає до наступної події "
                                f"(період {1.0 / et.rate:.4g} мс).")
        if self.active_faults():
            uf = self.utilization_during_faults()
            if uf >= 1:
                warnings.append(f"Під час відмов завантаження сягає U = {uf:.3f} ≥ 1: процесор перевантажений, "
                                "черги зростають, поки діє відмова.")
        if self.has_background() and self.background.mode != BG_CONTINUOUS and u < 1 \
                and u + self.background.load() >= 1:
            warnings.append(f"Завантаження разом з фоновими задачами U + U_BG = {u + self.background.load():.3f} ≥ 1: "
                            "фонова робота не встигає виконуватися, її черга зростає.")
        if self.N < 1000:
            warnings.append(f"N = {self.N} < 1000: оцінки показників можуть бути нестійкими.")
        return errors, warnings

    # ---------- JSON ----------
    def to_dict(self) -> Dict:
        d = asdict(self)
        for e in d["event_types"]:                # без (m, n) – той самий формат і ключ у БД, що й раніше
            if e.get("wh_m") is None and e.get("wh_n") is None:
                e.pop("wh_m", None); e.pop("wh_n", None)
        if not d.get("faults"):
            d.pop("faults", None)            # сценарії без відмов – у форматі версії 2 (той самий ключ у БД)
        if not self.has_background():
            d.pop("background", None)        # без фонових задач – той самий формат і ключ у БД
        d["schema_version"] = 2
        return d

    def to_json(self, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=indent)

    @staticmethod
    def from_dict(d: Dict) -> "Scenario":
        p = Platform(**d.get("platform", {}))
        ets = [EventType(**e) for e in d["event_types"]]
        return Scenario(name=d.get("name", "Сценарій"), platform=p, event_types=ets,
                        N=int(d.get("N", 10000)), warmup_fraction=d.get("warmup_fraction", 0.05),
                        description=d.get("description", ""),
                        faults=[Fault(**f) for f in d.get("faults", []) or []],
                        background=Background(**d["background"]) if d.get("background") else None)

    @staticmethod
    def from_json(text: str) -> "Scenario":
        return Scenario.from_dict(json.loads(text))

    @staticmethod
    def load(path: str) -> "Scenario":
        with open(path, encoding="utf-8") as f:
            return Scenario.from_json(f.read())

    def save(self, path: str) -> None:
        with open(path, "w", encoding="utf-8") as f:
            f.write(self.to_json())

    # ---------- сумісність зі старим форматом (8 параметрів Excel) ----------
    @staticmethod
    def from_legacy(p: Dict, name: str = "Сценарій") -> "Scenario":
        """lambda1, s1, s2, N, D1, lambda2, s2b, D2 → сценарій з одним типом МРЧ і одним типом ЖРЧ."""
        ets = [EventType(name="МРЧ", cls=SRT, rate=p["lambda1"], wcet_ms=p["s2"], deadline_ms=p["D1"])]
        if p["lambda2"] > 0:
            ets.append(EventType(name="ЖРЧ", cls=HRT, rate=p["lambda2"], wcet_ms=p["s2b"], deadline_ms=p["D2"]))
        return Scenario(name=name, platform=Platform(t_irq_ms=p["s1"]), event_types=ets, N=int(p["N"]))
