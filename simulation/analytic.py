"""
Аналітичні наближення класичної теорії масового обслуговування (розділ 2, п. 2.5.2) і
системні показники з розділу 1 (п. 1.3), що обчислюються за формулами, а не моделюванням.

Наближення середнього часу перебування подій МРЧ у системі для моделі з двома типами подій
(параметри стовпця таблиці: λ1, s1, s2, λ2, s2b). Час обслуговування події = первинна + вторинна обробка:
    S_S = s1 + s2 (МРЧ),   S_H = s1 + s2b (ЖРЧ);   ρ_S = λ1·S_S,  ρ_H = λ2·S_H,  ρ = ρ_S + ρ_H.

    approx1 – Поллачек–Хінчин (2.7), спільна черга FIFO без пріоритетів
    approx2 – Кобгем (2.9), відносний пріоритет ЖРЧ над МРЧ (без витіснення)
    approx3 – формула (2.10), абсолютний пріоритет з дообслуговуванням (потік ЖРЧ вважається пуассонівським)

Якщо ρ ≥ 1, стаціонарного режиму немає і функції повертають None.
"""
from typing import Dict, Optional


def _base(p: Dict):
    l1, l2 = float(p['lambda1']), float(p.get('lambda2', 0.0) or 0.0)
    SS, SH = float(p['s1']) + float(p['s2']), float(p['s1']) + float(p['s2b'])
    return l1, l2, SS, SH, l1 * SS, l2 * SH


def approx1(p: Dict) -> Optional[float]:
    l1, l2, SS, SH, rs, rh = _base(p)
    rho = rs + rh
    if rho >= 1:
        return None
    W = (l1 * SS ** 2 + l2 * SH ** 2) / (2 * (1 - rho))
    return W + SS


def approx2(p: Dict) -> Optional[float]:
    l1, l2, SS, SH, rs, rh = _base(p)
    if rs + rh >= 1:
        return None
    R = (l1 * SS ** 2 + l2 * SH ** 2) / 2
    return R / ((1 - rh) * (1 - rh - rs)) + SS


def approx3(p: Dict) -> Optional[float]:
    l1, l2, SS, SH, rs, rh = _base(p)
    if rs + rh >= 1:
        return None
    R = (l1 * SS ** 2 + l2 * SH ** 2) / 2
    return SS / (1 - rh) + R / ((1 - rh) * (1 - rh - rs))


APPROX = (("approx1", approx1), ("approx2", approx2), ("approx3", approx3))


def approximations(p: Dict) -> Dict[str, Optional[float]]:
    return {k: f(p) for k, f in APPROX}


def deviation(theory: Optional[float], sim: float) -> Optional[float]:
    """Відносне відхилення наближення від імітації, % (формула 2.25)."""
    if theory is None or not sim:
        return None
    return (theory - sim) / sim * 100.0


def system_load(u: float, u_os: float = 0.0) -> Dict[str, float]:
    """
    Системні показники п. 1.3:
        H   – запас обчислювальної потужності, H = 1 − U  (показник 9; U уже містить накладні витрати ОС)
        U_os – частка процесорного часу на накладні витрати ОС (показник 8)
        eta – коефіцієнт ефективності планування η = (U − U_os) / U  (показник 12)
    """
    return {'H': 1.0 - u, 'U_os': u_os, 'eta': (u - u_os) / u if u > 0 else 0.0}


def scenario_os_share(sc) -> float:
    """U_os для сценарію з типами подій: Σ λ_k · T_core (T_core – накладні витрати ОС на одну подію)."""
    core = sc.platform.core_ms()
    return sum(et.rate for et in sc.event_types) * core
