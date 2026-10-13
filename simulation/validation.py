"""
Перевірка вхідних параметрів імітаційної моделі (розділ 2, таблиця 2.3).

Повертає два списки:
  errors   – параметри, з якими моделювання неможливе;
  warnings – параметри, з якими моделювання можливе, але результати
             слід інтерпретувати обережно (перевантаження, недосяжні дедлайни).
"""
from typing import Dict, List, Tuple


def utilization(p: Dict) -> float:
    """Коефіцієнт завантаження процесора, формула (2.17)."""
    return p['lambda1'] * (p['s1'] + p['s2']) + p['lambda2'] * (p['s1'] + p['s2b'])


def check_params(p: Dict) -> Tuple[List[str], List[str]]:
    errors, warnings = [], []

    if p['lambda1'] <= 0:
        errors.append("λ1 має бути більшою за 0.")
    if p['lambda2'] < 0:
        errors.append("λ2 не може бути від'ємною.")
    for key, name in (('s1', 's1'), ('s2', 's2'), ('s2b', 's2b')):
        if p[key] <= 0:
            errors.append(f"{name} має бути більшою за 0.")
    if p['N'] < 1:
        errors.append("N має бути не меншою за 1.")
    if errors:
        return errors, warnings

    u = utilization(p)
    if u >= 1:
        warnings.append(
            f"Завантаження U = {u:.3f} ≥ 1: стаціонарного режиму не існує, "
            "черга МРЧ зростає необмежено, показники МРЧ залежать від N.")
    elif u > 0.95:
        warnings.append(
            f"Завантаження U = {u:.3f} близьке до 1: потрібні більші N і кількість прогонів.")

    if p['D1'] < p['s1'] + p['s2']:
        warnings.append(
            f"D1 = {p['D1']} < s1 + s2 = {p['s1'] + p['s2']:.4g}: дедлайн МРЧ недосяжний, P_miss = 1.")
    if p['D2'] < p['s1'] + p['s2b']:
        warnings.append(
            f"D2 = {p['D2']} < s1 + s2b = {p['s1'] + p['s2b']:.4g}: дедлайн ЖРЧ недосяжний.")
    if p['lambda2'] > 0 and p['s1'] + p['s2b'] >= 1.0 / p['lambda2']:
        warnings.append(
            "s1 + s2b ≥ 1/λ2: подія ЖРЧ не встигає обробитися до надходження наступної.")
    if p['N'] < 1000:
        warnings.append(f"N = {int(p['N'])} < 1000: оцінки показників можуть бути нестійкими.")

    return errors, warnings
