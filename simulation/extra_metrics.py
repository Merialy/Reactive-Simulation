"""
Додаткові показники реактивності з розділу 1 (п. 1.3), які обчислюються за послідовністю подій:

    jitter            – джиттер часу реакції J = R_max − R_min (показник 5)
    max_consecutive   – найбільша кількість пропусків дедлайну поспіль
    weakly_hard       – перевірка обмеження (m, n) у ковзному вікні (показник 14, п. 1.4.4)
    readiness         – коефіцієнт готовності: частка подій, оброблених без жодного очікування в чергах
                        (ні перед ISR, ні перед вторинною обробкою) (показник 18)
    queue_max         – максимальна довжина черги за журналом стану (показник 15)

Функції працюють з масивами NumPy і не залежать від ядра моделі – їх використовують обидва ядра
(simulation/engine.py і simulation/priority_simulator.py).
"""
from typing import Dict

import numpy as np

EPS = 1e-12


def jitter(sojourn: np.ndarray) -> float:
    return float(sojourn.max() - sojourn.min()) if sojourn.size else 0.0


def max_consecutive(late: np.ndarray) -> int:
    """Найдовша серія True поспіль (late – у порядку надходження подій)."""
    if not late.size or not late.any():
        return 0
    x = np.concatenate(([0], late.astype(np.int8), [0]))
    d = np.diff(x)
    starts, ends = np.flatnonzero(d == 1), np.flatnonzero(d == -1)
    return int((ends - starts).max())


def window_misses(late: np.ndarray, n: int) -> np.ndarray:
    """Кількість пропусків дедлайну в кожному вікні з n послідовних подій."""
    if late.size < n:
        return np.empty(0, dtype=int)
    c = np.concatenate(([0], np.cumsum(late.astype(np.int64))))
    return c[n:] - c[:-n]


def weakly_hard(late: np.ndarray, m: int, n: int) -> Dict[str, float]:
    """
    Обмеження (m, n): у будь-якому вікні з n послідовних подій хоча б m оброблено вчасно,
    тобто пропусків не більше n − m.

        wh_worst_miss – найбільша кількість пропусків в одному вікні
        wh_viol       – частка вікон, у яких обмеження порушено
        wh_ok         – 1, якщо обмеження виконано в усіх вікнах прогону, інакше 0
                        (середнє за R прогонів = частка прогонів без порушень)
    """
    w = window_misses(late, n)
    if not w.size:
        return {'wh_worst_miss': float(late.sum()), 'wh_viol': 0.0, 'wh_ok': 1.0}
    bad = w > (n - m)
    return {'wh_worst_miss': float(w.max()), 'wh_viol': float(bad.mean()), 'wh_ok': 0.0 if bad.any() else 1.0}


def readiness(wait_total: np.ndarray) -> float:
    """Частка подій, які не чекали в жодній черзі (показник 18). wait_total = очікування ISR + вторинної."""
    return float((wait_total <= EPS).mean()) if wait_total.size else 0.0


def queue_max(times: np.ndarray, lengths: np.ndarray, t_start: float, t_end: float) -> float:
    """Максимальна довжина черги на інтервалі [t_start; t_end] за журналом стану (показник 15)."""
    if not times.size:
        return 0.0
    m = (times >= t_start) & (times <= t_end)
    # стан, що діяв на момент t_start, теж враховується
    i0 = int(np.searchsorted(times, t_start, side='right')) - 1
    vals = lengths[m]
    if i0 >= 0:
        vals = np.concatenate((vals, [lengths[i0]]))
    return float(vals.max()) if vals.size else 0.0
