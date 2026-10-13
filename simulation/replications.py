"""
Повторні прогони імітаційної моделі та довірчі інтервали (розділ 2, п. 2.6.5).

Для кожного сценарію виконується R незалежних прогонів з початковими значеннями
генератора base_seed, base_seed + 1, ..., base_seed + R - 1. Ті самі початкові
значення використовуються для ВСІХ сценаріїв серії (метод загальних випадкових
чисел), тому різниця між сценаріями зумовлена лише зміною параметрів.
"""
from typing import Dict, List, Optional
import math
import numpy as np

from simulation.priority_simulator import PriorityQueueSimulation

METRIC_IDS = list(range(1, 19))

# Квантилі розподілу Стьюдента t(df, 0.975) для 95 % довірчого інтервалу
_T975 = {1: 12.706, 2: 4.303, 3: 3.182, 4: 2.776, 5: 2.571, 6: 2.447, 7: 2.365,
         8: 2.306, 9: 2.262, 10: 2.228, 11: 2.201, 12: 2.179, 13: 2.160, 14: 2.145,
         15: 2.131, 16: 2.120, 17: 2.110, 18: 2.101, 19: 2.093, 20: 2.086, 25: 2.060,
         30: 2.042, 40: 2.021, 60: 2.000, 120: 1.980}


def t_quantile_975(df: int) -> float:
    """Квантиль t(df, 0.975); для проміжних df береться найближче менше табличне значення."""
    if df <= 0:
        return float('nan')
    if df in _T975:
        return _T975[df]
    if df > 120:
        return 1.960
    return _T975[max(k for k in _T975 if k <= df)]


def summarize(values: List[float]) -> Dict[str, float]:
    """Середнє, стандартне відхилення та напівширина 95 % довірчого інтервалу (2.24)."""
    a = np.asarray(values, dtype=float)
    r = len(a)
    mean = float(a.mean())
    if r < 2:
        return {'mean': mean, 'std': 0.0, 'ci': 0.0, 'n': r}
    std = float(a.std(ddof=1))
    ci = t_quantile_975(r - 1) * std / math.sqrt(r)
    return {'mean': mean, 'std': std, 'ci': ci, 'n': r}


def run_replications(params: Dict, replications: int = 5, base_seed: int = 42) -> Dict:
    """
    Виконує R прогонів одного сценарію.

    Повертає словник:
      'mean' : {номер_метрики: середнє}
      'ci'   : {номер_метрики: напівширина 95 % ДІ}
      'runs' : список результатів окремих прогонів
    """
    replications = max(1, int(replications))
    runs = []
    for r in range(replications):
        sim = PriorityQueueSimulation(seed=base_seed + r)
        runs.append(sim.run_simulation_priority2_full(**params))
    mean, ci = {}, {}
    for m in METRIC_IDS:
        s = summarize([run[m] for run in runs])
        mean[m], ci[m] = s['mean'], s['ci']
    return {'mean': mean, 'ci': ci, 'runs': runs}


# ---------------------------------------------------------------------------
# Паралельний запуск серії сценаріїв (версія 1.2)
# ---------------------------------------------------------------------------
def _run_one(task):
    """Один прогон одного сценарію. Функція верхнього рівня – потрібна для multiprocessing."""
    scenario_idx, rep_idx, params, seed = task
    result = PriorityQueueSimulation(seed=seed).run_simulation_priority2_full(**params)
    return scenario_idx, rep_idx, result


def run_batch(param_list, replications: int = 5, base_seed: int = 42,
              workers: Optional[int] = None, progress_cb=None) -> List[Optional[Dict]]:
    """
    Виконує R прогонів для кожного сценарію зі списку param_list.

    Прогони розподіляються між процесами (workers). Початкові значення генератора
    не залежать від кількості процесів, тому результат завжди однаковий.
    Для сценаріїв, що дорівнюють None, повертається None.
    progress_cb(done, total) викликається після кожного завершеного прогону.
    """
    import os
    from concurrent.futures import ProcessPoolExecutor, as_completed

    replications = max(1, int(replications))
    tasks = [(i, r, p, base_seed + r)
             for i, p in enumerate(param_list) if p is not None
             for r in range(replications)]
    total = len(tasks)
    runs = {i: [None] * replications for i, p in enumerate(param_list) if p is not None}

    if workers is None:
        workers = max(1, (os.cpu_count() or 2) - 1)
    workers = min(workers, max(1, total))

    done = 0
    if workers > 1 and total > 1:
        try:
            with ProcessPoolExecutor(max_workers=workers) as ex:
                for fut in as_completed([ex.submit(_run_one, t) for t in tasks]):
                    i, r, res = fut.result()
                    runs[i][r] = res
                    done += 1
                    if progress_cb:
                        progress_cb(done, total)
        except (OSError, RuntimeError, ImportError):
            workers = 1        # якщо процеси недоступні – послідовний режим
            done = 0
    if workers <= 1:
        for t in tasks:
            i, r, res = _run_one(t)
            runs[i][r] = res
            done += 1
            if progress_cb:
                progress_cb(done, total)

    out = []
    for i, p in enumerate(param_list):
        if p is None:
            out.append(None)
            continue
        mean, ci = {}, {}
        for m in METRIC_IDS:
            s = summarize([run[m] for run in runs[i]])
            mean[m], ci[m] = s['mean'], s['ci']
        out.append({'mean': mean, 'ci': ci, 'runs': runs[i]})
    return out
