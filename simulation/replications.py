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

# Квантилі розподілу Стьюдента t(df, 0.975) для 95 % довірчого інтервалу: точні значення для df ≤ 30,
# далі – розклад Корніша–Фішера (похибка < 0,0005 для df > 30).
_T975 = {1: 12.7062, 2: 4.3027, 3: 3.1824, 4: 2.7764, 5: 2.5706, 6: 2.4469, 7: 2.3646, 8: 2.3060, 9: 2.2622,
         10: 2.2281, 11: 2.2010, 12: 2.1788, 13: 2.1604, 14: 2.1448, 15: 2.1314, 16: 2.1199, 17: 2.1098,
         18: 2.1009, 19: 2.0930, 20: 2.0860, 21: 2.0796, 22: 2.0739, 23: 2.0687, 24: 2.0639, 25: 2.0595,
         26: 2.0555, 27: 2.0518, 28: 2.0484, 29: 2.0452, 30: 2.0423}


def t_quantile_975(df: int) -> float:
    """Квантиль t(df, 0.975) розподілу Стьюдента."""
    if df <= 0:
        return float('nan')
    if df in _T975:
        return _T975[df]
    z = 1.959963985
    v = float(df)
    return (z + (z ** 3 + z) / (4 * v) + (5 * z ** 5 + 16 * z ** 3 + 3 * z) / (96 * v ** 2)
            + (3 * z ** 7 + 19 * z ** 5 + 17 * z ** 3 - 15 * z) / (384 * v ** 3))


def check_settings(replications, base_seed) -> None:
    """R ≥ 1 і seed ≥ 0 (генератор NumPy не приймає від'ємних seed)."""
    if int(replications) < 1:
        raise ValueError("Кількість прогонів R має бути не меншою за 1.")
    if int(base_seed) < 0:
        raise ValueError("Seed має бути невід'ємним цілим числом (0, 1, 2, …).")


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
    check_settings(replications, base_seed)
    replications = max(1, int(replications))
    runs = []
    for r in range(replications):
        sim = PriorityQueueSimulation(seed=base_seed + r)
        runs.append(sim.run_simulation_priority2_full(**params))
    mean, ci = {}, {}
    for m in METRIC_IDS:
        s = summarize([run[m] for run in runs])
        mean[m], ci[m] = s['mean'], s['ci']
    return {'mean': mean, 'ci': ci, 'extra': summarize_extra(runs), 'runs': runs}


def summarize_extra(runs) -> Dict:
    """Середнє і ДІ для додаткових показників (рядкові ключі, див. priority_simulator.EXTRA_KEYS).
    Показник, якого немає в жодному прогоні (наприклад, немає подій ЖРЧ), пропускається."""
    from simulation.priority_simulator import EXTRA_KEYS
    out = {}
    for k in EXTRA_KEYS:
        vals = [run[k] for run in runs if k in run]
        if vals:
            s = summarize(vals)
            out[k] = {'mean': s['mean'], 'ci': s['ci']}
    return out


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

    check_settings(replications, base_seed)
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
        out.append({'mean': mean, 'ci': ci, 'extra': summarize_extra(runs[i]), 'runs': runs[i]})
    return out


# ---------------------------------------------------------------------------
# Сценарії з кількома типами подій (версія 2.0)
# ---------------------------------------------------------------------------
def _run_scenario_one(task):
    scenario_dict, seed = task
    from simulation.engine import EventSimulation
    from simulation.model import Scenario
    return EventSimulation(seed=seed).run(Scenario.from_dict(scenario_dict))


def run_scenario_replications(scenario, replications: int = 5, base_seed: int = 42,
                              workers: Optional[int] = None, progress_cb=None) -> Dict:
    """
    R прогонів сценарію з кількома типами подій.

    Повертає {'classic': {'mean', 'ci'}, 'per_type': {тип: {показник: {'mean', 'ci'}}}, 'U', 'runs'}
    (і 'faults', якщо в сценарії є відмови). progress_cb(done, total) – після кожного прогону.
    """
    import os
    from concurrent.futures import ProcessPoolExecutor, as_completed

    check_settings(replications, base_seed)
    replications = max(1, int(replications))
    d = scenario.to_dict()
    tasks = [(d, base_seed + r) for r in range(replications)]
    if workers is None:
        workers = max(1, (os.cpu_count() or 2) - 1)
    workers = min(workers, replications)
    runs = None
    if workers > 1:
        try:
            runs = [None] * replications
            with ProcessPoolExecutor(max_workers=workers) as ex:
                futs = {ex.submit(_run_scenario_one, t): r for r, t in enumerate(tasks)}
                for done, fut in enumerate(as_completed(futs), 1):
                    runs[futs[fut]] = fut.result()
                    if progress_cb:
                        progress_cb(done, replications)
        except (OSError, RuntimeError, ImportError):
            runs = None
    if runs is None:
        runs = []
        for done, t in enumerate(tasks, 1):
            runs.append(_run_scenario_one(t))
            if progress_cb:
                progress_cb(done, replications)

    classic = {'mean': {}, 'ci': {}}
    for m in METRIC_IDS:
        s = summarize([r['classic'][m] for r in runs])
        classic['mean'][m], classic['ci'][m] = s['mean'], s['ci']
    per_type = {}
    for name in runs[0]['per_type']:
        per_type[name] = {}
        for key in runs[0]['per_type'][name]:
            s = summarize([r['per_type'][name][key] for r in runs])
            per_type[name][key] = {'mean': s['mean'], 'ci': s['ci']}
    out = {'classic': classic, 'per_type': per_type, 'U': runs[0]['U'], 'runs': runs}
    if 'system' in runs[0]:
        out['system'] = {key: {'mean': s['mean'], 'ci': s['ci']}
                         for key in runs[0]['system']
                         for s in [summarize([r.get('system', {}).get(key, 0.0) for r in runs])]}
    if 'faults' in runs[0]:
        out['faults'] = {key: {'mean': s['mean'], 'ci': s['ci']}
                         for key in runs[0]['faults']
                         for s in [summarize([r['faults'][key] for r in runs])]}
    return out
