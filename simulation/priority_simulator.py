import numpy as np # type: ignore
from collections import deque
from typing import List, Tuple, Dict

# ======================================
# ОСНОВНИЙ КЛАС СИМУЛЯЦІЇ
# ======================================
class PriorityQueueSimulation:
    """
    Імітаційна модель системи масового обслуговування з пріоритетами.

    - Два потоки подій:
        * type1 — пуассонівський потік (λ1)
        * type2 — детермінований потік (λ2)
    - Первинна обробка: s1 (спільна, неперервна)
    - Вторинна обробка: s2 — для type1, s2b — для type2; черги S2 > S1
    - Вторинна обробка переривається надходженням будь-якої події
    - стаціонарність: відкидаємо перші 5% і останні 5% по потоку type1
    - точний розрахунок часових середніх на інтервалі [tStart,tEnd] по журналу подій
    """    
    def __init__(self, seed=None):
        """
        Ініціалізуйте симуляцію.

        Аргументи:
            початкове значення: Випадкове початкове значення для відтворюваності. Якщо немає, результати будуть відрізнятися після кожного запуску.
        """
        self.rng = np.random.default_rng(seed)
    
    def simulate_multiple_systems(self, parameters: List[Dict]) -> List[Dict]:
        """
        Запуск моделювання для кількох наборів параметрів.

        Аргументи:
            параметри: Список словників з ключами: lambda1, s1, s2, N, D1, lambda2, s2b, D2

        Повернення:
            Список словників результатів
        """
        all_results = []
        
        for params in parameters:
            results = self.run_simulation_priority2_full(
                lambda1=params['lambda1'],
                s1=params['s1'],
                s2=params['s2'],
                N=params['N'],
                D1=params['D1'],
                lambda2=params['lambda2'],
                s2b=params['s2b'],
                D2=params['D2']
            )
            all_results.append(results)
        
        return all_results
    
    def run_simulation_priority2_full(self, lambda1: float, s1: float, s2: float,
                                     N: int, D1: float, lambda2: float,
                                     s2b: float, D2: float) -> Dict:
        """
        Запустити одне моделювання із заданими параметрами.

        Повертає dict з 18 показниками:
        1-5: Середній час очікування та час перебування
        6-10: Максимальний час очікування та час перебування
        11-13: Середня довжина черги (основна, вторинна1, вторинна2)
        14-16: Коефіцієнти використання (основна, вторинна1, вторинна2)
        17-18: Частка запізнілих завдань

        Реалізація оптимізована за швидкодією (версія 1.2): у головному циклі
        використовуються списки Python і локальні змінні замість масивів NumPy,
        а підсумкові статистики обчислюються векторно. Логіка моделі не змінилась.
        """
        if N <= 0:
            return {i: 0.0 for i in range(1, 19)}
        if lambda1 <= 0:
            raise ValueError("λ1 має бути додатною")
        if lambda2 < 0 or s1 <= 0 or s2 <= 0 or s2b <= 0:
            raise ValueError("λ2 ≥ 0, s1, s2, s2b > 0")
        N = int(N)
        INF = 1e30

        # ---- Моменти надходження подій МРЧ (пуассонівський потік, формула 2.3) ----
        r = self.rng.random(N)
        r[r == 0] = 1e-7
        arr1 = [0.0] + np.cumsum(-np.log(r) / lambda1).tolist()

        # Масиви часових міток (індекс 0 не використовується)
        start_prim1 = [0.0] * (N + 1)
        end_prim1 = [0.0] * (N + 1)
        start_sec1 = [0.0] * (N + 1)
        end_sec1 = [0.0] * (N + 1)
        rem_sec1 = [0.0] * (N + 1)
        # Для подій ЖРЧ – списки, що зростають (кількість заздалегідь невідома)
        arr2 = [0.0]; start_prim2 = [0.0]; end_prim2 = [0.0]
        start_sec2 = [0.0]; end_sec2 = [0.0]; rem_sec2 = [0.0]
        cnt2 = 0

        period2 = 1.0 / lambda2 if lambda2 > 0 else INF
        next_arr1 = arr1[1]
        next_arr2 = period2
        next_prim_done = INF
        next_sec_done = INF

        # Стан процесора: 0 – простій, 1 – первинна обробка, 2 – вторинна обробка
        state = 0
        primary_q = deque(); q1 = deque(); q2 = deque()
        cur_prim_type = cur_prim_idx = 0
        cur_sec_type = cur_sec_idx = 0

        # Журнал подій
        ev_t = []; ev_pq = []; ev_q1 = []; ev_q2 = []; ev_state = []; ev_sec = []
        log_t = ev_t.append; log_pq = ev_pq.append; log_q1 = ev_q1.append
        log_q2 = ev_q2.append; log_state = ev_state.append; log_sec = ev_sec.append

        arrivals1 = 0
        t = 0.0

        while arrivals1 < N or state != 0 or primary_q or q1 or q2:
            # ---- Найближча подія ----
            t_next = INF; ev = 0
            if arrivals1 < N and next_arr1 < t_next:
                t_next = next_arr1; ev = 1
            if next_arr2 < t_next:
                t_next = next_arr2; ev = 2
            if next_prim_done < t_next:
                t_next = next_prim_done; ev = 3
            if next_sec_done < t_next:
                t_next = next_sec_done; ev = 4
            t = t_next

            if ev == 1 or ev == 2:
                # ---- Надходження події МРЧ (1) або ЖРЧ (2) ----
                if ev == 1:
                    arrivals1 += 1
                    idx = arrivals1
                    next_arr1 = arr1[arrivals1 + 1] if arrivals1 < N else INF
                else:
                    cnt2 += 1
                    idx = cnt2
                    arr2.append(t); start_prim2.append(0.0); end_prim2.append(0.0)
                    start_sec2.append(0.0); end_sec2.append(0.0); rem_sec2.append(0.0)
                    next_arr2 = t + period2

                if state == 1:
                    primary_q.append((ev, idx))
                else:
                    if state == 2:
                        # Перервана задача повертається в ГОЛОВУ черги СВОГО класу
                        rem = next_sec_done - t
                        if rem < 0:
                            rem = 0.0
                        if cur_sec_type == 1:
                            rem_sec1[cur_sec_idx] = rem
                            q1.appendleft((1, cur_sec_idx))
                        else:
                            rem_sec2[cur_sec_idx] = rem
                            q2.appendleft((2, cur_sec_idx))
                        cur_sec_type = cur_sec_idx = 0
                        next_sec_done = INF
                    state = 1
                    cur_prim_type = ev; cur_prim_idx = idx
                    if ev == 1:
                        start_prim1[idx] = t
                    else:
                        start_prim2[idx] = t
                    next_prim_done = t + s1

            else:
                if ev == 3:
                    # ---- Завершення первинної обробки ----
                    if cur_prim_type == 1:
                        end_prim1[cur_prim_idx] = t
                        if rem_sec1[cur_prim_idx] <= 0:
                            rem_sec1[cur_prim_idx] = s2
                        q1.append((1, cur_prim_idx))
                    else:
                        end_prim2[cur_prim_idx] = t
                        if rem_sec2[cur_prim_idx] <= 0:
                            rem_sec2[cur_prim_idx] = s2b
                        q2.append((2, cur_prim_idx))
                    cur_prim_type = cur_prim_idx = 0
                    next_prim_done = INF
                else:
                    # ---- Завершення вторинної обробки ----
                    if cur_sec_type == 1:
                        end_sec1[cur_sec_idx] = t
                    else:
                        end_sec2[cur_sec_idx] = t
                    cur_sec_type = cur_sec_idx = 0
                    next_sec_done = INF

                # ---- Вибір наступної роботи: Q0 -> Q_H -> Q_S -> простій ----
                if primary_q:
                    cur_prim_type, cur_prim_idx = primary_q.popleft()
                    if cur_prim_type == 1:
                        start_prim1[cur_prim_idx] = t
                    else:
                        start_prim2[cur_prim_idx] = t
                    next_prim_done = t + s1
                    state = 1
                elif q2:
                    cur_sec_type, cur_sec_idx = q2.popleft()
                    if start_sec2[cur_sec_idx] == 0:
                        start_sec2[cur_sec_idx] = t
                    rem = rem_sec2[cur_sec_idx]
                    rem_sec2[cur_sec_idx] = 0.0
                    next_sec_done = t + (rem if rem > 0 else s2b)
                    state = 2
                elif q1:
                    cur_sec_type, cur_sec_idx = q1.popleft()
                    if start_sec1[cur_sec_idx] == 0:
                        start_sec1[cur_sec_idx] = t
                    rem = rem_sec1[cur_sec_idx]
                    rem_sec1[cur_sec_idx] = 0.0
                    next_sec_done = t + (rem if rem > 0 else s2)
                    state = 2
                else:
                    state = 0

            # ---- Запис стану в журнал ----
            log_t(t); log_pq(len(primary_q)); log_q1(len(q1)); log_q2(len(q2))
            log_state(state); log_sec(cur_sec_type)

        # ---- Стаціонарний інтервал: відкидаємо перші 5 % і останні 5 % подій МРЧ ----
        i_start = int(np.ceil(N * 0.05))
        i_end = int(np.floor(N * 0.95))
        if i_end <= i_start:
            return {i: 0.0 for i in range(1, 19)}
        t_start = arr1[i_start]
        t_end = arr1[i_end]
        if t_end <= t_start:
            return {i: 0.0 for i in range(1, 19)}

        A = lambda lst: np.asarray(lst, dtype=float)
        results = self._calculate_job_statistics(
            i_start, i_end, cnt2, t_start, t_end,
            A(arr1), A(arr2), A(start_prim1), A(start_prim2), A(end_prim1), A(end_prim2),
            A(start_sec1), A(start_sec2), A(end_sec1), A(end_sec2), D1, D2)

        st = np.asarray(ev_state, dtype=np.int8)
        sec = np.asarray(ev_sec, dtype=np.int8)
        results.update(self._calculate_time_averages(
            A(ev_t), A(ev_pq), A(ev_q1), A(ev_q2),
            (st == 1).astype(float), ((st == 2) & (sec == 1)).astype(float),
            ((st == 2) & (sec == 2)).astype(float), t_start, t_end))
        return results

    def _schedule_next(self, t, primary_q, secondary_q1, secondary_q2,
                      start_prim1, start_prim2, start_sec1, start_sec2,
                      rem_sec1, rem_sec2, s1, s2, s2b):
        """Запланувати наступне обслуговування після завершення поточного."""
        state = "idle"
        cur_prim_type = 0
        cur_prim_idx = 0
        cur_sec_type = 0
        cur_sec_idx = 0
        next_prim_done = 1e30
        next_sec_done = 1e30
        
        if len(primary_q) > 0:
            item = primary_q.popleft()
            cur_prim_type, cur_prim_idx = item
            if cur_prim_type == 1:
                start_prim1[cur_prim_idx] = t
            else:
                start_prim2[cur_prim_idx] = t
            next_prim_done = t + s1
            state = "primary"
        elif len(secondary_q2) > 0:
            item = secondary_q2.popleft()
            cur_sec_type, cur_sec_idx = item
            if start_sec2[cur_sec_idx] == 0:
                start_sec2[cur_sec_idx] = t
            rem_t2 = rem_sec2[cur_sec_idx] if rem_sec2[cur_sec_idx] > 0 else s2b
            rem_sec2[cur_sec_idx] = 0
            next_sec_done = t + rem_t2
            state = "secondary"
        elif len(secondary_q1) > 0:
            item = secondary_q1.popleft()
            cur_sec_type, cur_sec_idx = item
            if start_sec1[cur_sec_idx] == 0:
                start_sec1[cur_sec_idx] = t
            rem_t1 = rem_sec1[cur_sec_idx] if rem_sec1[cur_sec_idx] > 0 else s2
            rem_sec1[cur_sec_idx] = 0
            next_sec_done = t + rem_t1
            state = "secondary"
        
        return state, cur_prim_type, cur_prim_idx, cur_sec_type, cur_sec_idx, next_prim_done, next_sec_done
    
    def _calculate_job_statistics(self, i_start, i_end, cnt2, t_start, t_end,
                                  arr1, arr2, start_prim1, start_prim2, end_prim1, end_prim2,
                                  start_sec1, start_sec2, end_sec1, end_sec2, D1, D2):
        """Статистика по окремих подіях (векторний розрахунок, формули 2.21–2.22)."""
        results = {}

        # Події МРЧ з номерами i_start..i_end, що завершили обробку
        sl = slice(i_start, i_end + 1)
        done = end_sec1[sl] > 0
        wp = (start_prim1[sl] - arr1[sl])[done]
        ws1 = (start_sec1[sl] - end_prim1[sl])[done]
        soj1 = (end_sec1[sl] - arr1[sl])[done]
        if wp.size:
            results[1] = float(wp.mean()); results[2] = float(ws1.mean()); results[4] = float(soj1.mean())
            results[6] = max(0.0, float(wp.max())); results[7] = max(0.0, float(ws1.max()))
            results[9] = max(0.0, float(soj1.max()))
            results[17] = float((soj1 > D1).mean())
        else:
            for k in (1, 2, 4, 6, 7, 9, 17):
                results[k] = 0.0

        # Події ЖРЧ з номерами 100..cnt2-100, що надійшли в [t_start, t_end] і завершили обробку
        lo, hi = 100, cnt2 - 99
        if hi > lo:
            sl2 = slice(lo, hi)
            a2 = arr2[sl2]
            m = (a2 >= t_start) & (a2 <= t_end) & (end_sec2[sl2] > 0)
            ws2 = (start_sec2[sl2] - end_prim2[sl2])[m]
            soj2 = (end_sec2[sl2] - a2)[m]
        else:
            ws2 = soj2 = np.empty(0)
        if ws2.size:
            results[3] = float(ws2.mean()); results[5] = float(soj2.mean())
            results[8] = max(0.0, float(ws2.max())); results[10] = max(0.0, float(soj2.max()))
            results[18] = float((soj2 > D2).mean())
        else:
            for k in (3, 5, 8, 10, 18):
                results[k] = 0.0
        return results

    def _calculate_time_averages(self, ev_times, ev_pq, ev_sq1, ev_sq2,
                                 ev_busy_p, ev_busy_s1, ev_busy_s2,
                                 t_start, t_end):
        """Точні середні за часом за журналом подій (формула 2.23), векторно."""
        total_t = t_end - t_start
        if total_t <= 0 or len(ev_times) < 2:
            return {i: 0.0 for i in range(11, 17)}
        ta = ev_times[:-1]; tb = ev_times[1:]
        dt = np.clip(np.minimum(tb, t_end) - np.maximum(ta, t_start), 0.0, None)
        f = lambda x: float(np.dot(x[:-1], dt) / total_t)
        return {11: f(ev_pq), 12: f(ev_sq1), 13: f(ev_sq2),
                14: f(ev_busy_p), 15: f(ev_busy_s1), 16: f(ev_busy_s2)}
