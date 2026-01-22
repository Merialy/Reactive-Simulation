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
        """
        if N <= 0:
            return {i: 0.0 for i in range(1, 19)}
        
        # Ініціалізація масивів для подій типу 1
        arr1 = np.zeros(N + 1)
        start_prim1 = np.zeros(N + 1)
        end_prim1 = np.zeros(N + 1)
        start_sec1 = np.zeros(N + 1)
        end_sec1 = np.zeros(N + 1)
        rem_sec1 = np.zeros(N + 1)
        
        # Ініціалізація масивів для подій типу 2 (динамічна)
        cap2 = max(1024, int(4 * N * lambda2 / lambda1))
        cnt2 = 0
        arr2 = np.zeros(cap2)
        start_prim2 = np.zeros(cap2)
        end_prim2 = np.zeros(cap2)
        start_sec2 = np.zeros(cap2)
        end_sec2 = np.zeros(cap2)
        rem_sec2 = np.zeros(cap2)
        
        # Генерація часу прибуття для типу1 (процес Пуассона)
        t = 0.0
        for i in range(1, N + 1):
            r = self.rng.random()
            if r == 0:
                r = 1e-7
            t = t - np.log(r) / lambda1
            arr1[i] = t
        
        # Налаштування детермінованих прибуттів типу 2
        next_arr1 = arr1[1]
        next_arr2_scheduled = 1.0 / lambda2 if lambda2 > 0 else 1e30
        
        # Відстеження подій
        next_prim_done = 1e30
        next_sec_done = 1e30
        state = "idle"  # "простій", "основний", "додатковий"
        
        # Черги
        primary_q = deque()
        secondary_q1 = deque()
        secondary_q2 = deque()
        
        # Поточна послуга
        cur_prim_type = 0
        cur_prim_idx = 0
        cur_sec_type = 0
        cur_sec_idx = 0
        
        # Журнал подій
        ev_times = []
        ev_pq = []
        ev_sq1 = []
        ev_sq2 = []
        ev_busy_p = []
        ev_busy_s1 = []
        ev_busy_s2 = []
        
        arrivals1 = 0
        total_completed = 0
        
        # Головний цикл подій
        while (arrivals1 < N or state != "idle" or 
               len(primary_q) > 0 or len(secondary_q1) > 0 or len(secondary_q2) > 0):
            
            # Визначення наступної події
            t_next = 1e30
            ev = ""
            
            if arrivals1 < N and next_arr1 < t_next:
                t_next = next_arr1
                ev = "arr1"
            
            if next_arr2_scheduled < t_next:
                t_next = next_arr2_scheduled
                ev = "arr2"
            
            if next_prim_done < t_next:
                t_next = next_prim_done
                ev = "prim_done"
            
            if next_sec_done < t_next:
                t_next = next_sec_done
                ev = "sec_done"
            
            # Час просування
            t = t_next
            
            # Подія процесу
            if ev == "arr1":
                arrivals1 += 1
                
                # Запланувати наступне прибуття події типу 1
                if arrivals1 < N:
                    next_arr1 = arr1[arrivals1 + 1]
                else:
                    next_arr1 = 1e30
                
                # Обробка переривання та обслуговування
                if state == "primary":
                    primary_q.append((1, arrivals1))
                elif state == "secondary":
                    # Перервати вторинну обробку, повернути її в чергу
                    rem_time_cur = max(0, next_sec_done - t)
                    if cur_sec_type == 1:
                        rem_sec1[cur_sec_idx] = rem_time_cur
                        if len(secondary_q1) > 0:
                            secondary_q1.appendleft((1, cur_sec_idx))
                        else:
                            secondary_q2.append((1, cur_sec_idx))
                    elif cur_sec_type == 2:
                        rem_sec2[cur_sec_idx] = rem_time_cur
                        if len(secondary_q2) > 0:
                            secondary_q2.appendleft((2, cur_sec_idx))
                        else:
                            secondary_q2.append((2, cur_sec_idx))
                    
                    cur_sec_type = 0
                    cur_sec_idx = 0
                    next_sec_done = 1e30
                    
                    # Почати первинну обробку
                    state = "primary"
                    cur_prim_type = 1
                    cur_prim_idx = arrivals1
                    start_prim1[arrivals1] = t
                    next_prim_done = t + s1
                else:  # система в стані очікування
                    state = "primary"
                    cur_prim_type = 1
                    cur_prim_idx = arrivals1
                    start_prim1[arrivals1] = t
                    next_prim_done = t + s1
            
            elif ev == "arr2":
                # Створити нове прибуття події типу 2
                cnt2 += 1
                if cnt2 >= len(arr2):
                    # Розширити масиви
                    new_cap = len(arr2) * 2
                    arr2 = np.resize(arr2, new_cap)
                    start_prim2 = np.resize(start_prim2, new_cap)
                    end_prim2 = np.resize(end_prim2, new_cap)
                    start_sec2 = np.resize(start_sec2, new_cap)
                    end_sec2 = np.resize(end_sec2, new_cap)
                    rem_sec2 = np.resize(rem_sec2, new_cap)
                
                arr2[cnt2] = t
                
                # Запланувати наступне детерміністичне прибуття
                if lambda2 > 0:
                    next_arr2_scheduled = t + 1.0 / lambda2
                else:
                    next_arr2_scheduled = 1e30
                
                # Обробити аналогічно до arr1
                if state == "primary":
                    primary_q.append((2, cnt2))
                elif state == "secondary":
                    rem_time_cur2 = max(0, next_sec_done - t)
                    if cur_sec_type == 1:
                        rem_sec1[cur_sec_idx] = rem_time_cur2
                        if len(secondary_q1) > 0:
                            secondary_q1.appendleft((1, cur_sec_idx))
                        else:
                            secondary_q2.append((1, cur_sec_idx))
                    elif cur_sec_type == 2:
                        rem_sec2[cur_sec_idx] = rem_time_cur2
                        if len(secondary_q2) > 0:
                            secondary_q2.appendleft((2, cur_sec_idx))
                        else:
                            secondary_q2.append((2, cur_sec_idx))
                    
                    cur_sec_type = 0
                    cur_sec_idx = 0
                    next_sec_done = 1e30
                    
                    state = "primary"
                    cur_prim_type = 2
                    cur_prim_idx = cnt2
                    start_prim2[cnt2] = t
                    next_prim_done = t + s1
                else:
                    state = "primary"
                    cur_prim_type = 2
                    cur_prim_idx = cnt2
                    start_prim2[cnt2] = t
                    next_prim_done = t + s1
            
            elif ev == "prim_done":
                # Завершити первинну обробку
                if cur_prim_type == 1:
                    end_prim1[cur_prim_idx] = t
                    if rem_sec1[cur_prim_idx] <= 0:
                        rem_sec1[cur_prim_idx] = s2
                    secondary_q1.append((1, cur_prim_idx))
                elif cur_prim_type == 2:
                    end_prim2[cur_prim_idx] = t
                    if rem_sec2[cur_prim_idx] <= 0:
                        rem_sec2[cur_prim_idx] = s2b
                    secondary_q2.append((2, cur_prim_idx))
                
                cur_prim_type = 0
                cur_prim_idx = 0
                next_prim_done = 1e30
                
                # Вибрати наступну дію: primary_q -> S2 -> S1 -> idle
                state, cur_prim_type, cur_prim_idx, cur_sec_type, cur_sec_idx, next_prim_done, next_sec_done = \
                    self._schedule_next(t, primary_q, secondary_q1, secondary_q2, 
                                       start_prim1, start_prim2, start_sec1, start_sec2,
                                       rem_sec1, rem_sec2, s1, s2, s2b)
            
            elif ev == "sec_done":
                # Завершити вторинну обробку
                if cur_sec_type == 1:
                    end_sec1[cur_sec_idx] = t
                elif cur_sec_type == 2:
                    end_sec2[cur_sec_idx] = t
                
                cur_sec_type = 0
                cur_sec_idx = 0
                next_sec_done = 1e30
                total_completed += 1
                
                # Вибрати наступну дію
                state, cur_prim_type, cur_prim_idx, cur_sec_type, cur_sec_idx, next_prim_done, next_sec_done = \
                    self._schedule_next(t, primary_q, secondary_q1, secondary_q2,
                                       start_prim1, start_prim2, start_sec1, start_sec2,
                                       rem_sec1, rem_sec2, s1, s2, s2b)
            
            # Записати стан події в журнал
            ev_times.append(t)
            ev_pq.append(len(primary_q))
            ev_sq1.append(len(secondary_q1))
            ev_sq2.append(len(secondary_q2))
            ev_busy_p.append(1 if state == "primary" else 0)
            ev_busy_s1.append(1 if state == "secondary" and cur_sec_type == 1 else 0)
            ev_busy_s2.append(1 if state == "secondary" and cur_sec_type == 2 else 0)
        
        # Обчислити стаціонарний інтервал (відкинути перші 5% і останні 5% подій типу 1)
        i_start = int(np.ceil(N * 0.05))
        i_end = int(np.floor(N * 0.95))
        
        if i_end <= i_start:
            return {i: 0.0 for i in range(1, 19)}
        
        t_start = arr1[i_start]
        t_end = arr1[i_end]
        
        if t_end <= t_start:
            return {i: 0.0 for i in range(1, 19)}
        
        # Обчислити статистику по окремих завданнях
        results = self._calculate_job_statistics(
            i_start, i_end, cnt2, t_start, t_end,
            arr1, arr2, start_prim1, start_prim2, end_prim1, end_prim2,
            start_sec1, start_sec2, end_sec1, end_sec2, D1, D2
        )
        
        # Обчислити часові середні з журналу подій
        time_avg_results = self._calculate_time_averages(
            ev_times, ev_pq, ev_sq1, ev_sq2, ev_busy_p, ev_busy_s1, ev_busy_s2,
            t_start, t_end
        )
        
        results.update(time_avg_results)
        
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
        """Розрахування статистики для окремих завдань."""
        results = {}
        
        # Статистика типу 1
        sum_wp = 0.0
        sum_ws1 = 0.0
        sum_soj1 = 0.0
        max_wp = 0.0
        max_ws1 = 0.0
        max_soj1 = 0.0
        late1 = 0
        n_done1 = 0
        
        for idx in range(i_start, i_end + 1):
            if end_sec1[idx] > 0:
                wp = start_prim1[idx] - arr1[idx]
                ws1 = start_sec1[idx] - end_prim1[idx]
                soj1 = end_sec1[idx] - arr1[idx]
                
                sum_wp += wp
                sum_ws1 += ws1
                sum_soj1 += soj1
                max_wp = max(max_wp, wp)
                max_ws1 = max(max_ws1, ws1)
                max_soj1 = max(max_soj1, soj1)
                
                if soj1 > D1:
                    late1 += 1
                n_done1 += 1
        
        # Статистика типу 2
        sum_ws2 = 0.0
        sum_soj2 = 0.0
        max_ws2 = 0.0
        max_soj2 = 0.0
        late2 = 0
        n_done2 = 0
        
        for jdx in range(100, cnt2 - 99):
            if t_start <= arr2[jdx] <= t_end and end_sec2[jdx] > 0:
                ws2 = start_sec2[jdx] - end_prim2[jdx]
                soj2 = end_sec2[jdx] - arr2[jdx]
                
                sum_ws2 += ws2
                sum_soj2 += soj2
                max_ws2 = max(max_ws2, ws2)
                max_soj2 = max(max_soj2, soj2)
                
                if soj2 > D2:
                    late2 += 1
                n_done2 += 1
        
        # Результати заповнення
        if n_done1 > 0:
            results[1] = sum_wp / n_done1
            results[2] = sum_ws1 / n_done1
            results[4] = sum_soj1 / n_done1
            results[6] = max_wp
            results[7] = max_ws1
            results[9] = max_soj1
            results[17] = late1 / n_done1
        else:
            results[1] = results[2] = results[4] = 0.0
            results[6] = results[7] = results[9] = 0.0
            results[17] = 0.0
        
        if n_done2 > 0:
            results[3] = sum_ws2 / n_done2
            results[5] = sum_soj2 / n_done2
            results[8] = max_ws2
            results[10] = max_soj2
            results[18] = late2 / n_done2
        else:
            results[3] = results[5] = 0.0
            results[8] = results[10] = 0.0
            results[18] = 0.0
        
        return results
    
    def _calculate_time_averages(self, ev_times, ev_pq, ev_sq1, ev_sq2,
                                ev_busy_p, ev_busy_s1, ev_busy_s2,
                                t_start, t_end):
        """Обчислити часові середні статистики з журналу подій."""
        area_p = 0.0
        area_s1 = 0.0
        area_s2 = 0.0
        busy_p = 0.0
        busy_s1 = 0.0
        busy_s2 = 0.0
        
        for k in range(len(ev_times) - 1):
            ta = ev_times[k]
            tb = ev_times[k + 1]
            
            seg_a = max(ta, t_start)
            seg_b = min(tb, t_end)
            
            if seg_b > seg_a:
                dt_seg = seg_b - seg_a
                area_p += ev_pq[k] * dt_seg
                area_s1 += ev_sq1[k] * dt_seg
                area_s2 += ev_sq2[k] * dt_seg
                busy_p += ev_busy_p[k] * dt_seg
                busy_s1 += ev_busy_s1[k] * dt_seg
                busy_s2 += ev_busy_s2[k] * dt_seg
        
        total_t = t_end - t_start
        
        if total_t <= 0:
            return {i: 0.0 for i in range(11, 17)}
        
        return {
            11: area_p / total_t,
            12: area_s1 / total_t,
            13: area_s2 / total_t,
            14: busy_p / total_t,
            15: busy_s1 / total_t,
            16: busy_s2 / total_t
        }

