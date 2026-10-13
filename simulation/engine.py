"""
Узагальнене ядро дискретно-подієвого моделювання (версія 2.0).

Модель з трьома чергами (розділ 2, п. 2.4), узагальнена на довільну кількість типів подій:
  Q0 – спільна черга первинного обробника (FIFO, без витіснення ISR);
  QH – черга вторинних обробників подій ЖРЧ (усі типи класу HRT, FIFO);
  QS – черга вторинних обробників подій МРЧ (усі типи класу SRT, FIFO).
Пріоритет: Q0 → QH → QS. Будь-яка нова подія перериває вторинну обробку; перервана
задача повертається в голову черги свого класу (preemptive-resume).

Для сценарію з одним типом МРЧ і одним типом ЖРЧ результати збігаються з ядром версії 1.x.
"""
from collections import deque
from typing import Dict, List
import numpy as np

from simulation.model import Scenario, HRT, SRT, PERIODIC, POISSON

INF = 1e30
METRICS_PER_TYPE = ("count", "wait_primary_mean", "wait_secondary_mean", "sojourn_mean",
                    "sojourn_max", "sojourn_p95", "sojourn_p99", "p_miss", "utilization_secondary")


class EventSimulation:
    def __init__(self, seed=None):
        self.rng = np.random.default_rng(seed)

    # ------------------------------------------------------------------
    def run(self, sc: Scenario) -> Dict:
        """
        Повертає словник:
          'classic'  – 18 показників у форматі версії 1.x (агреговано за класами МРЧ / ЖРЧ);
          'per_type' – {назва типу: {показник: значення}};
          'U'        – теоретичне завантаження (2.16).
        """
        errors, _ = sc.check()
        if errors:
            raise ValueError("; ".join(errors))

        types = sc.event_types
        K = len(types)
        s1 = sc.s1()
        service = [et.service_ms(sc.platform) for et in types]
        is_hrt = [et.cls == HRT for et in types]
        N = int(sc.N)

        # ---- Пуассонівські потоки: об'єднаний потік з інтенсивністю Λ і вибором типу ----
        pois = [k for k, et in enumerate(types) if et.arrival_law() == POISSON]
        peri = [k for k, et in enumerate(types) if et.arrival_law() == PERIODIC]
        ref_is_srt = any(not is_hrt[k] for k in range(K))

        if pois:
            lam = np.array([types[k].rate for k in pois])
            Lam = float(lam.sum())
            r = self.rng.random(N)
            r[r == 0] = 1e-7
            p_arr = np.cumsum(-np.log(r) / Lam).tolist()
            if len(pois) == 1:
                p_type = [pois[0]] * N
            else:
                p_type = [pois[i] for i in self.rng.choice(len(pois), size=N, p=lam / Lam)]
            n_pois = N
        else:
            p_arr, p_type, n_pois = [], [], 0

        # Періодичні потоки: наступний момент для кожного типу
        next_per = {k: (types[k].phase_ms if types[k].phase_ms is not None else 1.0 / types[k].rate)
                    for k in peri}
        period = {k: 1.0 / types[k].rate for k in peri}

        # ---- Записи подій ----
        ev_type: List[int] = []; arr: List[float] = []
        sp: List[float] = []; ep: List[float] = []
        ss: List[float] = []; es: List[float] = []; rem: List[float] = []
        type_seq: List[List[int]] = [[] for _ in range(K)]   # id подій кожного типу в порядку надходження
        ref_seq: List[int] = []                               # id «опорних» подій (МРЧ) для стаціонарного вікна

        ref_flag = [((not is_hrt[k]) if ref_is_srt else True) for k in range(K)]
        a_type, a_arr, a_sp, a_ep, a_ss, a_es, a_rem = (ev_type.append, arr.append, sp.append, ep.append,
                                                        ss.append, es.append, rem.append)
        a_ref = ref_seq.append
        a_tseq = [type_seq[k].append for k in range(K)]

        # Якщо МРЧ немає, «опорними» є всі події; зупинка – після N подій
        n_ref_target = N
        p_idx = 0
        next_p = p_arr[0] if n_pois else INF

        state = 0                    # 0 – простій, 1 – ISR, 2 – вторинна обробка
        q0 = deque(); qh = deque(); qs = deque()
        cur_p = -1; cur_s = -1
        next_prim_done = INF; next_sec_done = INF

        journal = []
        log = journal.append
        t = 0.0

        n_ref = 0
        n_ev = 0
        per_next = [next_per[k] for k in peri]
        per_period = [period[k] for k in peri]
        n_per = len(peri)

        while n_ref < n_ref_target or state != 0 or q0 or qh or qs:
            # ---- Найближча подія (порядок при рівності: Пуассон, періодичні, ISR, вторинна) ----
            t_next = INF; ev = 0; ev_k = -1
            if n_ref < n_ref_target and p_idx < n_pois and next_p < t_next:
                t_next = next_p; ev = 1
            if n_per == 1:
                if per_next[0] < t_next:
                    t_next = per_next[0]; ev = 2; ev_k = 0
            elif n_per:
                for j in range(n_per):
                    if per_next[j] < t_next:
                        t_next = per_next[j]; ev = 2; ev_k = j
            if next_prim_done < t_next:
                t_next = next_prim_done; ev = 3
            if next_sec_done < t_next:
                t_next = next_sec_done; ev = 4
            if ev == 0:
                break                 # немає подій (наприклад, лише періодичні типи і N досягнуто)
            t = t_next

            if ev == 1 or ev == 2:
                if ev == 1:
                    k = p_type[p_idx]
                    p_idx += 1
                    next_p = p_arr[p_idx] if p_idx < n_pois else INF
                else:
                    k = peri[ev_k]
                    per_next[ev_k] = t + per_period[ev_k]
                i = n_ev
                n_ev += 1
                a_type(k); a_arr(t); a_sp(0.0); a_ep(0.0); a_ss(0.0); a_es(0.0); a_rem(0.0)
                a_tseq[k](i)
                if ref_flag[k]:
                    a_ref(i)
                    n_ref += 1

                if state == 1:
                    q0.append(i)
                else:
                    if state == 2:
                        rr = next_sec_done - t
                        rem[cur_s] = rr if rr > 0 else 0.0
                        (qh if is_hrt[ev_type[cur_s]] else qs).appendleft(cur_s)
                        cur_s = -1
                        next_sec_done = INF
                    state = 1
                    cur_p = i
                    sp[i] = t
                    next_prim_done = t + s1
            else:
                if ev == 3:
                    i = cur_p
                    ep[i] = t
                    if rem[i] <= 0:
                        rem[i] = service[ev_type[i]]
                    (qh if is_hrt[ev_type[i]] else qs).append(i)
                    cur_p = -1
                    next_prim_done = INF
                else:
                    es[cur_s] = t
                    cur_s = -1
                    next_sec_done = INF

                if q0:
                    cur_p = q0.popleft()
                    sp[cur_p] = t
                    next_prim_done = t + s1
                    state = 1
                elif qh or qs:
                    cur_s = qh.popleft() if qh else qs.popleft()
                    if ss[cur_s] == 0:
                        ss[cur_s] = t
                    rr = rem[cur_s]
                    rem[cur_s] = 0.0
                    next_sec_done = t + (rr if rr > 0 else service[ev_type[cur_s]])
                    state = 2
                else:
                    state = 0

            log((t, len(q0), len(qh), len(qs), state, ev_type[cur_s] if cur_s >= 0 else -1))

        J = np.asarray(journal, dtype=float).reshape(-1, 6)
        return self._statistics(sc, types, is_hrt, ref_seq, type_seq, ev_type, arr, sp, ep, ss, es,
                                J[:, 0], J[:, 1], J[:, 2], J[:, 3], J[:, 4], J[:, 5].astype(int))

    # ------------------------------------------------------------------
    def _statistics(self, sc, types, is_hrt, ref_seq, type_seq, ev_type, arr, sp, ep, ss, es,
                    ev_t, ev_q0, ev_qh, ev_qs, ev_state, ev_sec_type) -> Dict:
        K = len(types)
        n_ref = len(ref_seq)
        i_start = int(np.ceil(n_ref * sc.warmup_fraction))
        i_end = int(np.floor(n_ref * (1 - sc.warmup_fraction)))
        empty = {'classic': {m: 0.0 for m in range(1, 19)}, 'per_type': {}, 'U': sc.utilization()}
        if n_ref == 0 or i_end <= i_start or i_start < 1:
            return empty
        t_start = arr[ref_seq[i_start - 1]]
        t_end = arr[ref_seq[i_end - 1]]
        if t_end <= t_start:
            return empty

        A = lambda x: np.asarray(x, dtype=float)
        arr_a, sp_a, ep_a, ss_a, es_a = A(arr), A(sp), A(ep), A(ss), A(es)
        typ = np.asarray(ev_type, dtype=int)
        deadline = np.array([et.deadline_ms for et in types])[typ]

        # Відбір подій стаціонарного інтервалу:
        #  МРЧ – опорні події з номерами i_start..i_end; ЖРЧ – номер у своєму типі 100..count-100
        #  і момент надходження в [t_start, t_end] (як у версії 1.x)
        selected = np.zeros(len(arr), dtype=bool)
        ref_ids = np.asarray(ref_seq[i_start - 1:i_end], dtype=int)
        srt_ref = ref_ids[~np.asarray(is_hrt, dtype=bool)[typ[ref_ids]]] if ref_ids.size else ref_ids
        selected[srt_ref] = True
        for k in range(K):
            if is_hrt[k]:
                ids = np.asarray(type_seq[k][99:max(99, len(type_seq[k]) - 100)], dtype=int)
                if ids.size:
                    a = arr_a[ids]
                    selected[ids[(a >= t_start) & (a <= t_end)]] = True
        if not any(not h for h in is_hrt):     # лише ЖРЧ: опорні – усі події
            selected[ref_ids] = True
        done = selected & (es_a > 0)

        wp = sp_a - arr_a; ws = ss_a - ep_a; soj = es_a - arr_a; late = soj > deadline
        is_hrt_arr = np.asarray(list(is_hrt) + [False], dtype=bool)   # індекс -1 → «немає задачі»
        hrt_mask = is_hrt_arr[typ]

        def agg(mask):
            m = done & mask
            if not m.any():
                return None
            return dict(wp_mean=float(wp[m].mean()), ws_mean=float(ws[m].mean()), soj_mean=float(soj[m].mean()),
                        wp_max=max(0.0, float(wp[m].max())), ws_max=max(0.0, float(ws[m].max())),
                        soj_max=max(0.0, float(soj[m].max())), p_miss=float(late[m].mean()))

        srt = agg(~hrt_mask); hrt = agg(hrt_mask)
        c = {}
        c[1] = srt['wp_mean'] if srt else 0.0
        c[2] = srt['ws_mean'] if srt else 0.0
        c[3] = hrt['ws_mean'] if hrt else 0.0
        c[4] = srt['soj_mean'] if srt else 0.0
        c[5] = hrt['soj_mean'] if hrt else 0.0
        c[6] = srt['wp_max'] if srt else 0.0
        c[7] = srt['ws_max'] if srt else 0.0
        c[8] = hrt['ws_max'] if hrt else 0.0
        c[9] = srt['soj_max'] if srt else 0.0
        c[10] = hrt['soj_max'] if hrt else 0.0
        c[17] = srt['p_miss'] if srt else 0.0
        c[18] = hrt['p_miss'] if hrt else 0.0

        # ---- Середні за часом (2.23) ----
        T = t_end - t_start
        et_ = A(ev_t)
        dt = np.clip(np.minimum(et_[1:], t_end) - np.maximum(et_[:-1], t_start), 0.0, None)
        st = np.asarray(ev_state, dtype=np.int8)[:-1]
        sec = np.asarray(ev_sec_type, dtype=int)[:-1]
        sec_hrt = is_hrt_arr[sec]
        f = lambda x: float(np.dot(x, dt) / T)
        c[11] = f(A(ev_q0)[:-1]); c[12] = f(A(ev_qs)[:-1]); c[13] = f(A(ev_qh)[:-1])
        c[14] = f((st == 1).astype(float))
        c[15] = f(((st == 2) & ~sec_hrt).astype(float))
        c[16] = f(((st == 2) & sec_hrt).astype(float))

        per_type = {}
        for k, et in enumerate(types):
            m = done & (typ == k)
            n = int(m.sum())
            if n == 0:
                per_type[et.name] = {key: 0.0 for key in METRICS_PER_TYPE}
                per_type[et.name]['count'] = 0
                continue
            so = soj[m]
            per_type[et.name] = {
                'count': n,
                'wait_primary_mean': float(wp[m].mean()),
                'wait_secondary_mean': float(ws[m].mean()),
                'sojourn_mean': float(so.mean()),
                'sojourn_max': float(so.max()),
                'sojourn_p95': float(np.percentile(so, 95)),
                'sojourn_p99': float(np.percentile(so, 99)),
                'p_miss': float(late[m].mean()),
                'utilization_secondary': f(((st == 2) & (sec == k)).astype(float)),
            }
        return {'classic': c, 'per_type': per_type, 'U': sc.utilization()}
