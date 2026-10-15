"""
Пояснення простими словами: що відбувається в сценарії (перед запуском) і що означають
результати (після запуску). Не залежить від Qt – повертає HTML-рядки для QLabel.

    scenario_facts(sc)               → коротко про сценарій: 4 плитки, попередження, підказка
    scenario_conclusion(sc, res, P)  → висновок за результатами сценарію з типами подій
    batch_conclusion(batch, warns, errs, P) → висновок за серією з таблиці параметрів
"""
from html import escape

from simulation.model import HRT, STORM
from ui.core import fmt
from ui.i18n import is_en, tr

_DEF = {"hrt": "#C2410C", "srt": "#2563EB", "ok": "#15803D", "warn": "#B45309", "err": "#B91C1C", "muted": "#64748B"}


def _c(P, key):
    return (P or {}).get(key) or _DEF[key]


def pct(p: float) -> str:
    """Ймовірність у відсотках: 0.0237 → «2,4 %», 0.0004 → «< 0,1 %»."""
    if p is None:
        return "—"
    if p <= 0:
        return "0 %"
    if p < 0.001:
        return tr("< {v} %", v=fmt(0.1))
    return f"{fmt(p * 100, 2 if p < 0.1 else 3)} %"


def _q(name: str) -> str:
    return ("“{}”" if is_en() else "«{}»").format(escape(str(name)))


def model_time_ms(sc) -> float:
    """Приблизна тривалість модельного часу: N «опорних» подій (МРЧ, а якщо їх немає – усіх)."""
    srt = [et.rate for et in sc.event_types if et.cls != HRT]
    lam = sum(srt) if srt else sum(et.rate for et in sc.event_types)
    return sc.N / lam if lam > 0 else 0.0


def load_words(u: float) -> str:
    if u >= 1:
        return tr("процесор перевантажений – він не встигає обробляти події, черги ростуть без кінця")
    if u > 0.95:
        return tr("процесор майже повністю зайнятий – черги можуть бути дуже довгими")
    if u > 0.7:
        return tr("запас продуктивності невеликий")
    return tr("запас продуктивності достатній")


def _load_sentence(u: float, P=None) -> str:
    lvl = "err" if u >= 1 else ("warn" if u > 0.95 else "ok")
    if u >= 1:
        return tr("Процесору потрібно <b style='color:{c}'>{u}</b> його потужності – більше, ніж він має: "
                  "{w}.", c=_c(P, lvl), u=pct(u), w=load_words(u))
    return tr("Процесор зайнятий у середньому на <b style='color:{c}'>{u}</b> часу – {w}.",
              c=_c(P, lvl), u=pct(u), w=load_words(u))


def _soft_words(p: float) -> str:
    if p < 0.05:
        return tr("Частка мала – для м'яких подій це зазвичай допустимо.")
    if p < 0.2:
        return tr("Частка помітна – варто оцінити, чи це прийнятно.")
    return tr("Запізнюється значна частина подій – бракує продуктивності процесора або дедлайн надто жорсткий.")


def _sec(ms: float) -> str:
    return fmt(ms / 1000.0, 3)


def _fault_short(f) -> str:
    if f.kind == STORM:
        return tr("шторм")
    return tr("повільна обробка {t}", t=_q(f.target) if f.target else tr("усіх типів"))


def scenario_facts(sc) -> dict:
    """
    Коротко про сценарій для сторінки «Типи подій і відмови»:
        facts  – 4 плитки (значення, підпис, рівень ok/warn/err/"")
        alerts – список (рівень, текст): помилки, попередження або «✔ можна запускати»
        hint   – один рядок: на що дивитися після запуску
    Деталі (які саме типи, як часто, параметри відмови) – у таблицях нижче, тут не повторюються.
    """
    from ui.i18n import tr_msg
    hrt = sum(1 for et in sc.event_types if et.cls == HRT)
    srt = len(sc.event_types) - hrt
    facts = [(str(len(sc.event_types)), tr("типів подій: {h} ЖРЧ, {s} МРЧ", h=hrt, s=srt), "")]
    errors, warns = sc.check()
    alerts = [("err", tr_msg(e)) for e in errors]
    if errors:
        return {"facts": facts, "alerts": alerts, "hint": ""}

    u = sc.utilization()
    lvl = "err" if u >= 1 else ("warn" if u > 0.95 else "ok")
    short = (tr("перевантаження") if u >= 1 else tr("майже на межі") if u > 0.95 else
             tr("запас невеликий") if u > 0.7 else tr("запас є"))
    facts.append((pct(u), tr("завантаження процесора – {w}", w=short), lvl))

    active = sc.active_faults()
    T = model_time_ms(sc)
    if active:
        uf = sc.utilization_during_faults()
        lvlf = "err" if uf >= 1 else ("warn" if uf > 0.95 else "ok")
        a = min(f.start_ms for f in active)
        ends = [f.end_ms() for f in active if f.duration_ms is not None]
        span = (f"{_sec(a)}–{_sec(max(ends))} " + tr("с") if len(ends) == len(active) else tr("від {a} с", a=_sec(a)))
        what = _fault_short(active[0]) if len(active) == 1 else tr("{n} відмови", n=len(active))
        facts.append((pct(uf), tr("під час відмови: {w}, {t}", w=what, t=span), lvlf))
    else:
        facts.append((tr("немає"), tr("відмов – нормальна робота"), ""))
    facts.append((f"≈ {_sec(T)} " + tr("с"), tr("модельного часу ({n} подій МРЧ)", n=f"{sc.N:,}".replace(",", " ")), ""))

    for f in active:
        if f.start_ms >= T > 0:
            alerts.append(("err", tr("Відмова {f} почнеться вже після кінця моделювання (≈ {t} с) – збільште N.",
                                     f=_q(f.label()), t=_sec(T))))
        elif f.duration_ms is not None and f.end_ms() > T > 0:
            alerts.append(("warn", tr("Моделювання закінчиться раніше, ніж відмова {f} (≈ {t} с).",
                                      f=_q(f.label()), t=_sec(T))))
    for w in warns:
        if w.startswith("Під час відмов завантаження"):      # те саме, що й плитка «під час відмови», – простіше
            w_txt = tr("Під час відмови процесор не встигає – черги ростимуть, поки вона діє.")
        else:
            w_txt = tr_msg(w)
        alerts.append(("warn", w_txt))
    if not alerts:
        alerts = [("ok", tr("Сценарій коректний, можна запускати."))]
    hint = (tr("Після запуску порівняйте таблицю «Під час дії відмов» з основною. Щоб побачити вплив відмови, "
               "запустіть сценарій ще раз без неї (зніміть галочку) і порівняйте обидва запуски в «Історії».")
            if active else tr("Після запуску головне – «P порушення дедлайну»: для ЖРЧ вона має бути 0."))
    return {"facts": facts, "alerts": alerts, "hint": hint}


def _get(res, name, key):
    v = (res.get("per_type", {}).get(name) or {}).get(key)
    return None if v is None else v.get("mean")


def scenario_conclusion(sc, res, P=None) -> str:
    """Висновок простими словами за результатами сценарію з типами подій."""
    out = []
    ok, warn, err = (_c(P, k) for k in ("ok", "warn", "err"))
    hrt = [et for et in sc.event_types if et.cls == HRT]
    srt = [et for et in sc.event_types if et.cls != HRT]
    u = sc.utilization()
    out.append(_load_sentence(u, P))

    def misses(types, key):
        return [(et, _get(res, et.name, key)) for et in types if (_get(res, et.name, key) or 0) > 0]

    if hrt:
        m = misses(hrt, "p_miss")
        if not m:
            worst = max(hrt, key=lambda et: (_get(res, et.name, "sojourn_max") or 0) / et.deadline_ms)
            out.append(f"<span style='color:{ok}'>✔ " + tr("Усі жорсткі події (ЖРЧ) обробляються вчасно.") + "</span> " +
                       tr("Найближче до межі – {n}: найдовша реакція {t} мс при дедлайні {d} мс.",
                          n=_q(worst.name), t=fmt(_get(res, worst.name, "sojourn_max"), 3), d=fmt(worst.deadline_ms)))
        else:
            out.append(f"<span style='color:{err}'>✖ " + tr("Жорсткі дедлайни порушуються:") + " " +
                       "; ".join(tr("{n} – у {p} випадків", n=_q(et.name), p=pct(p)) for et, p in m) +
                       ". " + tr("Для бортової системи це неприпустимо.") + "</span>")
    if srt:
        m = misses(srt, "p_miss")
        if not m:
            out.append(f"<span style='color:{ok}'>✔ " + tr("М'які події (МРЧ) теж укладаються в дедлайни.") + "</span>")
        else:
            out.append(f"<span style='color:{warn}'>⚠ " + tr("М'які події (МРЧ) запізнюються:") + " " +
                       "; ".join(tr("{n} – у {p} випадків", n=_q(et.name), p=pct(p)) for et, p in m) +
                       ".</span> " + _soft_words(max(p for _, p in m)))

    for et in sc.event_types:
        wh = et.weakly_hard()
        if not wh:
            continue
        m, n = wh
        viol = _get(res, et.name, "wh_viol") or 0.0
        runs = [r.get("per_type", {}).get(et.name, {}).get("wh_worst_miss") for r in res.get("runs") or []]
        runs = [x for x in runs if x is not None]
        worst = max(runs) if runs else (_get(res, et.name, "wh_worst_miss") or 0.0)   # найгірше за всі прогони
        rule = tr("«{m} з {n}»", m=m, n=n)
        if viol <= 0:
            out.append(f"<span style='color:{ok}'>✔ " +
                       tr("{t}: обмеження {r} виконується – у будь-яких {n} подіях поспіль пропусків не більше {k}.",
                          t=_q(et.name), r=rule, n=n, k=n - m) + "</span>")
        else:
            c = err if et.cls == HRT else warn
            out.append(f"<span style='color:{c}'>✖ " +
                       tr("{t}: обмеження {r} порушується у {p} вікон (у найгіршому вікні {w} пропусків з {n}, "
                          "допустимо {k}).", t=_q(et.name), r=rule, p=pct(viol), w=fmt(worst, 3), n=n, k=n - m) +
                       "</span>")

    if sc.active_faults():
        hit_h = misses(hrt, "p_miss_fault")
        hit_s = misses(srt, "p_miss_fault")
        names = "; ".join(_q(f.label()) for f in sc.active_faults())
        n_fault = sum(_get(res, et.name, "count_fault") or 0 for et in sc.event_types)
        if n_fault < 1:
            out.append(f"<span style='color:{warn}'>⚠ " +
                       tr("Відмова ({f}) не встигла початися – моделювання закінчилося раніше. Збільште N, "
                          "щоб модельний час охопив інтервал відмови.", f=names) + "</span>")
        elif not hit_h and not hit_s:
            out.append(f"<span style='color:{ok}'>✔ " +
                       tr("Відмова ({f}) не призвела до порушень дедлайнів – система її витримала.", f=names) + "</span>")
        else:
            parts = []
            for et, p in hit_h + hit_s:
                parts.append(tr("{n} – {p} (за весь час {q})", n=_q(et.name), p=pct(p), q=pct(_get(res, et.name, "p_miss"))))
            out.append(tr("<b>Під час відмови</b> ({f}) запізнюються:", f=names) + " " + "; ".join(parts) + ".")
            if hit_s and not hit_h and hrt:
                out.append(tr("Жорсткі події не постраждали – їх захищає вищий пріоритет; весь удар прийняли "
                              "м'які події."))
            elif hit_h:
                out.append(tr("Відмова зачепила жорсткі події – пріоритет від цього не захищає, бо збій стався "
                              "в обробці самої жорсткої задачі або процесор перевантажений."))
    return "<br>".join(out)


def _ranges(xs) -> str:
    """[1, 2, 3, 5, 7, 8] → «1–3, 5, 7, 8»."""
    xs = sorted(xs)
    parts, i = [], 0
    while i < len(xs):
        j = i
        while j + 1 < len(xs) and xs[j + 1] == xs[j] + 1:
            j += 1
        parts.append(f"{xs[i]}–{xs[j]}" if j - i >= 2 else ", ".join(str(x) for x in xs[i:j + 1]))
        i = j + 1
    return ", ".join(parts)


def batch_conclusion(batch, warns, errs, P=None) -> str:
    """Висновок за серією сценаріїв (стовпці таблиці параметрів, 18 показників)."""
    ok, warn, err = (_c(P, k) for k in ("ok", "warn", "err"))
    done = [c for c, b in enumerate(batch) if b is not None]
    out = [tr("Змодельовано сценаріїв: {n} із {m}.", n=len(done), m=len(batch))]
    hard = [c + 1 for c in done if batch[c]["mean"][18] > 0]
    soft = [c + 1 for c in done if batch[c]["mean"][17] > 0]
    nums = _ranges

    def share(m, idx):
        ps = [batch[c - 1]["mean"][idx] for c in m]
        lo, hi = min(ps), max(ps)
        return tr(" (частка запізнень {p})", p=pct(hi)) if lo == hi else \
            tr(" (частка запізнень від {a} до {b})", a=pct(lo), b=pct(hi))

    def where(m):
        if len(m) == len(done) and len(m) > 1:
            return tr("у всіх змодельованих сценаріях") if errs else tr("у всіх сценаріях")
        return tr("у сценаріях {s}", s=nums(m))
    if done:
        if hard:
            out.append(f"<span style='color:{err}'>✖ " + tr("Жорсткі дедлайни (ЖРЧ) порушуються {w}", w=where(hard)) +
                       share(hard, 18) + ".</span>")
        else:
            out.append(f"<span style='color:{ok}'>✔ " + tr("У жодному сценарії жорсткі дедлайни (ЖРЧ) не порушуються.") + "</span>")
        if soft:
            out.append(f"<span style='color:{warn}'>⚠ " + tr("М'які дедлайни (МРЧ) порушуються {w}", w=where(soft)) +
                       share(soft, 17) + ".</span>")
        else:
            out.append(f"<span style='color:{ok}'>✔ " + tr("М'які дедлайни (МРЧ) теж виконуються всюди.") + "</span>")
        if len(done) > 1:
            worst = max(done, key=lambda c: batch[c]["mean"][4])
            out.append(tr("Найповільніша реакція – у сценарії {c}: у середньому {a} мс для МРЧ і {b} мс для ЖРЧ.",
                          c=worst + 1, a=fmt(batch[worst]["mean"][4], 3), b=fmt(batch[worst]["mean"][5], 3)))
    if errs:
        out.append(f"<span style='color:{err}'>" + tr("Не змодельовано через помилки в даних: {s}.",
                                                       s=nums(c + 1 for c in sorted(errs))) + "</span>")
    if warns:
        out.append(tr("З попередженнями (⚠ у заголовку – наведіть мишу): {s}.", s=nums(c + 1 for c in sorted(warns))))
    out.append(tr("Щоб побачити, як показники змінюються від параметра, відкрийте «Графіки»."))
    return "<br>".join(out)
