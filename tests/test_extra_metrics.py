"""
Тести версії 4.5: обмеження (m, n), джиттер, пропуски поспіль, готовність, максимальні черги,
аналітичні наближення п. 2.5.2 і системні показники п. 1.3.
"""
import os

import numpy as np
import pytest

from simulation import extra_metrics as xm
from simulation.analytic import approx1, approx2, approx3, system_load
from simulation.model import Scenario
from simulation.replications import run_scenario_replications

SCEN = os.path.join(os.path.dirname(__file__), "..", "scenarios")


def test_max_consecutive_and_windows():
    late = np.array([0, 1, 1, 0, 1, 1, 1, 0, 0, 1], dtype=bool)
    assert xm.max_consecutive(late) == 3
    assert xm.max_consecutive(np.zeros(5, dtype=bool)) == 0
    assert list(xm.window_misses(late, 3)) == [2, 2, 2, 2, 3, 2, 1, 1]


def test_weakly_hard_rule():
    late = np.array([0, 1, 1, 0, 1, 1, 1, 0, 0, 1], dtype=bool)
    # «1 з 3»: у вікні з 3 подій допустимо 2 пропуски → порушене лише одне вікно (3 пропуски)
    r = xm.weakly_hard(late, 1, 3)
    assert r['wh_worst_miss'] == 3 and r['wh_ok'] == 0.0
    assert r['wh_viol'] == pytest.approx(1 / 8)
    # без пропусків навіть найсуворіше обмеження «3 з 3» виконується
    assert xm.weakly_hard(np.zeros(10, dtype=bool), 3, 3)['wh_ok'] == 1.0


def test_readiness_and_queue_max():
    assert xm.readiness(np.array([0.0, 0.0, 0.1, 0.0])) == 0.75
    t = np.array([0.0, 1.0, 2.0, 3.0, 4.0])
    q = np.array([0.0, 5.0, 1.0, 2.0, 9.0])
    assert xm.queue_max(t, q, 1.5, 3.5) == 5.0       # стан на момент 1,5 (після t = 1) теж враховується
    assert xm.queue_max(t, q, 0.0, 4.0) == 9.0


def test_scenario_weakly_hard_normal_vs_fault():
    ok = Scenario.load(os.path.join(SCEN, "boryviter_adcs.json"))
    bad = Scenario.load(os.path.join(SCEN, "boryviter_fault_wcet.json"))
    for sc in (ok, bad):
        sc.N = 12000
    r_ok = run_scenario_replications(ok, 2, 1, workers=1)['per_type']["ADCS: датчики"]
    r_bad = run_scenario_replications(bad, 2, 1, workers=1)['per_type']["ADCS: датчики"]
    assert r_ok['wh_ok']['mean'] == 1.0 and r_ok['wh_viol']['mean'] == 0.0
    assert r_bad['wh_viol']['mean'] > 0 and r_bad['wh_ok']['mean'] == 0.0
    # джиттер і готовність мають сенс
    assert r_ok['jitter']['mean'] >= 0 and 0 <= r_ok['readiness']['mean'] <= 1


def test_weakly_hard_validation_and_json():
    sc = Scenario.load(os.path.join(SCEN, "base_two_types.json"))
    assert all('wh_m' not in e for e in sc.to_dict()['event_types'])   # старий формат і ключ БД без змін
    sc.event_types[0].wh_m = 3
    assert any("обидва" in e for e in sc.check()[0])
    sc.event_types[0].wh_n = 2
    assert any("1 ≤ m ≤ n" in e for e in sc.check()[0])
    sc.event_types[0].wh_n = 5
    assert not sc.check()[0]
    again = Scenario.from_dict(sc.to_dict())
    assert again.event_types[0].weakly_hard() == (3, 5)


def test_approximations():
    md1 = dict(lambda1=1.4, s1=1e-9, s2=0.5, lambda2=0.0, s2b=0.1)
    exact = 1.4 * 0.25 / (2 * (1 - 0.7)) + 0.5                       # M/D/1, Поллачек–Хінчин
    for f in (approx1, approx2, approx3):
        assert f(md1) == pytest.approx(exact, rel=1e-6)                # без ЖРЧ усі три збігаються
    p = dict(lambda1=1.0, s1=0.002, s2=0.5, lambda2=2.0, s2b=0.1)
    assert approx1(p) < approx2(p) < approx3(p)
    assert approx1(dict(p, s2=0.9)) is None                            # U ≥ 1 – стаціонарного режиму немає


def test_system_load():
    v = system_load(0.8, 0.1)
    assert v['H'] == pytest.approx(0.2) and v['eta'] == pytest.approx(0.875)


def test_parse_wh():
    pytest.importorskip("PySide6")
    from ui.page_scenario import parse_wh
    assert parse_wh("4 з 5") == (4, 5)
    assert parse_wh("4/5") == (4, 5)
    assert parse_wh("(4, 5)") == (4, 5)
    assert parse_wh("") is None
    with pytest.raises(ValueError):
        parse_wh("4")
