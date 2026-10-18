"""
Тести фонових задач (версія 4.6): черга QBG з найнижчим пріоритетом, перемикання контексту t_cs,
критичні секції з вимкненими перериваннями (час блокування).
"""
import copy
import os

import numpy as np
import pytest

from simulation.engine import EventSimulation
from simulation.model import Background, EventType, Platform, Scenario, SRT

SCEN = os.path.join(os.path.dirname(__file__), "..", "scenarios")


@pytest.fixture(scope="module")
def boryviter():
    sc = Scenario.load(os.path.join(SCEN, "boryviter_adcs.json"))
    sc.N = 6000
    return sc


@pytest.mark.parametrize("mode", ["periodic", "poisson", "continuous"])
def test_free_background_changes_nothing(boryviter, mode):
    """Без перемикання контексту і критичних секцій фонова задача не впливає на ЖРЧ і МРЧ."""
    base = EventSimulation(5).run(boryviter)
    sc = copy.deepcopy(boryviter)
    sc.background = Background(mode=mode, interval_ms=50, work_ms=10)
    r = EventSimulation(5).run(sc)
    for k in range(1, 19):
        assert r['classic'][k] == pytest.approx(base['classic'][k], rel=1e-9, abs=1e-12)
    for name, t in base['per_type'].items():
        assert r['per_type'][name]['sojourn_max'] == pytest.approx(t['sojourn_max'], rel=1e-9)


def test_continuous_background_takes_all_idle_time(boryviter):
    sc = copy.deepcopy(boryviter)
    sc.background = Background(mode="continuous")
    r = EventSimulation(1).run(sc)
    assert r['system']['bg_util'] == pytest.approx(1 - sc.utilization(), abs=0.02)


def test_background_job_time_matches_md1():
    """Майже без подій реального часу фонова робота – це M/D/1: T = s + λs²/(2(1 − ρ))."""
    sc = Scenario(name="t", platform=Platform(name="x", t_irq_ms=1e-6),
                  event_types=[EventType(name="A", cls=SRT, rate=0.001, deadline_ms=10, wcet_ms=1e-6)], N=400,
                  background=Background(mode="poisson", interval_ms=10, work_ms=5))
    sims = [EventSimulation(sd).run(sc)['system']['bg_resp_mean'] for sd in range(10)]
    assert np.mean(sims) == pytest.approx(5 + 0.1 * 25 / (2 * 0.5), rel=0.05)


def test_critical_sections_block_but_bounded(boryviter):
    sc = copy.deepcopy(boryviter)
    sc.background = Background(mode="periodic", interval_ms=100, work_ms=20, t_cs_ms=0.005,
                               crit_ms=0.2, crit_share=0.1)
    base = EventSimulation(5).run(boryviter)
    r = EventSimulation(5).run(sc)
    sens = r['per_type']["ADCS: датчики"]
    assert 0 < sens['blocking_max'] <= 0.2 + 1e-12          # не довше за одну критичну секцію
    assert r['system']['tid_max'] <= 0.2 + 1e-12
    # найгірший час реакції зростає, але не більше ніж на t_id + t_cs (+ ISR, що стали в чергу за час блокування)
    grow = sens['sojourn_max'] - base['per_type']["ADCS: датчики"]['sojourn_max']
    assert 0 < grow <= 0.2 + 0.005 + 0.01
    assert r['system']['cs_per_ms'] > 0 and r['system']['cs_util'] < 0.01


def test_background_json_and_check():
    sc = Scenario.load(os.path.join(SCEN, "base_two_types.json"))
    assert 'background' not in sc.to_dict()                  # старий формат і ключ БД без змін
    sc.background = Background(mode="none")
    assert 'background' not in sc.to_dict()
    sc.background = Background(mode="periodic", interval_ms=10, work_ms=5, crit_share=1.5)
    assert sc.check()[0]
    sc.background.crit_share = 0.1
    assert not sc.check()[0]
    sc.background.work_ms = 9                                # U + U_BG ≥ 1 → попередження
    assert any("U + U_BG" in w for w in sc.check()[1])
    again = Scenario.from_dict(sc.to_dict())
    assert again.background == sc.background
