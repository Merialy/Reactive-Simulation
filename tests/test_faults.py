"""
Тести сценаріїв відмов (версія 3.2): шторм переривань і перевищення WCET.
"""
import copy
import os
import pytest

from simulation.engine import EventSimulation
from simulation.model import Scenario, Fault, HRT
from storage.repository import scenario_key

SCEN = os.path.join(os.path.dirname(__file__), "..", "scenarios")


@pytest.fixture(scope="module")
def base():
    sc = Scenario.load(os.path.join(SCEN, "boryviter_adcs.json"))
    sc.N = 6000
    return sc


def _with(sc, *faults):
    s = copy.deepcopy(sc)
    s.faults = list(faults)
    return s


def test_inactive_fault_changes_nothing(base):
    """Відмова, що не настає, не змінює результат (окремий генератор випадкових чисел)."""
    ref = EventSimulation(seed=5).run(base)
    far = _with(base, Fault(kind="storm", start_ms=1e12, duration_ms=1, rate=5),
                Fault(kind="overrun", start_ms=1e12, duration_ms=1, factor=3))
    r = EventSimulation(seed=5).run(far)
    assert r["classic"] == ref["classic"]
    assert r["faults"] == {"storm_events": 0.0, "overrun_events": 0.0}


def test_overrun_factor_one_changes_nothing(base):
    ref = EventSimulation(seed=5).run(base)
    r = EventSimulation(seed=5).run(_with(base, Fault(kind="overrun", factor=1.0, probability=1.0)))
    assert r["classic"] == ref["classic"]


def test_storm_hurts_srt_but_not_hrt(base):
    sc = _with(base, Fault(kind="storm", start_ms=5000, duration_ms=3000, rate=20, wcet_ms=0.03))
    r = EventSimulation(seed=1).run(sc)
    # кількість хибних переривань ≈ rate · duration
    assert r["faults"]["storm_events"] == pytest.approx(20 * 3000, rel=0.05)
    for et in sc.event_types:
        m = r["per_type"][et.name]
        assert m["count_fault"] > 0
        if et.cls == HRT:
            assert m["p_miss_fault"] == 0.0           # ЖРЧ захищені пріоритетом QH
    assert r["per_type"]["Команди НКК (UART)"]["p_miss_fault"] > 0.5


def test_overrun_probability_matches(base):
    """Подія датчиків з перевищенням ×2,5 (1,25 мс > D = 1 мс) завжди запізнюється – P_miss ≈ p."""
    sc = _with(base, Fault(kind="overrun", target="ADCS: датчики", probability=0.3, factor=2.5))
    r = EventSimulation(seed=2).run(sc)
    m = r["per_type"]["ADCS: датчики"]
    assert m["p_miss_fault"] == pytest.approx(0.3, abs=0.03)
    assert r["per_type"]["ADCS: закон керування"]["p_miss_fault"] == 0.0


def test_fault_json_and_check(base, tmp_path):
    sc = _with(base, Fault(kind="storm", start_ms=100, duration_ms=50, rate=2))
    path = tmp_path / "f.json"
    sc.save(str(path))
    assert Scenario.load(str(path)).to_dict() == sc.to_dict()
    # сценарій без відмов має той самий ключ у базі, що й у версії 3.1
    assert "faults" not in base.to_dict()
    assert scenario_key(base) != scenario_key(sc)
    bad = _with(base, Fault(kind="overrun", target="немає такого", factor=0.5))
    assert len(bad.check()[0]) == 2


def test_example_fault_scenarios_load():
    for name in ("boryviter_fault_storm.json", "boryviter_fault_wcet.json"):
        sc = Scenario.load(os.path.join(SCEN, name))
        assert sc.faults and not sc.check()[0]
