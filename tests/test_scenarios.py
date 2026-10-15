"""
Тести сценаріїв з кількома типами подій (версія 2.0).
"""
import os
import pytest

from simulation.engine import EventSimulation
from simulation.model import Scenario, Platform, EventType, HRT, SRT, cycles_to_ms
from simulation.priority_simulator import PriorityQueueSimulation
from simulation.replications import run_scenario_replications

HERE = os.path.dirname(__file__)
SCEN = os.path.join(HERE, "..", "scenarios")

LEGACY = [
    dict(lambda1=1.0, s1=0.01, s2=0.75, N=8000, D1=1.0, lambda2=2.0, s2b=0.1, D2=0.2),
    dict(lambda1=1.0, s1=0.002, s2=0.3, N=8000, D1=0.4, lambda2=2.0, s2b=0.1, D2=0.2),
    dict(lambda1=1.0, s1=1e-9, s2=0.5, N=8000, D1=1.0, lambda2=0.0, s2b=0.1, D2=0.2),
]


# ---------- Сумісність: нове ядро = ядро версії 1.x для двох типів ----------
@pytest.mark.parametrize("p", LEGACY)
def test_engine_matches_legacy_exactly(p):
    old = PriorityQueueSimulation(seed=3).run_simulation_priority2_full(**p)
    out = EventSimulation(seed=3).run(Scenario.from_legacy(p))
    new = out['classic']
    for k in range(1, 19):
        assert new[k] == pytest.approx(old[k], rel=1e-9, abs=1e-12)
    # додаткові показники (версія 4.5) обох ядер теж збігаються
    sc = Scenario.from_legacy(p)
    names = {et.cls: et.name for et in sc.event_types}
    pairs = [('srt', SRT), ('hrt', HRT)]
    for pre, cls in pairs:
        if f'{pre}_p95' not in old:
            continue
        t = out['per_type'][names[cls]]
        assert t['sojourn_p95'] == pytest.approx(old[f'{pre}_p95'], rel=1e-9, abs=1e-12)
        assert t['sojourn_p99'] == pytest.approx(old[f'{pre}_p99'], rel=1e-9, abs=1e-12)
        assert t['jitter'] == pytest.approx(old[f'{pre}_jitter'], rel=1e-9, abs=1e-12)
        assert t['max_consec_miss'] == old[f'{pre}_consec']
        assert t['readiness'] == pytest.approx(old[f'{pre}_ready'], abs=1e-12)
    for k in ('q0_max', 'qs_max', 'qh_max'):
        assert out['system'][k] == old[k]


# ---------- Формат JSON ----------
def test_json_roundtrip(tmp_path):
    sc = Scenario.load(os.path.join(SCEN, "boryviter_adcs.json"))
    path = tmp_path / "copy.json"
    sc.save(str(path))
    assert Scenario.load(str(path)).to_dict() == sc.to_dict()


# ---------- Апаратні параметри, формула (2.12) ----------
def test_cycles_conversion():
    assert cycles_to_ms(195, 100) == pytest.approx(0.00195)
    p = Platform(f_clk_mhz=100, c_irq=195, c_core=300)
    et = EventType(name="x", cls=SRT, rate=1, deadline_ms=1, wcet_cycles=50000)
    assert p.irq_ms() == pytest.approx(0.00195)
    assert et.service_ms(p) == pytest.approx(0.5 + 0.003)


# ---------- Кілька типів: збереження роботи і гарантія ЖРЧ ----------
@pytest.fixture(scope="module")
def boryviter():
    sc = Scenario.load(os.path.join(SCEN, "boryviter_adcs.json"))
    return sc, run_scenario_replications(sc, replications=3, base_seed=1, workers=1)


def test_multitype_work_conservation(boryviter):
    sc, res = boryviter
    for et in sc.event_types:
        expected = et.rate * et.service_ms(sc.platform)
        got = res['per_type'][et.name]['utilization_secondary']['mean']
        assert got == pytest.approx(expected, rel=0.05)


def test_multitype_hrt_guarantee(boryviter):
    sc, res = boryviter
    for et in sc.event_types:
        if et.cls == HRT:
            assert res['per_type'][et.name]['p_miss']['mean'] == 0.0


# ---------- Лише події ЖРЧ ----------
def test_only_hrt_scenario_runs():
    sc = Scenario(name="ЖРЧ", platform=Platform(t_irq_ms=0.002),
                  event_types=[EventType(name="A", cls=HRT, rate=1.0, wcet_ms=0.2, deadline_ms=0.5),
                               EventType(name="B", cls=HRT, rate=0.5, wcet_ms=0.3, deadline_ms=1.0, phase_ms=0.3)],
                  N=3000)
    r = EventSimulation(seed=1).run(sc)
    assert r['per_type']['A']['count'] > 0 and r['per_type']['B']['count'] > 0
    assert r['per_type']['A']['p_miss'] == 0.0


# ---------- Перевірка сценарію (таблиця 2.3) ----------
def test_scenario_check():
    sc = Scenario.from_legacy(dict(lambda1=1, s1=0.002, s2=0.9, N=5000, D1=0.5, lambda2=2, s2b=0.1, D2=0.2))
    errors, warnings = sc.check()
    assert not errors
    assert any("U =" in w for w in warnings)
    assert any("недосяжний" in w for w in warnings)
    bad = Scenario(name="x", platform=Platform(), event_types=[])
    assert bad.check()[0]
