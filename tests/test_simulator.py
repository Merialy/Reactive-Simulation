"""
Автоматизовані тести верифікації імітаційної моделі (розділ 2, п. 2.8.2).

Запуск з кореня проєкту:  python -m pytest -q
"""
import pytest

from simulation.priority_simulator import PriorityQueueSimulation
from simulation.replications import run_replications, summarize
from simulation.validation import check_params, utilization

BASE = dict(lambda1=1.0, s1=0.002, s2=0.5, N=30000, D1=1.0, lambda2=2.0, s2b=0.1, D2=0.2)


@pytest.fixture(scope="module")
def base_result():
    return run_replications(BASE, replications=3, base_seed=1)['mean']


# ---------- Тест 1: граничний випадок M/D/1 (формула 2.8) ----------
@pytest.mark.parametrize("rho", [0.3, 0.5, 0.7])
def test_md1_limit_case(rho):
    p = dict(lambda1=1.0, s1=1e-9, s2=rho, N=50000, D1=10.0, lambda2=0.0, s2b=0.1, D2=1.0)
    w_theory = rho * rho / (2 * (1 - rho))          # λ s² / (2(1-ρ)), λ = 1
    w_sim = run_replications(p, replications=3, base_seed=7)['mean'][2]
    assert abs(w_sim / w_theory - 1) < 0.05          # критерій К1: ≤ 5 %


# ---------- Тест 2: закон збереження роботи ----------
def test_work_conservation(base_result):
    r = base_result
    assert r[15] == pytest.approx(BASE['lambda1'] * BASE['s2'], rel=0.03)
    assert r[16] == pytest.approx(BASE['lambda2'] * BASE['s2b'], rel=0.01)
    assert r[14] == pytest.approx((BASE['lambda1'] + BASE['lambda2']) * BASE['s1'], rel=0.05)


# ---------- Тест 3: гарантія обробки подій ЖРЧ (умови 2.18–2.19) ----------
def test_hrt_deadline_guarantee(base_result):
    r = base_result
    assert r[18] == 0.0                              # жодна подія ЖРЧ не порушила дедлайн
    assert r[10] <= BASE['D2']                       # максимальний час перебування ЖРЧ


# ---------- Тест 4: закон Літтла (формула 2.27) ----------
def test_little_law_srt(base_result):
    r = base_result
    # кількість подій МРЧ на вторинному етапі: у черзі Q_S + в обслуговуванні
    L = r[12] + r[15]
    # час від завершення ISR до завершення реакції
    W = r[4] - r[1] - BASE['s1']
    assert L == pytest.approx(BASE['lambda1'] * W, rel=0.03)


def test_little_law_hrt(base_result):
    r = base_result
    L = r[13] + r[16]
    W = r[5] - BASE['s1']
    assert L == pytest.approx(BASE['lambda2'] * W, rel=0.03)


# ---------- Відтворюваність (критерій К4) ----------
def test_reproducibility():
    p = dict(BASE, N=5000)
    a = PriorityQueueSimulation(seed=123).run_simulation_priority2_full(**p)
    b = PriorityQueueSimulation(seed=123).run_simulation_priority2_full(**p)
    c = PriorityQueueSimulation(seed=124).run_simulation_priority2_full(**p)
    assert a == b
    assert a != c


# ---------- Граничні значення параметрів ----------
def test_no_hrt_flow_does_not_crash():
    p = dict(BASE, lambda2=0.0, N=5000)
    r = PriorityQueueSimulation(seed=1).run_simulation_priority2_full(**p)
    assert r[16] == 0.0 and r[18] == 0.0


def test_invalid_lambda1_raises():
    with pytest.raises(ValueError):
        PriorityQueueSimulation(seed=1).run_simulation_priority2_full(**dict(BASE, lambda1=0.0))


# ---------- Перевірка параметрів (таблиця 2.3) ----------
def test_validation_warnings():
    assert utilization(BASE) == pytest.approx(0.706)
    errors, warnings = check_params(dict(BASE, s2=0.9))      # U > 1
    assert not errors and any('U =' in w for w in warnings)
    errors, warnings = check_params(dict(BASE, D1=0.3))      # D1 < s1 + s2
    assert any('D1' in w for w in warnings)
    errors, _ = check_params(dict(BASE, lambda1=0.0))
    assert errors


def test_confidence_interval():
    s = summarize([1.0, 2.0, 3.0, 4.0, 5.0])
    assert s['mean'] == 3.0
    assert s['ci'] == pytest.approx(2.776 * (2.5 ** 0.5) / 5 ** 0.5, rel=1e-3)


# ---------- Паралельний запуск дає ті самі результати, що й послідовний ----------
def test_parallel_batch_is_deterministic():
    from simulation.replications import run_batch
    p = dict(BASE, N=3000)
    plist = [p, None, dict(p, s2=0.3)]
    par = run_batch(plist, replications=3, base_seed=5, workers=2)
    seq = run_batch(plist, replications=3, base_seed=5, workers=1)
    assert par[1] is None
    assert par[0]['mean'] == seq[0]['mean'] and par[2]['mean'] == seq[2]['mean']
    assert par[0]['mean'] == run_replications(p, replications=3, base_seed=5)['mean']


def test_t_quantiles_and_settings():
    """Квантилі Стьюдента (довідкові значення) і перевірка R та seed (версія 4.7)."""
    from simulation.replications import t_quantile_975, check_settings
    for df, q in ((1, 12.706), (4, 2.776), (22, 2.074), (40, 2.021), (60, 2.000), (150, 1.976), (199, 1.972)):
        assert t_quantile_975(df) == pytest.approx(q, abs=1e-3)
    check_settings(1, 0)
    with pytest.raises(ValueError):
        check_settings(0, 42)
    with pytest.raises(ValueError):
        check_settings(5, -1)
