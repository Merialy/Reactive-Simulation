"""
Тести підсистеми збереження (етап 3).

Тести JSON-сховища виконуються завжди. Тест MongoDB виконується лише тоді, коли встановлено
pymongo і запущено сервер (наприклад, `docker compose up -d`); інакше він пропускається.
"""
import os
import pytest

from simulation.model import Scenario
from simulation.replications import run_scenario_replications, run_replications
from storage.repository import JsonRepository, MongoRepository, result_document

SCEN = os.path.join(os.path.dirname(__file__), "..", "scenarios")


def _roundtrip(repo):
    sc = Scenario.load(os.path.join(SCEN, "boryviter_adcs.json"))
    sid = repo.save_scenario(sc)
    assert repo.get_scenario(sid).to_dict() == sc.to_dict()
    assert any(s["id"] == sid for s in repo.list_scenarios())

    res = run_scenario_replications(sc, replications=2, base_seed=1, workers=1)
    rid = repo.save_result(sc, res, {"replications": 2, "seed": 1}, scenario_id=sid, tags=["тест"])
    doc = repo.get_result(rid)
    assert doc["scenario_id"] == sid
    assert doc["settings"] == {"replications": 2, "seed": 1}
    assert len(doc["classic"]) == 18
    names = [t["name"] for t in doc["per_type"]]
    assert names == [et.name for et in sc.event_types]
    assert any(r["id"] == rid for r in repo.list_results(scenario_id=sid))

    assert repo.delete_result(rid) and repo.get_result(rid) is None
    assert repo.delete_scenario(sid) and repo.get_scenario(sid) is None


def test_json_repository(tmp_path):
    _roundtrip(JsonRepository(str(tmp_path)))


LEG = dict(lambda1=1.0, s1=0.01, s2=0.5, N=2000, D1=1.0, lambda2=2.0, s2b=0.1, D2=0.2)


def _no_duplicates(repo):
    rep = run_replications(LEG, replications=2, base_seed=1)
    st = {"replications": 2, "seed": 1, "source_file": "A.xlsx"}
    sc_a = Scenario.from_legacy(LEG, name="A.xlsx – сценарій 1")
    sid = repo.save_scenario(sc_a)
    rid1, new1 = repo.save_result_checked(sc_a, rep, st, scenario_id=sid, tags=["A.xlsx"])
    rid2, new2 = repo.save_result_checked(sc_a, rep, st, scenario_id=sid, tags=["A.xlsx"])
    assert new1 and not new2 and rid1 == rid2                    # повторне натискання – без дубліката
    # той самий набір параметрів з іншого файлу – той самий запис, додається тег
    sc_b = Scenario.from_legacy(LEG, name="B.xlsx – сценарій 7")
    assert repo.save_scenario(sc_b) == sid
    rid3, new3 = repo.save_result_checked(sc_b, rep, dict(st, source_file="B.xlsx"), tags=["B.xlsx"])
    assert rid3 == rid1 and not new3
    assert repo.get_result(rid1)["tags"] == ["A.xlsx", "B.xlsx"]
    # інший seed – це інший запуск
    _, new4 = repo.save_result_checked(sc_a, rep, {"replications": 2, "seed": 2}, tags=["A.xlsx"])
    assert new4
    assert len(repo.list_scenarios()) == 1 and len(repo.list_results()) == 2


def test_json_no_duplicates(tmp_path):
    _no_duplicates(JsonRepository(str(tmp_path)))


def test_json_deduplicate_old_records(tmp_path):
    """Записи версії 3.0 мали випадкові _id; dedupe зводить їх до одного."""
    from storage.repository import _new_id
    repo = JsonRepository(str(tmp_path))
    rep = run_replications(LEG, replications=2, base_seed=1)
    sc = Scenario.from_legacy(LEG, name="старий")
    for _ in range(3):                                   # імітація трьох натискань у версії 3.0
        sid = repo.save_scenario(sc, scenario_id=_new_id())
        doc = result_document(sc, rep, {"replications": 2, "seed": 1}, scenario_id=sid)
        doc["_id"] = _new_id()
        repo.store_result(doc)
    assert len(repo.list_results()) == 3 and len(repo.list_scenarios()) == 3
    info = repo.deduplicate()
    assert info == {"results_removed": 2, "scenarios_removed": 2}
    assert len(repo.list_results()) == 1 and len(repo.list_scenarios()) == 1


def test_legacy_result_document():
    p = dict(lambda1=1.0, s1=0.002, s2=0.5, N=3000, D1=1.0, lambda2=2.0, s2b=0.1, D2=0.2)
    rep = run_replications(p, replications=2, base_seed=1)
    doc = result_document(Scenario.from_legacy(p), rep, {"replications": 2, "seed": 1})
    assert len(doc["classic"]) == 18 and doc["per_type"] == []
    assert doc["classic"][16]["metric"] == 17


def test_mongo_repository():
    pytest.importorskip("pymongo")
    try:
        repo = MongoRepository(os.environ.get("REACTIVE_MONGO_URI", "mongodb://localhost:27017"),
                               db_name="reactive_simulation_test", timeout_ms=800)
    except Exception:
        pytest.skip("MongoDB не запущена")
    try:
        _roundtrip(repo)
        repo.client.drop_database("reactive_simulation_test")
        repo = MongoRepository(os.environ.get("REACTIVE_MONGO_URI", "mongodb://localhost:27017"),
                               db_name="reactive_simulation_test", timeout_ms=800)
        _no_duplicates(repo)
    finally:
        repo.client.drop_database("reactive_simulation_test")
