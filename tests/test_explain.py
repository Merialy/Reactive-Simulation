"""Пояснення простими словами (ui/explain.py): тексти формуються для всіх готових сценаріїв і відповідають результатам."""
import glob
import re

import pytest

pytest.importorskip("numpy")
from simulation.model import Scenario
from simulation.replications import run_scenario_replications
from ui.explain import pct, scenario_conclusion, scenario_facts


def _plain(html):
    return re.sub(r"<[^>]+>", "", html)


def test_pct():
    assert pct(0) == "0 %"
    assert pct(0.0004).startswith("<")
    assert pct(0.5).startswith("50")


@pytest.mark.parametrize("path", sorted(glob.glob("scenarios/*.json")))
def test_story_and_conclusion(path):
    sc = Scenario.load(path)
    info = scenario_facts(sc)
    assert len(info["facts"]) == 4
    assert (info["facts"][2][0] == "немає") == (not sc.active_faults())
    assert info["alerts"] and info["hint"]
    sc.N = 3000
    res = run_scenario_replications(sc, 2, 1)
    concl = _plain(scenario_conclusion(sc, res))
    hrt_miss = any(res["per_type"][et.name]["p_miss"]["mean"] > 0 for et in sc.event_types if et.cls == "HRT")
    assert ("Жорсткі дедлайни порушуються" in concl) == hrt_miss
