"""Переклад інтерфейсу: англійською не має лишатися українських фраз."""
import glob
import re

import pytest

from simulation.model import Scenario
from ui import i18n

CYR = re.compile("[А-Яа-яІіЇїЄєҐґ]")


@pytest.fixture
def en():
    i18n.set_lang("en")
    yield
    i18n.set_lang("uk")


def test_ready_scenarios_are_fully_english(en):
    from ui import explain
    from ui.core import delocalize_scenario, localize_scenario, type_label
    from ui.i18n import tr_msg
    for path in glob.glob("scenarios/*.json"):
        sc = Scenario.load(path)
        loc, back = localize_scenario(sc)
        assert delocalize_scenario(loc, back).to_dict() == sc.to_dict(), path     # у файл – початкові назви
        assert not loc.check()[0], path                                            # ціль відмови теж перекладено
        texts = [loc.name, loc.description] + [type_label(et.cls, et.name) for et in loc.event_types] + \
                [tr_msg(f.label()) for f in loc.active_faults()] + [tr_msg(m) for m in sum(loc.check(), [])]
        texts += [str(explain.scenario_facts(loc))]
        bad = [t for t in texts if CYR.search(t)]
        assert not bad, (path, bad)


def test_model_messages_are_translated(en):
    from ui.i18n import tr_msg
    for msg in ("Тривалість первинної обробки має бути більшою за 0.", "Тактова частота має бути більшою за 0.",
                "Тип «A»: інтенсивність має бути більшою за 0.", "Відмова «A»: інтенсивність шторму має бути більшою за 0.",
                "s2b має бути більшою за 0.", "MongoDB недоступна (timeout): використовуються JSON-файли."):
        assert not CYR.search(tr_msg(msg).replace("«A»", "")), tr_msg(msg)


def test_english_dictionary_has_no_cyrillic():
    from ui.i18n_en import EN, NAMES
    assert not [v for v in list(EN.values()) + list(NAMES.values()) if CYR.search(v)]
    assert all(set(re.findall(r"\{(\w+)\}", k)) == set(re.findall(r"\{(\w+)\}", v)) for k, v in EN.items())
