"""
Тести логіки нового інтерфейсу, що не залежить від Qt (ui/core.py).
"""
import pytest

from ui.core import ParamTable, fmt, fmt_ci, parse_number, prepare_batch


def test_fmt():
    assert fmt(0.000183) == "0,000183"
    assert fmt(1.31449) == "1,314"
    assert fmt(10000) == "10000"
    assert fmt(0) == "0"
    assert fmt_ci(1.27715, 0.03203) == "1,277 ± 0,032"
    assert fmt(1e-6) == "0,000001"


def test_parse_number():
    assert parse_number("0,5") == 0.5
    assert parse_number(" 2.25 ") == 2.25
    assert parse_number("") is None
    with pytest.raises(ValueError):
        parse_number("abc")


def _table():
    t = ParamTable(columns=3)
    for c, s2 in enumerate((0.3, 0.5)):
        for r, v in enumerate((1, 0.01, s2, 5000, 1, 2, 0.1, 0.2)):
            t.set(r, c, v)
    t.set(0, 2, 1.0)                      # третій стовпець заповнено частково
    return t


def test_param_table_roundtrip(tmp_path):
    t = _table()
    p = tmp_path / "p.xlsx"
    t.save(str(p))
    t2, problems = ParamTable.load(str(p))
    assert problems == []
    assert t2.values == t.values
    assert t2.column_params(1)["s2"] == 0.5
    assert t2.column_u(0) == pytest.approx(1 * 0.31 + 2 * 0.11)


def test_prepare_batch():
    t = _table()
    t.set(1, 1, 5.0)                      # s1 = 5 мс > межі 1 мс
    run_list, warns, errs = prepare_batch(t)
    assert run_list[0] is not None
    assert run_list[1] is None and 1 in errs
    assert run_list[2] is None and "не заповнено" in errs[2]


def test_row_oriented_format(tmp_path):
    """«Рядковий» формат: перший рядок – назви параметрів, кожен наступний – сценарій."""
    import pandas as pd
    rows = [["lambda1", "s1", "s2", "N", "D1", "lambda2", "s2b", "D2"],
            [1, 0.01, 0.75, 10000, 1, 2, 0.1, 0.2],
            [1, 0.01, 0.5, 10000, 1, 2, 0.1, 0.2]]
    p = tmp_path / "rows.xlsx"
    pd.DataFrame(rows).to_excel(p, header=False, index=False)
    t, problems = ParamTable.load(str(p))
    assert t.n_cols == 2
    assert t.column_params(0)["s2"] == 0.75 and t.column_params(1)["s2"] == 0.5
    assert "рядковому" in problems[0]
