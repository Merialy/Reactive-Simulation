"""
Запуск сценарію з JSON-файлу з командного рядка.

    python -m simulation.run_scenario scenarios/boryviter_adcs.json -R 5 --seed 42
    python -m simulation.run_scenario scenarios/boryviter_fault_storm.json -R 5   (сценарій з відмовою)

Виводить показники для кожного типу подій (середнє ± 95 % довірчий інтервал).
"""
import argparse
import sys

from simulation.model import Scenario
from simulation.replications import run_scenario_replications

ROWS = [
    ("count", "Кількість оброблених подій", "{:.0f}"),
    ("wait_primary_mean", "Сер. очікування ISR, мс", "{:.4g}"),
    ("wait_secondary_mean", "Сер. очікування вторинної, мс", "{:.4g}"),
    ("sojourn_mean", "Сер. час реакції, мс", "{:.4g}"),
    ("sojourn_p95", "Час реакції, 95-й перцентиль, мс", "{:.4g}"),
    ("sojourn_p99", "Час реакції, 99-й перцентиль, мс", "{:.4g}"),
    ("sojourn_max", "Макс. час реакції, мс", "{:.4g}"),
    ("p_miss", "Ймовірність порушення дедлайну", "{:.4g}"),
    ("utilization_secondary", "Завантаження вторинним обробником", "{:.4g}"),
]

FAULT_ROWS = [
    ("count_fault", "ПІД ЧАС ВІДМОВИ: кількість подій", "{:.0f}"),
    ("sojourn_mean_fault", "  сер. час реакції, мс", "{:.4g}"),
    ("sojourn_max_fault", "  макс. час реакції, мс", "{:.4g}"),
    ("p_miss_fault", "  ймовірність порушення дедлайну", "{:.4g}"),
]


def main(argv=None):
    ap = argparse.ArgumentParser(description="Моделювання сценарію з JSON-файлу")
    ap.add_argument("path")
    ap.add_argument("-R", "--replications", type=int, default=5)
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--save", action="store_true", help="зберегти сценарій і результат у базу (MongoDB або JSON)")
    ap.add_argument("--tag", action="append", default=[], help="мітка результату, наприклад «серія Е1»")
    args = ap.parse_args(argv)

    sc = Scenario.load(args.path)
    errors, warnings = sc.check()
    print(f"Сценарій: {sc.name}")
    print(f"Платформа: {sc.platform.name}; T_IRQ = {sc.s1():.4g} мс; T_Core = {sc.platform.core_ms():.4g} мс")
    print(f"Завантаження U = {sc.utilization():.3f}")
    for e in errors:
        print("ПОМИЛКА:", e)
    if errors:
        return 1
    for w in warnings:
        print("⚠", w)

    res = run_scenario_replications(sc, replications=args.replications, base_seed=args.seed)
    names = [et.name for et in sc.event_types]
    w0 = max(len(r[1]) for r in ROWS) + 2
    header = "".join(f"{n[:22]:>24}" for n in names)
    print()
    print(" " * w0 + header)
    for key, label, fmt in ROWS:
        line = f"{label:<{w0}}"
        for n in names:
            v = res['per_type'][n][key]
            cell = fmt.format(v['mean'])
            if args.replications > 1 and key != "count":
                cell += f" ± {v['ci']:.2g}"
            line += f"{cell:>24}"
        print(line)
    if sc.faults:
        print("\nВідмови:")
        for f in sc.faults:
            print("  •", f.label())
        print(f"Завантаження під час відмов U = {sc.utilization_during_faults():.3f}")
        for key, label, fmt in FAULT_ROWS:
            line = f"{label:<{w0}}"
            for n in names:
                v = res['per_type'][n][key]
                cell = fmt.format(v['mean'])
                if args.replications > 1 and key != "count_fault":
                    cell += f" ± {v['ci']:.2g}"
                line += f"{cell:>24}"
            print(line)
        fs = res.get('faults', {})
        print(f"Хибних переривань: {fs.get('storm_events', {}).get('mean', 0):.0f}; "
              f"подій з перевищенням WCET: {fs.get('overrun_events', {}).get('mean', 0):.0f} (у середньому за прогін)")
    print(f"\nR = {args.replications}, seed = {args.seed}")

    if args.save:
        from storage import get_repository
        from storage.repository import last_storage_warning
        repo = get_repository()
        if last_storage_warning():
            print("⚠", last_storage_warning())
        sid = repo.save_scenario(sc)
        rid = repo.save_result(sc, res, {"replications": args.replications, "seed": args.seed},
                               scenario_id=sid, tags=args.tag)
        print(f"Збережено в {repo.describe()}: сценарій {sid}, результат {rid}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
