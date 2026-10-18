"""
Робота зі сховищем з командного рядка.

    python -m storage.cli status                 – яке сховище використовується
    python -m storage.cli import scenarios/*.json – зберегти сценарії з файлів у базу
    python -m storage.cli scenarios              – список сценаріїв
    python -m storage.cli results                – список збережених результатів
    python -m storage.cli show <id результату>   – показники по типах подій
    python -m storage.cli dedupe                 – прибрати дублікати, збережені версією 3.0
"""
import argparse
import glob
import sys

from storage.repository import get_repository, last_storage_warning


def main(argv=None):
    ap = argparse.ArgumentParser(description="Сховище сценаріїв і результатів")
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("status")
    imp = sub.add_parser("import"); imp.add_argument("files", nargs="+")
    sub.add_parser("scenarios")
    sub.add_parser("results")
    sh = sub.add_parser("show"); sh.add_argument("id")
    sub.add_parser("dedupe")
    args = ap.parse_args(argv)

    repo = get_repository()
    if last_storage_warning():
        print("⚠", last_storage_warning())
    print("Сховище:", repo.describe())

    if args.cmd == "import":
        from simulation.model import Scenario
        for pattern in args.files:
            for path in glob.glob(pattern) or [pattern]:
                sc = Scenario.load(path)
                print(f"  {repo.save_scenario(sc)}  {sc.name}")
    elif args.cmd == "scenarios":
        for s in repo.list_scenarios():
            print(f"  {s['id']}  {s['updated_at']}  типів: {s['n_types']}  {s['name']}")
    elif args.cmd == "results":
        for r in repo.list_results():
            st = r.get("settings") or {}
            print(f"  {r['id']}  {r['created_at']}  R={st.get('replications')} seed={st.get('seed')}  "
                  f"U={r.get('utilization', 0):.3f}  {r['scenario_name']}")
    elif args.cmd == "dedupe":
        before_r, before_s = len(repo.list_results(limit=10 ** 9)), len(repo.list_scenarios())
        info = repo.deduplicate()
        print(f"  Результатів: було {before_r}, видалено дублікатів {info['results_removed']}, "
              f"лишилось {len(repo.list_results(limit=10 ** 9))}")
        print(f"  Сценаріїв:   було {before_s}, видалено дублікатів {info['scenarios_removed']}, "
              f"лишилось {len(repo.list_scenarios())}")
    elif args.cmd == "show":
        doc = repo.get_result(args.id)
        if not doc:
            print("Результат не знайдено"); return 1
        print(f"{doc['scenario_name']}  ({doc['created_at']}, R={doc['settings'].get('replications')})")
        for t in doc.get("per_type", []):
            m = t["metrics"]
            g = lambda k: (m.get(k) or {}).get("mean", float("nan"))
            print(f"  {t['name']:<28} [{t['class']}]  T_сер = {g('sojourn_mean'):.4g} мс,"
                  f"  p99 = {g('sojourn_p99'):.4g} мс,  P_miss = {g('p_miss'):.4g}")
        if not doc.get("per_type"):                      # серія з таблиці параметрів – 18 показників
            for c in doc.get("classic", []):
                print(f"  {c['metric']:>2}. {c['name']:<55} {c['mean']:.4g} ± {c['ci']:.2g}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
