"""
Перевірка роботи з MongoDB на цьому комп'ютері.

    python tools/check_mongodb.py

Що робить:
  1. Підключається до MongoDB (адреса з config.py або змінної REACTIVE_MONGO_URI).
  2. У ТИМЧАСОВІЙ базі reactive_simulation_check зберігає результати всіх готових сценаріїв
     і серії з таблиці параметрів, перевіряє: відсутність дублікатів, теги, різні R і seed,
     однаковість з JSON-сховищем, назви з крапками і символом $, видалення, dedupe,
     швидкість. Після перевірки тимчасова база видаляється.
  3. Вашу робочу базу reactive_simulation лише ЧИТАЄ: перевіряє, що всі збережені документи
     відкриваються, ключі відповідають вмісту і немає дублікатів. Нічого в ній не змінює.
  4. Перевіряє перехід на JSON-файли, якщо MongoDB недоступна.

Підсумок пишеться на екран і у файл logs/mongodb_check.txt.
"""
import copy
import glob
import os
import shutil
import sys
import tempfile
import time
import traceback

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
sys.path.insert(0, ROOT)

from config import Config                                               # noqa: E402
from simulation.model import Scenario                                   # noqa: E402
from simulation.replications import run_batch, run_scenario_replications  # noqa: E402
from storage import repository as R                                     # noqa: E402

CHECK_DB = "reactive_simulation_check"
LINES, PROBLEMS = [], []


def out(msg=""):
    print(msg, flush=True)
    LINES.append(msg)


def ok(cond, msg):
    out(("  ✔ " if cond else "  ✖ ") + msg)
    if not cond:
        PROBLEMS.append(msg)
    return cond


def clean(doc):
    """Документ без полів, що змінюються з часом або залежать від сховища."""
    d = {k: v for k, v in (doc or {}).items() if k not in ("_id", "created_at", "last_saved_at")}
    return d


def main():
    uri = os.environ.get("REACTIVE_MONGO_URI", Config.MONGO_URI)
    out(f"Перевірка MongoDB · програма {Config.APP_VERSION} · модель {Config.MODEL_VERSION} · {uri}")
    try:
        import pymongo
        out(f"pymongo {pymongo.version}")
    except ImportError:
        out("✖ pymongo не встановлено: pip install pymongo")
        return finish()

    # ------------------------------------------------------------------ 1. підключення
    out("\n1. Підключення")
    t0 = time.time()
    try:
        repo = R.MongoRepository(uri, db_name=CHECK_DB, timeout_ms=3000)
    except Exception as e:
        ok(False, f"не вдалося підключитися: {type(e).__name__}: {e}")
        out("  Запустіть MongoDB (docker compose up -d) і повторіть.")
        return finish()
    info = repo.client.server_info()
    ok(True, f"підключено за {time.time() - t0:.2f} с, MongoDB {info.get('version')}")
    repo.client.drop_database(CHECK_DB)
    repo = R.MongoRepository(uri, db_name=CHECK_DB, timeout_ms=3000)
    tmp = tempfile.mkdtemp(prefix="rs_mongo_check_")
    jrepo = R.JsonRepository(tmp)

    try:
        # -------------------------------------------------------------- 2. сценарії
        out("\n2. Збереження результатів готових сценаріїв (R = 2, N = 3000)")
        settings = {"replications": 2, "seed": 7}
        saved = []
        for path in sorted(glob.glob(os.path.join("scenarios", "*.json"))):
            sc = Scenario.load(path)
            sc.N = 3000
            res = run_scenario_replications(sc, 2, 7, workers=1)
            for rp in (repo, jrepo):
                sid = rp.save_scenario(sc)
                rid, new = rp.save_result_checked(sc, res, settings, scenario_id=sid, tags=[sc.name])
            ok(new, f"{os.path.basename(path)}: збережено ({rid})")
            saved.append((sc, res, rid))
            m, j = repo.get_result(rid), jrepo.get_result(rid)
            ok(m is not None and clean(m) == clean(j), "  документ у MongoDB такий самий, як у JSON-файлі")
            ok(bool(m.get("per_type")) and len(m["classic"]) == 18, "  є показники за типами і 18 основних")
            if sc.has_background():
                ok("bg_util" in (m.get("system") or {}), "  збережено показники фонових задач")
            if sc.active_faults():
                ok(bool(m.get("fault_stats")), "  збережено статистику відмов")
            back = Scenario.from_dict(m["scenario"])
            ok(back.to_dict() == sc.to_dict(), "  сценарій відновлюється з бази без змін")

        # -------------------------------------------------------------- 3. дублікати і теги
        out("\n3. Дублікати і теги")
        sc, res, rid = saved[1]
        n0 = len(repo.list_results(limit=10 ** 6))
        rid2, new = repo.save_result_checked(sc, res, settings, scenario_id=repo.save_scenario(sc), tags=["інший файл.xlsx"])
        ok(not new and rid2 == rid, "повторне збереження того самого запуску не створює копію")
        ok(len(repo.list_results(limit=10 ** 6)) == n0, "кількість результатів не змінилась")
        ok("інший файл.xlsx" in repo.get_result(rid)["tags"] and sc.name in repo.get_result(rid)["tags"],
           "новий тег дописано до старого")
        sc2 = copy.deepcopy(sc)
        sc2.name = "Та сама конфігурація, інша назва"
        ok(repo.save_scenario(sc2) == repo.save_scenario(sc), "сценарій з іншою назвою, але тими самими параметрами – той самий запис")
        doc = repo.scenarios.find_one({"_id": repo.save_scenario(sc)})
        ok(sc2.name in doc.get("names", []) and sc.name in doc.get("names", []), "обидві назви збережено в полі names")
        rid3, new = repo.save_result_checked(sc, res, {"replications": 2, "seed": 8})
        ok(new and rid3 != rid, "інший seed – окремий запис")
        rid4, new = repo.save_result_checked(sc, res, {"replications": 3, "seed": 7})
        ok(new and rid4 not in (rid, rid3), "інше R – окремий запис")

        # -------------------------------------------------------------- 4. серія з таблиці
        out("\n4. Серія з таблиці параметрів (як кнопка «Зберегти в базу даних»)")
        P = [dict(lambda1=1.0, s1=0.01, s2=s2, N=3000, D1=1.0, lambda2=2.0, s2b=0.1, D2=0.2) for s2 in (0.3, 0.5)]
        batch = run_batch(P, 2, 7, 1, None)
        source = "Тест v1.2 $special.xlsx"
        for c, (p, b) in enumerate(zip(P, batch), 1):
            s = Scenario.from_legacy(p, name=f"{source} – сценарій {c}")
            keep = {k: b[k] for k in ("mean", "ci", "extra") if b.get(k)}
            st = dict(settings, source_file=source)
            sid = repo.save_scenario(s)
            r, new = repo.save_result_checked(s, keep, st, scenario_id=sid, tags=[source])
            d = repo.get_result(r)
            ok(new and len(d["classic"]) == 18, f"стовпець {c}: збережено, 18 показників")
            ok(d["tags"] == [source] and d["settings"]["source_file"] == source,
               "  назва файлу з крапкою і $ зберігається без помилок")
            ok(bool(d.get("extra")), "  додаткові показники серії теж збережено")

        # -------------------------------------------------------------- 5. список, видалення
        out("\n5. Список, пошук, видалення")
        rows = repo.list_results(limit=2000)
        ok(all(r.get("created_at") for r in rows), f"у списку {len(rows)} результатів, у всіх є дата")
        ok([r["created_at"] for r in rows] == sorted((r["created_at"] for r in rows), reverse=True),
           "список упорядковано від нових до старих")
        ok(repo.delete_result(rid4) and repo.get_result(rid4) is None, "видалення результату")
        ok(not repo.delete_result("немає-такого"), "видалення неіснуючого – без помилки")

        # -------------------------------------------------------------- 6. dedupe
        out("\n6. Очищення дублікатів версії 3.0 (storage.cli dedupe)")
        old = clean(repo.get_result(rid))
        for _ in range(2):                                   # дві «старі» копії з випадковими _id
            repo.results.insert_one(dict(old, _id=R._new_id(), created_at=R._now()))
        n_before = len(repo.list_results(limit=10 ** 6))
        info = repo.deduplicate()
        n_after = len(repo.list_results(limit=10 ** 6))
        ok(info["results_removed"] == 2 and n_after == n_before - 2, f"видалено копій: {info['results_removed']}")
        ok(repo.get_result(rid) is not None, "оригінальний запис лишився")

        # -------------------------------------------------------------- 7. швидкість
        out("\n7. Швидкість")
        t0 = time.time()
        for k in range(100):
            repo.save_result_checked(sc, res, {"replications": 2, "seed": 1000 + k})
        dt = time.time() - t0
        ok(dt < 30, f"100 збережень за {dt:.2f} с")
        t0 = time.time()
        n = len(repo.list_results(limit=2000))
        ok(time.time() - t0 < 5, f"список з {n} результатів за {time.time() - t0:.2f} с")
    except Exception:
        ok(False, "неочікувана помилка:\n" + traceback.format_exc())
    finally:
        repo.client.drop_database(CHECK_DB)
        shutil.rmtree(tmp, ignore_errors=True)
        out("\n  тимчасову базу і файли видалено")

    # ------------------------------------------------------------------ 8. робоча база (лише читання)
    out(f"\n8. Ваша робоча база {R.DB_NAME} (лише читання)")
    try:
        main_repo = R.MongoRepository(uri, timeout_ms=3000)
        res_docs = list(main_repo.results.find({}))
        sc_docs = list(main_repo.scenarios.find({}))
        out(f"  результатів: {len(res_docs)}, сценаріїв: {len(sc_docs)}")
        broken, wrong_key, versions = [], [], {}
        keys = {}
        for d in res_docs:
            versions[d.get("app_version", "?")] = versions.get(d.get("app_version", "?"), 0) + 1
            try:
                s = Scenario.from_dict(d["scenario"])
                k = R.run_key(s, d.get("settings") or {}, d.get("model_version"))
                keys.setdefault(k, []).append(d["_id"])
                if d["_id"] != k:
                    wrong_key.append(d["_id"])
                assert d.get("classic") is not None and d.get("created_at")
            except Exception as e:
                broken.append(f"{d.get('_id')}: {type(e).__name__}: {e}")
        ok(not broken, "усі результати відкриваються" + ("" if not broken else ": " + "; ".join(broken[:5])))
        dups = {k: v for k, v in keys.items() if len(v) > 1}
        ok(not dups, "дублікатів немає" if not dups else
           f"є дублікати ({sum(len(v) - 1 for v in dups.values())}) – виконайте python -m storage.cli dedupe")
        ok(not wrong_key, "ключі результатів відповідають вмісту" if not wrong_key else
           f"{len(wrong_key)} записів зі старими ключами (версія 3.0) – виконайте python -m storage.cli dedupe")
        out("  версії програми в записах: " + ", ".join(f"{v}: {n}" for v, n in sorted(versions.items())))
        bad_sc = []
        for d in sc_docs:
            try:
                Scenario.from_dict(d)
            except Exception as e:
                bad_sc.append(f"{d.get('_id')}: {e}")
        ok(not bad_sc, "усі сценарії відкриваються" + ("" if not bad_sc else ": " + "; ".join(bad_sc[:5])))
    except Exception as e:
        ok(False, f"не вдалося прочитати робочу базу: {type(e).__name__}: {e}")

    # ------------------------------------------------------------------ 9. MongoDB недоступна
    out("\n9. Якщо MongoDB недоступна")
    saved_env = os.environ.get("REACTIVE_MONGO_URI")
    os.environ["REACTIVE_MONGO_URI"] = "mongodb://127.0.0.1:1"
    try:
        t0 = time.time()
        r = R.get_repository("auto")
        dt = time.time() - t0
        ok(r.kind == "JSON", f"програма перейшла на JSON-файли за {dt:.1f} с")
        ok(bool(R.last_storage_warning()), "з'явилося попередження: " + R.last_storage_warning())
    finally:
        if saved_env is None:
            os.environ.pop("REACTIVE_MONGO_URI", None)
        else:
            os.environ["REACTIVE_MONGO_URI"] = saved_env
    return finish()


def finish():
    out(f"\nПРОБЛЕМ: {len(PROBLEMS)}")
    for p in PROBLEMS:
        out("  ✖ " + p.splitlines()[0])
    os.makedirs("logs", exist_ok=True)
    with open(os.path.join("logs", "mongodb_check.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(LINES) + "\n")
    print("Звіт: logs/mongodb_check.txt")
    return 1 if PROBLEMS else 0


if __name__ == "__main__":
    sys.exit(main())
