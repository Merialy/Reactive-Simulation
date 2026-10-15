"""
Репозиторій сценаріїв і результатів.

Колекції (MongoDB) / папки (JSON):
  scenarios – конфігурації сценаріїв (платформа, типи подій, N);
  results   – підсумкові результати запусків (середні ± ДІ, показники по типах, налаштування R і seed).

Структура документа результату:
{
  "_id": "...",
  "scenario_id": "...",            # посилання на сценарій (якщо його збережено)
  "scenario": {...},               # знімок сценарію на момент запуску
  "scenario_name": "...",
  "created_at": "2026-10-13T21:10:00",
  "app_version": "3.0",
  "settings": {"replications": 5, "seed": 42},
  "utilization": 0.465,
  "classic": [{"metric": 1, "name": "...", "mean": ..., "ci": ...}, ...],
  "per_type": [{"name": "ADCS: датчики", "class": "HRT", "metrics": {"sojourn_mean": {"mean":..,"ci":..}, ...}}],
  "tags": ["серія Е1"], "note": "..."
}

Захист від дублікатів (версія 3.1)
  * Ідентифікатор сценарію = хеш його ВМІСТУ (платформа, типи подій, N, частка розгону),
    без назви та опису. Той самий набір параметрів завжди має той самий _id.
  * Ідентифікатор результату (run_key) = хеш (сценарій + R + seed + MODEL_VERSION).
    Модель детермінована: за однакових входів результат той самий, тому повторне
    збереження не створює нового документа, а лише дописує нові теги (ім'я файлу).
  * MODEL_VERSION змінюється тільки тоді, коли змінюється сама модель (числа),
    і тоді нові результати зберігаються окремо від старих.
"""
import hashlib
import json
import os
import uuid
from datetime import datetime
from typing import Dict, List, Optional

from config import Config

DB_NAME = "reactive_simulation"


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _new_id() -> str:
    return uuid.uuid4().hex[:24]


# поля сценарію, що НЕ впливають на результат і не входять у ключ
_NAME_FIELDS = ("name", "description", "_id", "updated_at", "created_at", "names")
# налаштування запуску, що впливають на результат
RUN_SETTINGS = ("replications", "seed")


def _canon(x):
    """Нормалізація для хешу: 0.1 і 0.10000000000000002 вважаються однаковими."""
    if isinstance(x, bool) or x is None:
        return x
    if isinstance(x, (int, float)):
        x = float(x)
        return float(f"{x:.12g}")
    if isinstance(x, dict):
        return {str(k): _canon(v) for k, v in x.items()}
    if isinstance(x, (list, tuple)):
        return [_canon(v) for v in x]
    return x


def _hash(obj) -> str:
    text = json.dumps(_canon(obj), sort_keys=True, ensure_ascii=False)
    return hashlib.sha1(text.encode("utf-8")).hexdigest()[:24]


def scenario_key(scenario) -> str:
    """Ідентифікатор сценарію за вмістом (назва не враховується)."""
    d = scenario if isinstance(scenario, dict) else scenario.to_dict()
    return _hash({k: v for k, v in d.items() if k not in _NAME_FIELDS})


def run_key(scenario, settings: Dict, model_version: Optional[str] = None) -> str:
    """Ідентифікатор запуску: той самий сценарій + ті самі R і seed + та сама модель."""
    s = {k: (settings or {}).get(k) for k in RUN_SETTINGS}
    return _hash({"scenario": scenario_key(scenario), "settings": s,
                  "model": model_version or Config.MODEL_VERSION})


def result_document(scenario, result: Dict, settings: Dict, scenario_id: Optional[str] = None,
                    tags: Optional[List[str]] = None, note: str = "") -> Dict:
    """
    Перетворює результат run_scenario_replications(...) або run_replications(...) на документ БД.
    result: {'classic': {'mean': {1..18}, 'ci': {1..18}}, 'per_type': {...}, 'U': float}
            або {'mean': {...}, 'ci': {...}} (формат версії 1.x).
    """
    classic = result.get('classic', result)
    mean, ci = classic.get('mean', {}), classic.get('ci', {}) or {}
    classic_list = []
    for m in range(1, 19):
        if m in mean:
            classic_list.append({"metric": m, "name": Config.METRIC_NAMES[m - 1],
                                 "mean": float(mean[m]), "ci": float(ci.get(m, 0.0) or 0.0)})
    classes = {et.name: et.cls for et in scenario.event_types}
    per_type = [{"name": name, "class": classes.get(name, ""), "metrics": metrics}
                for name, metrics in (result.get('per_type') or {}).items()]
    skey = scenario_key(scenario)
    return {
        "run_key": run_key(scenario, settings),
        "scenario_key": skey,
        "scenario_id": scenario_id or skey,
        "scenario": scenario.to_dict(),
        "scenario_name": scenario.name,
        "created_at": _now(),
        "app_version": Config.APP_VERSION,
        "model_version": Config.MODEL_VERSION,
        "settings": dict(settings),
        "utilization": float(result.get('U', scenario.utilization())),
        "classic": classic_list,
        "per_type": per_type,
        "tags": list(tags or []),
        "note": note,
        **({"fault_stats": result["faults"]} if result.get("faults") else {}),
        **({"system": result["system"]} if result.get("system") else {}),
        **({"extra": result["extra"]} if result.get("extra") else {}),
    }


class Repository:
    """Спільний інтерфейс сховища."""
    kind = "base"

    # --- сценарії ---
    def save_scenario(self, scenario, scenario_id: Optional[str] = None) -> str: ...
    def get_scenario(self, scenario_id: str): ...
    def list_scenarios(self) -> List[Dict]: ...
    def delete_scenario(self, scenario_id: str) -> bool: ...
    # --- результати ---
    def insert_result(self, doc: Dict) -> str: ...
    def get_result(self, result_id: str) -> Optional[Dict]: ...
    def list_results(self, scenario_id: Optional[str] = None, limit: int = 100) -> List[Dict]: ...
    def delete_result(self, result_id: str) -> bool: ...

    def store_result(self, doc: Dict):
        """Зберігає документ результату. Повертає (id, True), якщо це новий запуск,
        або (id, False), якщо такий запуск уже був (тоді лише дописуються теги)."""
        ...

    def insert_result(self, doc: Dict) -> str:
        return self.store_result(doc)[0]

    def save_result(self, scenario, result: Dict, settings: Dict, scenario_id: Optional[str] = None,
                    tags: Optional[List[str]] = None, note: str = "") -> str:
        return self.save_result_checked(scenario, result, settings, scenario_id, tags, note)[0]

    def save_result_checked(self, scenario, result: Dict, settings: Dict, scenario_id: Optional[str] = None,
                            tags: Optional[List[str]] = None, note: str = ""):
        """Як save_result, але повертає (id, чи_новий)."""
        return self.store_result(result_document(scenario, result, settings, scenario_id, tags, note))

    def deduplicate(self) -> Dict[str, int]:
        """
        Переводить старі записи (з випадковими _id, версія 3.0) на ключі за вмістом
        і видаляє дублікати. Повертає кількість видалених документів.
        """
        from simulation.model import Scenario
        removed_r = removed_s = 0
        for r in self.list_results(limit=10 ** 9):
            doc = self.get_result(r["id"])
            if not doc:
                continue
            sc = Scenario.from_dict(doc["scenario"])
            key = run_key(sc, doc.get("settings") or {}, doc.get("model_version"))
            if doc["_id"] == key:
                continue
            self.delete_result(doc["_id"])
            doc = {k: v for k, v in doc.items() if k != "_id"}
            doc.update({"run_key": key, "scenario_key": scenario_key(sc), "scenario_id": scenario_key(sc)})
            doc.setdefault("model_version", Config.MODEL_VERSION)
            if not self.store_result(doc)[1]:
                removed_r += 1
        for s in self.list_scenarios():
            sc = self.get_scenario(s["id"])
            key = scenario_key(sc)
            if s["id"] == key:
                continue
            existed = self.get_scenario(key) is not None
            self.delete_scenario(s["id"])
            self.save_scenario(sc)
            if existed:
                removed_s += 1
        return {"results_removed": removed_r, "scenarios_removed": removed_s}

    def describe(self) -> str:
        return self.kind


# ---------------------------------------------------------------------------
class MongoRepository(Repository):
    kind = "MongoDB"

    def __init__(self, uri: str, db_name: str = DB_NAME, timeout_ms: int = 1500):
        from pymongo import MongoClient, DESCENDING
        self._DESC = DESCENDING
        self.client = MongoClient(uri, serverSelectionTimeoutMS=timeout_ms)
        self.client.admin.command("ping")             # перевірка з'єднання
        self.db = self.client[db_name]
        self.scenarios = self.db["scenarios"]
        self.results = self.db["results"]
        self.uri = uri
        self.results.create_index([("scenario_id", 1), ("created_at", -1)])
        self.scenarios.create_index("name")

    def describe(self) -> str:
        return f"MongoDB ({self.uri}, база {self.db.name})"

    def save_scenario(self, scenario, scenario_id=None):
        sid = scenario_id or scenario_key(scenario)
        doc = scenario.to_dict()
        doc["updated_at"] = _now()
        self.scenarios.update_one({"_id": sid},
                                  {"$set": doc, "$setOnInsert": {"created_at": doc["updated_at"]},
                                   "$addToSet": {"names": scenario.name}},
                                  upsert=True)
        return sid

    def get_scenario(self, scenario_id):
        from simulation.model import Scenario
        doc = self.scenarios.find_one({"_id": scenario_id})
        return Scenario.from_dict(doc) if doc else None

    def list_scenarios(self):
        cur = self.scenarios.find({}, {"name": 1, "updated_at": 1, "event_types": 1}).sort("updated_at", self._DESC)
        return [{"id": d["_id"], "name": d.get("name", ""), "updated_at": d.get("updated_at", ""),
                 "n_types": len(d.get("event_types", []))} for d in cur]

    def delete_scenario(self, scenario_id):
        return self.scenarios.delete_one({"_id": scenario_id}).deleted_count == 1

    def store_result(self, doc):
        doc = dict(doc)
        rid = doc.pop("_id", None) or doc.get("run_key") or _new_id()
        tags = list(doc.pop("tags", []) or [])
        doc.pop("last_saved_at", None)          # інакше конфлікт полів $set / $setOnInsert
        res = self.results.update_one(
            {"_id": rid},
            {"$setOnInsert": doc, "$addToSet": {"tags": {"$each": tags}},
             "$set": {"last_saved_at": _now()}},
            upsert=True)
        return rid, res.upserted_id is not None

    def get_result(self, result_id):
        return self.results.find_one({"_id": result_id})

    def list_results(self, scenario_id=None, limit=100):
        q = {"scenario_id": scenario_id} if scenario_id else {}
        cur = self.results.find(q, {"scenario_name": 1, "created_at": 1, "settings": 1, "utilization": 1,
                                    "scenario_id": 1, "tags": 1}).sort("created_at", self._DESC).limit(limit)
        return [{"id": d["_id"], **{k: v for k, v in d.items() if k != "_id"}} for d in cur]

    def delete_result(self, result_id):
        return self.results.delete_one({"_id": result_id}).deleted_count == 1


# ---------------------------------------------------------------------------
class JsonRepository(Repository):
    """Запасне сховище: по одному JSON-файлу на документ."""
    kind = "JSON"

    def __init__(self, folder: str):
        self.folder = folder
        self.sdir = os.path.join(folder, "scenarios")
        self.rdir = os.path.join(folder, "results")
        os.makedirs(self.sdir, exist_ok=True)
        os.makedirs(self.rdir, exist_ok=True)

    def describe(self) -> str:
        return f"JSON-файли ({os.path.abspath(self.folder)})"

    @staticmethod
    def _write(path, doc):
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(doc, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)

    @staticmethod
    def _read(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)

    def _all(self, d):
        out = []
        for name in os.listdir(d):
            if name.endswith(".json"):
                try:
                    out.append(self._read(os.path.join(d, name)))
                except (OSError, ValueError):
                    continue
        return out

    def save_scenario(self, scenario, scenario_id=None):
        sid = scenario_id or scenario_key(scenario)
        path = os.path.join(self.sdir, sid + ".json")
        old = self._read(path) if os.path.exists(path) else {}
        doc = scenario.to_dict()
        now = _now()
        names = list(old.get("names", []))
        if scenario.name not in names:
            names.append(scenario.name)
        doc.update({"_id": sid, "updated_at": now, "created_at": old.get("created_at", now), "names": names})
        self._write(path, doc)
        return sid

    def get_scenario(self, scenario_id):
        from simulation.model import Scenario
        p = os.path.join(self.sdir, scenario_id + ".json")
        return Scenario.from_dict(self._read(p)) if os.path.exists(p) else None

    def list_scenarios(self):
        docs = sorted(self._all(self.sdir), key=lambda d: d.get("updated_at", ""), reverse=True)
        return [{"id": d["_id"], "name": d.get("name", ""), "updated_at": d.get("updated_at", ""),
                 "n_types": len(d.get("event_types", []))} for d in docs]

    def delete_scenario(self, scenario_id):
        p = os.path.join(self.sdir, scenario_id + ".json")
        if os.path.exists(p):
            os.remove(p)
            return True
        return False

    def store_result(self, doc):
        doc = dict(doc)
        rid = doc.pop("_id", None) or doc.get("run_key") or _new_id()
        path = os.path.join(self.rdir, rid + ".json")
        is_new = not os.path.exists(path)
        if is_new:
            doc["_id"] = rid
            doc["tags"] = list(dict.fromkeys(doc.get("tags") or []))
        else:
            old = self._read(path)
            old["tags"] = list(dict.fromkeys(list(old.get("tags") or []) + list(doc.get("tags") or [])))
            doc = old
        doc["last_saved_at"] = _now()
        self._write(path, doc)
        return rid, is_new

    def get_result(self, result_id):
        p = os.path.join(self.rdir, result_id + ".json")
        return self._read(p) if os.path.exists(p) else None

    def list_results(self, scenario_id=None, limit=100):
        docs = [d for d in self._all(self.rdir) if not scenario_id or d.get("scenario_id") == scenario_id]
        docs.sort(key=lambda d: d.get("created_at", ""), reverse=True)
        keys = ("scenario_name", "created_at", "settings", "utilization", "scenario_id", "tags")
        return [{"id": d["_id"], **{k: d.get(k) for k in keys}} for d in docs[:limit]]

    def delete_result(self, result_id):
        p = os.path.join(self.rdir, result_id + ".json")
        if os.path.exists(p):
            os.remove(p)
            return True
        return False


# ---------------------------------------------------------------------------
_last_error = ""


def get_repository(prefer: Optional[str] = None) -> Repository:
    """
    Повертає сховище: MongoDB, якщо доступна, інакше JSON-файли.
    prefer: "mongo" / "json" / None (з налаштувань Config.STORAGE).
    """
    global _last_error
    prefer = (prefer or Config.STORAGE).lower()
    if prefer in ("mongo", "auto"):
        uri = os.environ.get("REACTIVE_MONGO_URI", Config.MONGO_URI)
        try:
            return MongoRepository(uri)
        except Exception as e:                       # pymongo не встановлено або сервер недоступний
            _last_error = f"MongoDB недоступна ({type(e).__name__}): використовуються JSON-файли."
            if prefer == "mongo":
                raise
    return JsonRepository(Config.JSON_STORAGE_DIR)


def last_storage_warning() -> str:
    return _last_error
