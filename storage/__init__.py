"""
Підсистема збереження сценаріїв і результатів моделювання (версія 3.0).

Основне сховище – MongoDB. Якщо MongoDB недоступна (не запущена, не встановлено pymongo),
автоматично використовується запасне сховище у JSON-файлах (папка data/).

    from storage import get_repository
    repo = get_repository()
    sid = repo.save_scenario(scenario)
    rid = repo.save_result(scenario, result, settings={'replications': 5, 'seed': 42})
"""
from storage.repository import get_repository, Repository  # noqa: F401
