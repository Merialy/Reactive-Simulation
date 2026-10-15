# Контейнер для відтворюваних обчислювальних експериментів (без графічного інтерфейсу).
#
# Графічна програма запускається на комп'ютері користувача: python app.py.
# У контейнері запускаються тести верифікації та розрахунок сценаріїв з командного рядка:
#   docker compose run --rm sim                                   – тести (за замовчуванням)
#   docker compose run --rm sim python -m simulation.run_scenario scenarios/boryviter_adcs.json -R 5 --save
#   docker compose run --rm sim python -m storage.cli results
FROM python:3.12-slim

# Не створювати .pyc і одразу виводити текст у консоль
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PYTHONIOENCODING=utf-8 \
    MPLBACKEND=Agg

WORKDIR /app

# Спочатку лише залежності – щоб Docker кешував цей крок
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Код проєкту
COPY . .

CMD ["python", "-m", "pytest", "-q"]
