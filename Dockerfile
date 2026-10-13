# Берём готовый образ с Python
FROM python:3.10-slim

# Устанавливаем рабочий каталог внутри контейнера
WORKDIR /app

# Копируем файлы проекта внутрь контейнера
COPY . .

# Устанавливаем все нужные библиотеки
RUN pip install pandas matplotlib openpyxl numpy

# Команда для запуска приложения
CMD ["python", "main.py"]