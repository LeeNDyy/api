# Используем базовый образ Python
FROM python:3.10-slim

# Устанавливаем рабочую директорию
WORKDIR /app

# Копируем requirements.txt и устанавливаем зависимости
COPY requirements.txt requirements.txt


RUN pip install --no-cache-dir -r requirements.txt

# Копируем все файлы приложения
COPY . .

# Используем Hypercorn для запуска приложения Quart
CMD ["hypercorn", "--bind", "0.0.0.0:5000", "mikroservices:app"]





