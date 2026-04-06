# Используем официальный образ Python
FROM python:3.11-slim

# Устанавливаем системные зависимости для Playwright и браузеров
RUN apt-get update && apt-get install -y \
    wget \
    gnupg \
    && rm -rf /var/lib/apt/lists/*

# Рабочая директория
WORKDIR /app

# Копируем зависимости
COPY requirements.txt .

# Устанавливаем зависимости Python
RUN pip install --no-cache-dir -r requirements.txt

# Устанавливаем Playwright и браузер Chromium
# Мы устанавливаем только Chromium, чтобы образ был легче
RUN playwright install --with-deps chromium

# Копируем остальной код
COPY . .

# Создаем папки для временных файлов
RUN mkdir -p uploads temp_pdfs

# Переменная окружения для Flask
ENV FLASK_APP=app.py
ENV PYTHONUNBUFFERED=1

# Порт для Render
EXPOSE 10000

# Команда запуска (используем gunicorn для продакшена)
CMD ["gunicorn", "--bind", "0.0.0.0:10000", "app:app"]
