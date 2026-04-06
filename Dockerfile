# Используем Python 3.11
FROM python:3.11-slim

# Устанавливаем LibreOffice для конвертации PDF
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-calc \
    fonts-liberation \
    fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

# Рабочая директория
WORKDIR /app

# Копируем зависимости
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Копируем остальной код
COPY . .

# Создаем папки для временных файлов
RUN mkdir -p uploads temp_pdfs

# Выставляем порт
EXPOSE 5000

# Команда запуска
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "app:app"]
