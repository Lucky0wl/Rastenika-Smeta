# Инструкция по деплою (Linux Hosting)

Чтобы PDF-генератор работал корректно на Linux-сервере (Ubuntu/Debian), необходимо выполнить следующие шаги:

1. **Установите зависимости Python**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Установите Playwright и его системные зависимости**:
   ```bash
   # Установка самого playwright
   pip install playwright
   # Установка браузера Chromium и необходимых библиотек Linux
   playwright install chromium
   playwright install-deps
   ```

3. **Запуск через Gunicorn**:
   ```bash
   gunicorn -w 4 -b 0.0.0.0:8000 app:app
   ```

## Почему это лучше LibreOffice?
1. **Качество 1:1**: PDF будет выглядеть точно так же, как если бы вы открыли его в браузере.
2. **Шрифты**: Используется современный шрифт Inter, который подгружается автоматически.
3. **Надежность**: Playwright — это движок Google Chrome, самый стабильный инструмент для рендеринга на сегодняшний день.
