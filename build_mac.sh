#!/bin/bash
set -e

echo "=== TelegramParser — сборка для macOS ==="
echo ""

# Проверка Python
if ! command -v python3 &>/dev/null; then
    echo "ОШИБКА: python3 не найден."
    echo "Установите Python 3 с https://www.python.org/downloads/"
    exit 1
fi

echo "Python: $(python3 --version)"

# Проверка tkinter
python3 -c "import tkinter" 2>/dev/null || {
    echo ""
    echo "ОШИБКА: tkinter не найден."
    echo "Нужен Python с поддержкой Tk."
    echo "Установите Python с https://www.python.org/downloads/"
    echo "(НЕ через homebrew без --with-tcl-tk)"
    exit 1
}

echo "tkinter: OK"
echo ""

# Зависимости
echo "--- Установка зависимостей ---"
pip3 install -r requirements.txt
pip3 install pyinstaller
echo ""

# Очистка
rm -rf build dist

# Сборка
echo "--- Сборка .app ---"
pyinstaller \
    --windowed \
    --name "TelegramParser" \
    --add-data "config.json.example:." \
    main.py

echo ""
echo "=== Готово ==="
echo ""
echo "Приложение: dist/TelegramParser.app"
cp config.json.example dist/config.json.example
echo "Пример настроек: dist/config.json.example"
echo ""
echo "Чтобы упаковать для передачи:"
echo "  cd dist && zip -r TelegramParser-mac.zip TelegramParser.app config.json.example"
