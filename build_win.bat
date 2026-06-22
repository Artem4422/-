@echo off
echo === TelegramParser — сборка для Windows ===
echo.

pip install -r requirements.txt
pip install pyinstaller

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

pyinstaller --onefile --windowed --name "TelegramParser" --add-data "config.json;." main.py

echo.
echo === Готово ===
echo Файл: dist\TelegramParser.exe
pause
