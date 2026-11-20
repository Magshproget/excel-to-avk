@echo off
chcp 65001 > nul
echo ====================================================
echo  Збірка Конвертера для АВК-5 в один .EXE файл
echo ====================================================
echo.

echo 1. Встановлення необхідних бібліотек...
python -m pip install --upgrade pip
python -m pip install openpyxl pyinstaller

echo.
echo 2. Створення файлу Конвертер_АВК.exe...
pyinstaller --noconsole --onefile --name "Конвертер_АВК" app_gui.py

echo.
echo ====================================================
echo  Готово! 
echo  Ваш файл знаходиться у папці: dist\Конвертер_АВК.exe
echo ====================================================
pause
