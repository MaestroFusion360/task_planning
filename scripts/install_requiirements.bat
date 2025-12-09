@echo off

:: Устанавливаем кодировку UTF-8
chcp 65001

:: Проверка наличия requirements.txt
IF NOT EXIST "requirements.txt" (
    echo requirements.txt не найден!
    exit /b 1
)

:: Проверка наличия виртуального окружения
IF NOT EXIST "..\.venv" (
    echo Виртуальное окружение не найдено, создаём его...
    python -m venv ..\.venv
)

:: Активация виртуального окружения
call ..\.venv\Scripts\activate.bat

:: Установка зависимостей
pip install -r requirements.txt

:: Задержка, чтобы окно не закрылось
echo Виртуальное окружение активно. Для выхода используйте команду 'deactivate'.
pause
