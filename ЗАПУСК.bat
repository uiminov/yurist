@echo off
chcp 65001 >nul
title Формирование судебных заявлений
cd /d "%~dp0"

echo ============================================================
echo   ФОРМИРОВАНИЕ СУДЕБНЫХ ЗАЯВЛЕНИЙ
echo ============================================================
echo.

set PY=
where py >nul 2>nul && set PY=py
if "%PY%"=="" (where python >nul 2>nul && set PY=python)

if "%PY%"=="" (
    echo [!] Python не найден.
    echo     Установите его с https://www.python.org/downloads/
    echo     ВАЖНО: при установке поставьте галочку "Add Python to PATH".
    echo.
    pause
    exit /b 1
)

%PY% -c "import docx, openpyxl, win32com.client" >nul 2>nul
if errorlevel 1 (
    echo Устанавливаю необходимые библиотеки, подождите...
    %PY% -m pip install --quiet --disable-pip-version-check python-docx openpyxl pywin32
    echo.
)

%PY% "%~dp0generate_sud.py"
