@echo off
chcp 65001 >nul
title Sud arizalari - Telegram bot
cd /d "%~dp0"

echo ============================================================
echo   TELEGRAM BOT - sud arizalarini shakllantirish
echo ============================================================
echo.

set PY=
where py >nul 2>nul && set PY=py
if "%PY%"=="" (where python >nul 2>nul && set PY=python)

if "%PY%"=="" (
    echo [!] Python topilmadi.
    echo     https://www.python.org/downloads/ dan o'rnating.
    echo     MUHIM: o'rnatishda "Add Python to PATH" belgisini qo'ying.
    echo.
    pause
    exit /b 1
)

%PY% -c "import telebot, docx, openpyxl, win32com.client" >nul 2>nul
if errorlevel 1 (
    echo Kerakli kutubxonalar o'rnatilmoqda, kuting...
    %PY% -m pip install --quiet --disable-pip-version-check pyTelegramBotAPI python-docx openpyxl pywin32
    echo.
)

echo Bot ishga tushmoqda... To'xtatish uchun bu oynani yoping.
echo.
%PY% "%~dp0bot.py"
