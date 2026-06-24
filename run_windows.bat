@echo off
chcp 65001 >nul
title Telegram Russian Groups Collector
python telegram_groups_windows.py
if errorlevel 1 (
    echo.
    echo Python не найден или произошла ошибка.
    echo Установи Python 3.10+ с https://www.python.org/downloads/ (галочка "Add to PATH")
    pause
)
