@echo off
chcp 65001 >nul
title Аналитика - DEBUG сборка
python -m pip install -r requirements.txt
python -m PyInstaller --noconfirm --clean Analitika_console.spec
pause
