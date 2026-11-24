@echo off
chcp 65001 > nul
title Solicitar Mauri
cd /d "%~dp0"

echo ======================================
echo Executando solicitar_mauri.py...
echo ======================================
echo.

python solicitar_mauri.py

echo.
pause
