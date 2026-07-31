@echo off
title Numerar Certidoes

cd /d "%~dp0"

python numerar_traducoes.py

if errorlevel 1 (
    echo.
    echo Ocorreu um erro durante a execucao.
)

echo.
pause
