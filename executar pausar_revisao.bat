@echo off
title Pausar / Retomar Revisao

cd /d "%~dp0"

python pausar_revisao.py

if errorlevel 1 (
    echo.
    echo Ocorreu um erro durante a execucao.
)

echo.
pause
