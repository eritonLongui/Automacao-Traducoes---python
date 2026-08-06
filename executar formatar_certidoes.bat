@echo off
title Formatar Certidoes

cd /d "%~dp0"

python formatar_certidoes.py

if errorlevel 1 (
    echo.
    echo Ocorreu um erro durante a execucao.
)

echo.
pause
