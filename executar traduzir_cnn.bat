@echo off
title Separador de Certidoes

cd /d "%~dp0"

python traduzir_cnn.py

if errorlevel 1 (
    echo.
    echo Ocorreu um erro durante a execucao.
)

echo.
pause