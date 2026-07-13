@echo off
title Separador de Certidoes

cd /d "%~dp0"

python separar_arquivos.py

if errorlevel 1 (
    echo.
    echo Ocorreu um erro durante a execucao.
)

echo.
pause