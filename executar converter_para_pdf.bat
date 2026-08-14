@echo off
title Converter para PDF

cd /d "%~dp0"

python converter_para_pdf.py

if errorlevel 1 (
    echo.
    echo Ocorreu um erro durante a execucao.
)

echo.
pause
