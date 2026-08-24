@echo off
chcp 65001 > nul
title Converter para PDF

cd /d "%~dp0"

python converter_para_pdf.py

if errorlevel 1 (
    echo.
    echo Ocorreu um erro durante a execucao.
)

echo.
pause

