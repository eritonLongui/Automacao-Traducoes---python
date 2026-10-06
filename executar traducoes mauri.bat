@echo off
chcp 65001 > nul
title Traducoes Mauri

cd /d "%~dp0"

echo ======================================
echo Executando numerar_traducoes.py...
echo ======================================
python numerar_traducoes.py

if errorlevel 1 (
    echo.
    echo Ocorreu um erro ao executar numerar_traducoes.py.
    echo O fluxo de conversao para PDF foi interrompido.
    echo.
    pause
    exit /b 1
)

echo.
echo ======================================
echo Executando converter_para_pdf.py...
echo ======================================
python converter_para_pdf.py

if errorlevel 1 (
    echo.
    echo Ocorreu um erro ao executar converter_para_pdf.py.
    echo.
    pause
    exit /b 1
)

echo.
echo ======================================
echo Fluxo Mauri concluido com sucesso!
echo ======================================
echo.
pause