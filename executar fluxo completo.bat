@echo off
title Separador de Certidoes

cd /d "%~dp0"

echo =================================
echo Executando separar_arquivos.py...
echo =================================
python separar_arquivos.py

if errorlevel 1 (
    echo.
    echo Ocorreu um erro ao executar separar_arquivos.py.
    pause
    exit /b 1
)

echo.
echo ===================================
echo Executando formatar_certidoes.py...
echo ===================================
python formatar_certidoes.py

if errorlevel 1 (
    echo.
    echo Ocorreu um erro ao executar formatar_certidoes.py.
    pause
    exit /b 1
)

echo.
echo ======================================
echo Fluxo concluido com sucesso!
echo ======================================
pause