@echo off
title Configurando Separador de Certidoes

echo ===============================
echo Configurando projeto...
echo ===============================
echo.

:: Verifica se o Python existe
python --version >nul 2>&1

if errorlevel 1 (
    echo Python nao encontrado.
    echo Instale o Python e marque a opcao "Add Python to PATH".
    pause
    exit /b
)

echo Python encontrado.
echo.

echo Instalando dependencias...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo.

:: Cria as pastas caso nao existam
if not exist entrada mkdir entrada
if not exist saida mkdir saida
if not exist __pycache__ mkdir __pycache__

:: Cria os .gitkeep se nao existirem
if not exist entrada\.gitkeep type nul > entrada\.gitkeep
if not exist saida\.gitkeep type nul > saida\.gitkeep

:: Oculta os arquivos
attrib +h entrada\.gitkeep
attrib +h saida\.gitkeep

:: Oculta a pasta __pycache__
attrib +h __pycache__

echo.
echo ===================================
echo Configuracao concluida com sucesso!
echo ===================================
echo.
pause