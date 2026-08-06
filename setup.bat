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

:: Oculta pastas internas/de desenvolvimento
if exist __pycache__ attrib +h __pycache__
if exist modelos attrib +h modelos
if exist .git attrib +h .git

:: Garante visibilidade das pastas de uso principal
if exist entrada attrib -h entrada
if exist divididos attrib -h divididos
if exist saida attrib -h saida

:: Oculta arquivos .gitkeep especificos
if exist entrada\.gitkeep attrib +h entrada\.gitkeep
if exist saida\.gitkeep attrib +h saida\.gitkeep

:: Oculta todos os arquivos no diretorio raiz que NAO sao .bat
for %%f in (*.*) do (
    if /i not "%%~xf"==".bat" (
        attrib +h "%%f"
    ) else (
        attrib -h "%%f"
    )
)

echo.
echo ===================================
echo Configuracao concluida com sucesso!
echo ===================================
echo.
pause