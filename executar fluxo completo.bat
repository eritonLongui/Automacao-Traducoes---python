@echo off
title Fluxo Completo de Traducoes

cd /d "%~dp0"

echo ======================================
echo Executando solicitar_mauri.py...
echo ======================================
python solicitar_mauri.py

if errorlevel 1 (
    echo.
    echo Ocorreu um erro ao executar solicitar_mauri.py.
    echo O fluxo completo foi interrompido para correcao.
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
echo Limpando a pasta de entrada...
echo ======================================
for /d %%p in ("entrada\*") do rmdir /s /q "%%p"
for %%f in ("entrada\*") do (
    if /i not "%%~nxf"==".gitkeep" del /q "%%f"
)

echo.
echo =======================================================
echo Atualizando data das certidoes para a data atual...
echo =======================================================
python atualizar_data_traducoes.py

if errorlevel 1 (
    echo.
    echo Ocorreu um erro ao atualizar a data das certidoes.
    pause
    exit /b 1
)

echo.
echo ======================================
echo Executando numerar_traducoes.py...
echo ======================================
python numerar_traducoes.py

if errorlevel 1 (
    echo.
    echo Ocorreu um erro ao executar numerar_traducoes.py.
    pause
    exit /b 1
)

echo.
echo ======================================
echo Fluxo completo concluido com sucesso!
echo ======================================
pause