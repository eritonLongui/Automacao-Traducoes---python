@echo off
title Separador de Certidoes

cd /d "%~dp0"

set "TEM_ARQUIVOS=0"
if exist "entrada" (
    for /r "entrada" %%f in (*) do (
        if /i not "%%~nxf"==".gitkeep" set "TEM_ARQUIVOS=1"
    )
)

if "%TEM_ARQUIVOS%"=="1" (
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
) else (
    echo ========================================================
    echo Nenhum documento na pasta entrada. Pulando separacao...
    echo ========================================================
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