@echo off
title Nomear Certidoes via OCR (PaddleOCR + OpenCV)

cd /d "%~dp0"

python nomear_certidoes_ocr.py

if errorlevel 1 (
    echo.
    echo Ocorreu um erro durante a execucao.
)

echo.
pause
