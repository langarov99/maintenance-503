@echo off
title Първоначална инсталация — Data Extraction Bot
chcp 65001 > nul

set "ROOT=%~dp0"
set "ROOT=%ROOT:~0,-1%"
set "PYTHON=%ROOT%\WinPython\python-3.12.4.amd64\python.exe"

if not exist "%PYTHON%" (
    where python > nul 2>&1
    if errorlevel 1 (
        echo [ГРЕШКА] Python не е открит. Инсталирайте WinPython в папка WinPython\.
        pause & exit /b 1
    )
    set "PYTHON=python"
)

echo.
echo  ================================================
echo   Инсталация на Data Extraction Bot
echo  ================================================
echo.

:: Install Python packages
echo [1/3] Инсталиране на Python пакети...
"%PYTHON%" -m pip install -r "%ROOT%\requirements.txt"
if errorlevel 1 ( echo [ГРЕШКА] pip install неуспешно. & pause & exit /b 1 )

:: Pre-download EasyOCR models
echo.
echo [2/3] Изтегляне на EasyOCR езикови модели (bg, en, pl, cs, it, de)...
echo       (Нужен е интернет само при първо стартиране — после работи offline)
"%PYTHON%" -c "import easyocr; easyocr.Reader(['bg','en','pl','cs','it','de'], model_storage_directory='models')"
if errorlevel 1 ( echo [ПРЕДУПРЕЖДЕНИЕ] OCR моделите не бяха изтеглени напълно. )

:: Check for LLM model
echo.
echo [3/3] Проверка за LLM модел...
if exist "%ROOT%\models\phi-3-mini.gguf" (
    echo [OK] Намерен phi-3-mini.gguf — LLM режимът е активен.
) else (
    echo [INFO] LLM модел не е намерен.
    echo        За да активирате AI резервния режим:
    echo        1. Свалете phi-3-mini-4k-instruct-q4.gguf
    echo        2. Преименувайте го на phi-3-mini.gguf
    echo        3. Копирайте в папка: models\
)

:: Create done flag
echo. > "%ROOT%\.deps_installed"

echo.
echo  ================================================
echo   Инсталацията е завършена! Стартирайте START.bat
echo  ================================================
pause
