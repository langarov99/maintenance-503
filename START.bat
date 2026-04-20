@echo off
title Data Extraction Bot
chcp 65001 > nul

:: -------------------------------------------------------
:: Determine script location (works from USB drive too)
:: -------------------------------------------------------
set "ROOT=%~dp0"
set "ROOT=%ROOT:~0,-1%"

:: -------------------------------------------------------
:: Locate Python — prefer portable WinPython on USB
:: -------------------------------------------------------
set "PYTHON=%ROOT%\WinPython\python-3.12.4.amd64\python.exe"
if not exist "%PYTHON%" (
    :: Fallback to system Python
    where python > nul 2>&1
    if errorlevel 1 (
        echo [ГРЕШКА] Python не е намерен!
        echo Моля, копирайте WinPython в папка WinPython\ на флашката.
        pause
        exit /b 1
    )
    set "PYTHON=python"
)

:: -------------------------------------------------------
:: Install dependencies if not already installed
:: -------------------------------------------------------
set "FLAG=%ROOT%\.deps_installed"
if not exist "%FLAG%" (
    echo Инсталиране на зависимости (само при първо стартиране)...
    "%PYTHON%" -m pip install -r "%ROOT%\requirements.txt" --quiet
    if errorlevel 1 (
        echo [ГРЕШКА] Не успя инсталацията на зависимости.
        pause
        exit /b 1
    )
    echo. > "%FLAG%"
    echo Готово!
)

:: -------------------------------------------------------
:: Start the application
:: -------------------------------------------------------
echo.
echo  ================================================
echo   Data Extraction Bot  ^|  http://localhost:5000
echo  ================================================
echo   Затворете този прозорец за да спрете бота.
echo.

start "" "http://localhost:5000"
"%PYTHON%" -m uvicorn app.main:app --host 127.0.0.1 --port 5000 --app-dir "%ROOT%"

pause
