@echo off
title Data Extraction Bot
chcp 65001 > nul

set "ROOT=%~dp0"
set "ROOT=%ROOT:~0,-1%"

:: Find python.exe inside WinPython (any subfolder depth)
set "PYTHON="
for /d %%a in ("%ROOT%\WinPython\*") do (
    for /d %%b in ("%%a\python-*") do (
        if exist "%%b\python.exe" set "PYTHON=%%b\python.exe"
    )
)
:: Also check one level directly (WinPython\python-x.x.x)
if not defined PYTHON (
    for /d %%b in ("%ROOT%\WinPython\python-*") do (
        if exist "%%b\python.exe" set "PYTHON=%%b\python.exe"
    )
)
:: Fallback to system Python
if not defined PYTHON (
    where python > nul 2>&1
    if not errorlevel 1 set "PYTHON=python"
)
if not defined PYTHON (
    echo [ГРЕШКА] Python не е намерен!
    echo Моля, копирайте WinPython в папка WinPython\ на флашката.
    pause
    exit /b 1
)
echo [OK] Python: %PYTHON%

:: Install dependencies if not already installed
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

echo.
echo  ================================================
echo   Data Extraction Bot  ^|  http://localhost:5000
echo  ================================================
echo   Затворете този прозорец за да спрете бота.
echo.

start "" "http://localhost:5000"
"%PYTHON%" -m uvicorn app.main:app --host 127.0.0.1 --port 5000 --app-dir "%ROOT%"

pause
