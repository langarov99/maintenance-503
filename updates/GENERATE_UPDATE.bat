@echo off
setlocal enabledelayedexpansion
title Data Extraction Bot - Генериране на обновление
chcp 65001 > nul

set "UPDATES_DIR=%~dp0"
set "UPDATES_DIR=%UPDATES_DIR:~0,-1%"
set "ROOT=%UPDATES_DIR%\.."
set "VERSION_FILE=%UPDATES_DIR%\version.txt"
set "GEN_SCRIPT=%UPDATES_DIR%\generate_update.py"

echo ================================================
echo   Data Extraction Bot - Генериране на обновление
echo ================================================
echo.

:: Намери Python (WinPython или системен)
set "PYTHON="
for /d %%a in ("%ROOT%\WinPython\*") do (
    for /d %%b in ("%%a\python-*") do (
        if exist "%%b\python.exe" set "PYTHON=%%b\python.exe"
    )
)
if not defined PYTHON (
    for /d %%b in ("%ROOT%\WinPython\python-*") do (
        if exist "%%b\python.exe" set "PYTHON=%%b\python.exe"
    )
)
if not defined PYTHON (
    where python > nul 2>&1
    if not errorlevel 1 set "PYTHON=python"
)
if not defined PYTHON (
    echo [ГРЕШКА] Python не е намерен. Инсталирайте WinPython в папката WinPython\
    pause
    exit /b 1
)

:: Прочети текущата версия
if not exist "%VERSION_FILE%" echo 1.0 > "%VERSION_FILE%"
set /p CURRENT_VER=<"%VERSION_FILE%"
set CURRENT_VER=%CURRENT_VER: =%
echo Текуща версия: %CURRENT_VER%

:: Раздели на major.minor
for /f "tokens=1,2 delims=." %%a in ("%CURRENT_VER%") do (
    set MAJOR=%%a
    set MINOR=%%b
)

:: Увеличи версията
set /a MINOR_NEW=%MINOR%+1
if %MINOR_NEW% GTR 9 (
    set /a MAJOR=%MAJOR%+1
    set MINOR_NEW=0
)
set NEW_VER=%MAJOR%.%MINOR_NEW%
echo Нова версия:   %NEW_VER%
echo.

:: Брой commit-и
set /p COMMITS=Брой commit-и за включване [1]:
if "%COMMITS%"=="" set COMMITS=1

:: Генерирай ZIP
set "ZIP_FILE=%UPDATES_DIR%\update_%NEW_VER%.zip"
echo.
echo Генериране на update_%NEW_VER%.zip ...
"%PYTHON%" "%GEN_SCRIPT%" "%ROOT%" "%ZIP_FILE%" %COMMITS%
if errorlevel 1 (
    echo [ГРЕШКА] Неуспешно генериране на обновлението.
    pause
    exit /b 1
)

:: Обнови version.txt
echo %NEW_VER%> "%VERSION_FILE%"

echo.
echo [OK] Файлът update_%NEW_VER%.zip е създаден в папката updates\
echo.
echo Изпратете го до другите машини и те трябва да го поставят
echo в тяхната папка updates\ и да пуснат APPLY_UPDATE.bat
echo.
pause
