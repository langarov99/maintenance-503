@echo off
setlocal enabledelayedexpansion
title Data Extraction Bot - Прилагане на обновление
chcp 65001 > nul

set "UPDATES_DIR=%~dp0"
set "UPDATES_DIR=%UPDATES_DIR:~0,-1%"
set "ROOT=%UPDATES_DIR%\.."
set "VERSION_FILE=%UPDATES_DIR%\version.txt"
set "APPLY_SCRIPT=%UPDATES_DIR%\apply_update.py"

echo ================================================
echo   Data Extraction Bot - Прилагане на обновление
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
    echo [ГРЕШКА] Python не е намерен. Стартирайте бота поне веднъж с START.bat
    pause
    exit /b 1
)

:: Прочети текущата версия
if not exist "%VERSION_FILE%" echo 1.0 > "%VERSION_FILE%"
set /p CURRENT_VER=<"%VERSION_FILE%"
set CURRENT_VER=%CURRENT_VER: =%
echo Инсталирана версия: %CURRENT_VER%
echo.

:: Намери ZIP файл
set "ZIP_FILE="
set "ZIP_NAME="
for /f "delims=" %%f in ('dir /b /a-d "%UPDATES_DIR%\update_*.zip" 2^>nul') do (
    set "ZIP_FILE=%UPDATES_DIR%\%%f"
    set "ZIP_NAME=%%~nf"
)

if "%ZIP_FILE%"=="" (
    echo [ГРЕШКА] Няма намерен update файл в:
    echo   %UPDATES_DIR%\
    echo.
    echo Поставете update_X.Y.zip в папката updates\ и опитайте пак.
    pause
    exit /b 1
)

:: Извлечи версията
set "NEW_VER=%ZIP_NAME:update_=%"
echo Намерен update: %ZIP_NAME%.zip  ^(версия %NEW_VER%^)
echo.

:: Приложи update-а
echo Прилагане на обновлението...
echo.
"%PYTHON%" "%APPLY_SCRIPT%" "%ZIP_FILE%" "%ROOT%"
if errorlevel 1 (
    echo.
    echo [ГРЕШКА] Неуспешно прилагане на обновлението.
    pause
    exit /b 1
)

:: Обнови version.txt
echo %NEW_VER%> "%VERSION_FILE%"

:: Изтрий ZIP файла
del "%ZIP_FILE%"

echo.
echo [OK] Обновлението е приложено успешно!
echo      Версия: %CURRENT_VER% --^> %NEW_VER%
echo.
echo Стартирайте бота с START.bat
echo.
pause
