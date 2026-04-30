@echo off
setlocal enabledelayedexpansion
title Data Extraction Bot - Генериране на обновление
chcp 65001 > nul

set "UPDATES_DIR=%~dp0"
set "UPDATES_DIR=%UPDATES_DIR:~0,-1%"
set "ROOT=%UPDATES_DIR%\.."
set "VERSION_FILE=%UPDATES_DIR%\version.txt"

echo ================================================
echo   Data Extraction Bot - Генериране на обновление
echo ================================================
echo.

:: Намери git
set "GIT=git"
where git > nul 2>&1
if errorlevel 1 (
    set "GIT="
    for %%P in (
        "C:\Program Files\Git\cmd\git.exe"
        "C:\Program Files\Git\bin\git.exe"
        "C:\Program Files (x86)\Git\cmd\git.exe"
        "C:\Program Files (x86)\Git\bin\git.exe"
    ) do (
        if "!GIT!"=="" if exist %%P set "GIT=%%~P"
    )
    if "!GIT!"=="" (
        echo [ГРЕШКА] Git не е намерен!
        echo Инсталирайте Git от https://git-scm.com/download/win
        pause
        exit /b 1
    )
    echo [INFO] Git намерен: !GIT!
    echo.
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

:: Генерирай patch
set "PATCH_FILE=%UPDATES_DIR%\update_%NEW_VER%.patch"
echo.
echo Генериране на update_%NEW_VER%.patch ...
"%GIT%" -C "%ROOT%" format-patch HEAD~%COMMITS% --stdout > "%PATCH_FILE%"
if errorlevel 1 (
    echo [ГРЕШКА] Неуспешно генериране на patch.
    pause
    exit /b 1
)

:: Обнови version.txt
echo %NEW_VER%> "%VERSION_FILE%"

echo.
echo [OK] Файлът update_%NEW_VER%.patch е създаден в папката updates\
echo.
echo Изпратете го до другите машини и те трябва да го поставят
echo в тяхната папка updates\ и да пуснат APPLY_UPDATE.bat
echo.
pause
