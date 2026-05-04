@echo off
setlocal enabledelayedexpansion
title Data Extraction Bot - Генериране на обновление
chcp 65001 > nul

set "UPDATES_DIR=%~dp0"
set "UPDATES_DIR=%UPDATES_DIR:~0,-1%"
set "ROOT=%UPDATES_DIR%\.."
set "VERSION_FILE=%UPDATES_DIR%\version.txt"
set "GEN_SCRIPT=%UPDATES_DIR%\generate_update.py"
set "LAST_COMMIT_FILE=%UPDATES_DIR%\last_exported_commit.txt"

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
    echo [ГРЕШКА] Python не е намерен.
    pause
    exit /b 1
)

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

:: Определи from_ref автоматично от последния export
set "FROM_REF="
if exist "%LAST_COMMIT_FILE%" (
    set /p LAST_COMMIT=<"%LAST_COMMIT_FILE%"
    set "LAST_COMMIT=!LAST_COMMIT: =!"

    for /f %%c in ('"%GIT%" -C "%ROOT%" rev-list --count !LAST_COMMIT!..HEAD 2^>nul') do set "COMMIT_COUNT=%%c"

    if "!COMMIT_COUNT!"=="0" (
        echo [INFO] Няма нови commit-и от последния update.
        echo.
        pause
        exit /b 0
    )

    echo Нови commit-и от последния update ^(!COMMIT_COUNT! бр.^):
    "%GIT%" -C "%ROOT%" log --oneline !LAST_COMMIT!..HEAD
    echo.
    set "FROM_REF=!LAST_COMMIT!"
) else (
    echo [INFO] Няма запис за предишен update. Последни commit-и:
    "%GIT%" -C "%ROOT%" log --oneline -10
    echo.
    set /p USER_COUNT=Брой commit-и за включване [1]:
    if "!USER_COUNT!"=="" set USER_COUNT=1
    set "FROM_REF=HEAD~!USER_COUNT!"
)

set /p CONFIRM=Натиснете Enter за потвърждение или Ctrl+C за отказ...

:: Генерирай ZIP
set "ZIP_FILE=%UPDATES_DIR%\update_%NEW_VER%.zip"
echo.
echo Генериране на update_%NEW_VER%.zip ...
"%PYTHON%" "%GEN_SCRIPT%" "%ROOT%" "%ZIP_FILE%" "!FROM_REF!"
if errorlevel 1 (
    echo [ГРЕШКА] Неуспешно генериране на обновлението.
    pause
    exit /b 1
)

:: Запази текущия HEAD hash за следващия update
for /f %%h in ('"%GIT%" -C "%ROOT%" rev-parse HEAD') do echo %%h> "%LAST_COMMIT_FILE%"

:: Обнови version.txt
echo %NEW_VER%> "%VERSION_FILE%"

echo.
echo [OK] Файлът update_%NEW_VER%.zip е създаден в папката updates\
echo.
echo Изпратете го до другите машини и те трябва да го поставят
echo в тяхната папка updates\ и да пуснат APPLY_UPDATE.bat
echo.
pause
