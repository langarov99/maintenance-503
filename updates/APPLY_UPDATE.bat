@echo off
setlocal enabledelayedexpansion
title Data Extraction Bot - Прилагане на обновление
chcp 65001 > nul

set "UPDATES_DIR=%~dp0"
set "UPDATES_DIR=%UPDATES_DIR:~0,-1%"
set "ROOT=%UPDATES_DIR%\.."
set "VERSION_FILE=%UPDATES_DIR%\version.txt"

echo ================================================
echo   Data Extraction Bot - Прилагане на обновление
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
echo Инсталирана версия: %CURRENT_VER%
echo.

:: Намери patch файл (update_X.Y.patch)
set "PATCH_FILE="
set "PATCH_NAME="
for /f "delims=" %%f in ('dir /b /a-d "%UPDATES_DIR%\update_*.patch" 2^>nul') do (
    set "PATCH_FILE=%UPDATES_DIR%\%%f"
    set "PATCH_NAME=%%~nf"
)

if "%PATCH_FILE%"=="" (
    echo [ГРЕШКА] Няма намерен patch файл в:
    echo   %UPDATES_DIR%\
    echo.
    echo Поставете update_X.Y.patch в папката updates\ и опитайте пак.
    pause
    exit /b 1
)

:: Извлечи версията от името на файла (update_1.1 -> 1.1)
set "PATCH_VER=%PATCH_NAME:update_=%"
echo Намерен patch: %PATCH_NAME%.patch  ^(версия %PATCH_VER%^)
echo.

:: Провери дали може да се приложи
"%GIT%" -C "%ROOT%" apply --check "%PATCH_FILE%" > nul 2>&1
if errorlevel 1 (
    echo [ГРЕШКА] Patch файлът не може да се приложи.
    echo Възможно е вече да е приложен или да има конфликт.
    echo.
    "%GIT%" -C "%ROOT%" apply --check "%PATCH_FILE%"
    pause
    exit /b 1
)

:: Приложи patch-а
echo Прилагане на обновлението...
"%GIT%" -C "%ROOT%" apply "%PATCH_FILE%"
if errorlevel 1 (
    echo [ГРЕШКА] Неуспешно прилагане на patch.
    pause
    exit /b 1
)

:: Обнови version.txt
echo %PATCH_VER%> "%VERSION_FILE%"

:: Изтрий patch файла
del "%PATCH_FILE%"

echo.
echo [OK] Обновлението е приложено успешно!
echo      Версия: %CURRENT_VER% --^> %PATCH_VER%
echo.
echo Стартирайте бота с START.bat
echo.
pause
