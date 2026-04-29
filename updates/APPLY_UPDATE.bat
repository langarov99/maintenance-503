@echo off
title Data Extraction Bot - Обновяване
chcp 65001 > nul

:: Тази папка (updates\)
set "UPDATES_DIR=%~dp0"
set "UPDATES_DIR=%UPDATES_DIR:~0,-1%"

:: Главната папка на бота (една ниво нагоре)
set "ROOT=%UPDATES_DIR%\.."

echo ================================================
echo   Data Extraction Bot - Прилагане на обновление
echo ================================================
echo.

:: Проверка дали има patch файл
if not exist "%UPDATES_DIR%\update.patch" (
    echo [ГРЕШКА] Файлът update.patch не е намерен в:
    echo   %UPDATES_DIR%\
    echo.
    echo Поставете update.patch в папката updates\ и опитайте пак.
    pause
    exit /b 1
)

:: Намери git
where git > nul 2>&1
if errorlevel 1 (
    echo [ГРЕШКА] Git не е намерен! Трябва да е инсталиран git.
    pause
    exit /b 1
)

echo Намерен patch файл: update.patch
echo.

:: Провери дали patch-ът може да се приложи
git -C "%ROOT%" apply --check "%UPDATES_DIR%\update.patch" > nul 2>&1
if errorlevel 1 (
    echo [ГРЕШКА] Patch файлът не може да се приложи.
    echo Възможно е вече да е приложен или да има конфликт.
    echo.
    git -C "%ROOT%" apply --check "%UPDATES_DIR%\update.patch"
    pause
    exit /b 1
)

:: Приложи patch-а
echo Прилагане на обновлението...
git -C "%ROOT%" apply "%UPDATES_DIR%\update.patch"
if errorlevel 1 (
    echo [ГРЕШКА] Неуспешно прилагане на patch.
    pause
    exit /b 1
)

echo.
echo [OK] Обновлението е приложено успешно!
echo.

:: Изтрий patch файла след успешно прилагане
del "%UPDATES_DIR%\update.patch"
echo Файлът update.patch е изтрит.
echo.

echo Можете да стартирате бота с START.bat
echo.
pause
