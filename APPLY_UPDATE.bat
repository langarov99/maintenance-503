@echo off
title Data Extraction Bot - Обновяване
chcp 65001 > nul

set "ROOT=%~dp0"
set "ROOT=%ROOT:~0,-1%"

echo ================================================
echo   Data Extraction Bot - Прилагане на обновление
echo ================================================
echo.

:: Проверка дали има patch файл
if not exist "%ROOT%\update.patch" (
    echo [ГРЕШКА] Файлът update.patch не е намерен в:
    echo   %ROOT%\
    echo.
    echo Поставете update.patch в папката на бота и опитайте пак.
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
git -C "%ROOT%" apply --check "%ROOT%\update.patch" > nul 2>&1
if errorlevel 1 (
    echo [ГРЕШКА] Patch файлът не може да се приложи.
    echo Възможно е вече да е приложен или да има конфликт.
    echo.
    git -C "%ROOT%" apply --check "%ROOT%\update.patch"
    pause
    exit /b 1
)

:: Приложи patch-а
echo Прилагане на обновлението...
git -C "%ROOT%" apply "%ROOT%\update.patch"
if errorlevel 1 (
    echo [ГРЕШКА] Неуспешно прилагане на patch.
    pause
    exit /b 1
)

echo.
echo [OK] Обновлението е приложено успешно!
echo.

:: Изтрий patch файла след успешно прилагане
del "%ROOT%\update.patch"
echo Файлът update.patch е изтрит.
echo.

echo Можете да стартирате бота с START.bat
echo.
pause
