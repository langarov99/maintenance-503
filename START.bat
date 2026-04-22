@echo off
title Data Extraction Bot
chcp 65001 > nul

set "ROOT=%~dp0"
set "ROOT=%ROOT:~0,-1%"

:: Find python.exe inside WinPython subfolders
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
    echo [GRESKA] Python ne e nameren!
    pause
    exit /b 1
)
echo [OK] Python: %PYTHON%

:: Install dependencies if not already done
set "FLAG=%ROOT%\.deps_installed"
if not exist "%FLAG%" (
    echo Instalirane na zavisimosti...
    "%PYTHON%" -m pip install -r "%ROOT%\requirements.txt" --quiet
    if errorlevel 1 (
        echo [GRESKA] pip install ne uspya.
        pause
        exit /b 1
    )
    echo done > "%FLAG%"
)

echo.
echo  ================================================
echo   Data Extraction Bot  -  http://localhost:5000
echo  ================================================
echo   Zatvorete prozoreca "Bot Server" za da spirete.
echo.

:: Start server in a separate minimized window
start "Bot Server" /min "%COMSPEC%" /k "%PYTHON%" -m uvicorn app.main:app --host 127.0.0.1 --port 5000 --app-dir "%ROOT%"

:: Wait until server responds (max ~15s, checks every 0.5s)
echo Izchakване na sarvara...
:wait_loop
timeout /t 1 /nobreak > nul
"%PYTHON%" -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:5000/', timeout=1)" 2>nul
if errorlevel 1 goto wait_loop

:: Server is ready — open browser
echo [OK] Sarvarot e gotov!
start "" "http://localhost:5000"

echo.
echo   Brauzyrat e otvoren. Zatvorete "Bot Server" za da spirete bota.
echo.
pause
