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
echo   Zatvorete tozi prozorec za da spirete bota.
echo.

:: Open browser in background: polls every 0.5s until server responds (max 20s)
start /b "" "%PYTHON%" -c "exec('import urllib.request,time,webbrowser\nfor _ in range(40):\n  time.sleep(0.5)\n  try:\n    urllib.request.urlopen(\'http://127.0.0.1:5000/\',timeout=1)\n    webbrowser.open(\'http://localhost:5000\')\n    break\n  except: pass')"

:: Kill any lingering process on port 5000 before starting
for /f "tokens=5" %%a in ('netstat -aon 2^>nul ^| findstr ":5000 "') do (
    taskkill /F /PID %%a >nul 2>&1
)

:: Start server in this window (foreground — log visible here)
"%PYTHON%" -m uvicorn app.main:app --host 127.0.0.1 --port 5000 --app-dir "%ROOT%"

pause
