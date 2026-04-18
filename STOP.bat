@echo off
title Спиране на Data Extraction Bot
chcp 65001 > nul
echo Спиране на Data Extraction Bot...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":5000"') do (
    taskkill /f /pid %%a > nul 2>&1
)
echo Ботът е спрян.
timeout /t 2 > nul
