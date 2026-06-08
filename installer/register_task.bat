@echo off
REM Optional helper - registers the listener as a Scheduled Task that runs
REM at logon and restarts on crash. The .exe also offers to do this from
REM its first-run dialog; this .bat is here for when you want to do it
REM manually (or re-do it after moving the .exe).
REM
REM Run this AS ADMINISTRATOR (right-click -> Run as administrator).

setlocal
set "EXE=%~dp0..\dist\CalListener.exe"
if not exist "%EXE%" set "EXE=%~dp0CalListener.exe"
if not exist "%EXE%" (
    echo Cannot find CalListener.exe.
    echo Place this .bat next to CalListener.exe and re-run.
    pause
    exit /b 1
)

echo Registering CalListener as a Scheduled Task...
schtasks /create /tn "CalListener" /sc onlogon /rl highest ^
         /tr "\"%EXE%\"" /f
if errorlevel 1 (
    echo Failed - did you run this as administrator?
    pause
    exit /b 1
)

echo.
echo Done. The listener will start at every logon.
echo Run it now from Task Scheduler ("Run") or just sign out and back in.
pause
