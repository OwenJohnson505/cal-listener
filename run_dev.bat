@echo off
REM ============================================================
REM  Cal Listener - DEV mode (run from source).
REM
REM  Skips the PyInstaller build + GitHub release cycle entirely.
REM  First run: clones the repo and installs deps.
REM  Every subsequent run: git pull + start listener.
REM
REM  Iteration loop is now:
REM    1. Claude pushes a fix to main.
REM    2. You hit Ctrl+C in the listener window to stop it.
REM    3. You double-click this .bat again - it pulls latest + restarts.
REM  No downloads, no replacing .exe.
REM ============================================================

setlocal
set REPO=C:\Users\jowen\cal-listener-dev
set PY=py
title Cal Listener (dev mode)

REM Make sure Python is available.
where %PY% >nul 2>nul
if errorlevel 1 (
    set PY=python
    where %PY% >nul 2>nul
    if errorlevel 1 (
        echo.
        echo Python isn't on PATH. Install Python 3.12+ from
        echo https://www.python.org/downloads/  and re-run this script.
        pause
        exit /b 1
    )
)

REM First-time setup.
if not exist "%REPO%\.git" (
    echo.
    echo === First run - cloning cal-listener into %REPO% ===
    git clone https://github.com/OwenJohnson505/cal-listener "%REPO%"
    if errorlevel 1 goto fail
    cd /d "%REPO%"
    echo.
    echo === Installing dependencies (one-time) ===
    %PY% -m pip install --upgrade pip
    %PY% -m pip install -e .[windows]
    if errorlevel 1 goto fail
    echo.
    echo === Setup complete ===
    echo.
)

cd /d "%REPO%"

echo.
echo === Pulling latest from origin/main ===
git pull --ff-only
if errorlevel 1 (
    echo.
    echo git pull failed - check above. Continuing with whatever's locally checked out.
)

echo.
echo === Starting listener (Ctrl+C in this window stops it) ===
echo.
%PY% -m cal_listener

echo.
echo Listener exited. Press any key to close.
pause >nul
goto end

:fail
echo.
echo Setup failed - see error above.
pause

:end
endlocal
