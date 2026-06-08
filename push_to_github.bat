@echo off
REM ----------------------------------------------------------------------
REM Push the cal-listener repo to a new GitHub repo, then watch
REM GitHub Actions build the .exe.
REM
REM Requirements:
REM   - You're signed into GitHub in your Windows credential manager
REM     (this is already the case since cal-toolkit-web pushes worked).
REM   - You've created an EMPTY repo at github.com/OwenJohnson505/cal-listener
REM     (no README, no .gitignore, no licence — just empty).
REM
REM Run this once.
REM ----------------------------------------------------------------------

setlocal
cd /d "%~dp0"
title cal-listener - first push

set "GIT=C:\Program Files\Git\mingw64\bin\git.exe"
if not exist "%GIT%" set "GIT=C:\Program Files\Git\bin\git.exe"
if not exist "%GIT%" set "GIT=git"

echo.
echo === cal-listener: first push ===
echo.

"%GIT%" init -b main
if errorlevel 1 goto fail

"%GIT%" add -A
"%GIT%" commit -m "Initial commit: lean job-queue worker for cal-toolkit-web"
if errorlevel 1 goto fail

"%GIT%" remote remove origin >nul 2>&1
"%GIT%" remote add origin https://github.com/OwenJohnson505/cal-listener.git
"%GIT%" push -u origin main
if errorlevel 1 goto fail

echo.
echo ============================================
echo  Pushed to https://github.com/OwenJohnson505/cal-listener
echo.
echo  Next:
echo    1. Open the Actions tab on the repo.
echo    2. Wait ~2 minutes for the "Build CalListener.exe" workflow
echo       to finish.
echo    3. On the run's summary page, download the
echo       "CalListener-<sha>" artifact.
echo    4. Inside, you'll find CalListener.exe.
echo    5. Double-click that .exe on your laptop and follow the
echo       first-run dialog.
echo ============================================
goto end

:fail
echo.
echo PUSH FAILED. Possible causes:
echo   - You haven't created the empty repo on github.com yet.
echo   - Git credentials aren't configured.
echo   - The repo already has content (delete it on github.com and retry).
:end
pause
endlocal
