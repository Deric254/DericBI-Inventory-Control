@echo off
title St. Anne ICT Command Centre — Installer Builder
color 0A
cd /d "%~dp0"

echo ============================================================
echo   St. Anne Mission Hospital — ICT Command Centre
echo   Building Installer EXE
echo ============================================================
echo.

:: ── Node.js check ────────────────────────────────────────────────────────────
node --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Node.js not found. Install from https://nodejs.org
    pause & exit /b 1
)
for /f "tokens=* delims=v" %%V in ('node --version') do set NODE_VER=%%V
echo [OK] Node.js v%NODE_VER%

:: ── Bundled Python ────────────────────────────────────────────────────────────
if not exist "python\python.exe" (
    echo.
    echo [INFO] Downloading embedded Python - needs internet, ~2 min...
    call scripts\bundle-python.bat
    if errorlevel 1 ( pause & exit /b 1 )
) else (
    echo [OK] Embedded Python ready.
)

:: ── Node dependencies ─────────────────────────────────────────────────────────
if not exist "node_modules\electron" (
    echo.
    echo [INFO] Installing build tools - one time only...
    npm install
    if errorlevel 1 ( echo [ERROR] npm install failed. & pause & exit /b 1 )
)
echo [OK] Build tools ready.

:: ── Build installer ───────────────────────────────────────────────────────────
echo.
echo Building installer... (3-6 minutes, downloading Electron if first time)
echo.
npm run build:win
if errorlevel 1 (
    echo.
    echo [ERROR] Build failed - see messages above.
    pause & exit /b 1
)

echo.
echo ============================================================
echo   DONE!
echo.
echo   dist\ICT_CommandCentre_Setup.exe
echo.
echo   This is a proper installer - send it to any Windows PC.
echo   It will install to Program Files with a Start Menu shortcut.
echo   Fresh installs start with sample data; the user sets their own PIN on first launch.
echo ============================================================
echo.
pause
