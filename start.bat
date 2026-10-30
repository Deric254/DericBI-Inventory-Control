@echo off
title DericBI Inventory Control
cd /d "%~dp0"
echo ================================================
echo  DericBI Inventory Control
echo ================================================
echo.

:: ── Bundled Python ───────────────────────────────────────────────────────────
if not exist "python\python.exe" (
    echo [INFO] First run - setting up Python. Needs internet, takes ~2 min...
    call scripts\bundle-python.bat
    if errorlevel 1 (
        echo.
        echo [ERROR] Python setup failed.
        pause & exit /b 1
    )
    echo.
)

:: ── Node.js ───────────────────────────────────────────────────────────────────
node --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Node.js not found. Install from https://nodejs.org
    pause & exit /b 1
)

:: ── Node modules ──────────────────────────────────────────────────────────────
if not exist "node_modules\electron" (
    echo [INFO] Installing Electron - one time only...
    npm install
    if errorlevel 1 ( echo [ERROR] npm install failed. & pause & exit /b 1 )
    echo [OK] Done.
    echo.
)

:: ── Launch ────────────────────────────────────────────────────────────────────
set ICT_PYTHON=%~dp0python\python.exe
echo [OK] Python: %ICT_PYTHON%
echo [OK] Launching...
echo.

npm start
set EXIT_CODE=%errorlevel%

echo.
if %EXIT_CODE% neq 0 (
    echo [ERROR] App crashed with code %EXIT_CODE%
) else (
    echo [OK] App closed.
)
pause
