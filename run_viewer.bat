@echo off
setlocal
cd /d "%~dp0"

echo [1/4] Checking Python...
where py >nul 2>nul
if %errorlevel%==0 (
    set "PY_CMD=py -3"
) else (
    where python >nul 2>nul
    if %errorlevel% neq 0 (
        echo Python not found. Install Python 3.10+ and try again.
        pause
        exit /b 1
    )
    set "PY_CMD=python"
)

echo [2/4] Creating virtual environment (if needed)...
if not exist ".venv\Scripts\python.exe" (
    %PY_CMD% -m venv .venv
    if %errorlevel% neq 0 (
        echo Failed to create virtual environment.
        pause
        exit /b 1
    )
)

echo [3/4] Installing dependencies...
call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo Failed to install required packages.
    pause
    exit /b 1
)

REM Optional: PPTX support (PowerPoint + pywin32)
python -m pip install pywin32 >nul 2>nul

echo [4/4] Starting viewer...
python viewer.py
set "APP_EXIT=%errorlevel%"

if not "%APP_EXIT%"=="0" (
    echo.
    echo Viewer closed with exit code %APP_EXIT%.
    echo If PPTX mode fails, ensure Microsoft PowerPoint is installed.
)

pause
exit /b %APP_EXIT%
