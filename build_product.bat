@echo off
setlocal
cd /d "%~dp0"

echo [1/5] Checking Python...
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

echo [2/5] Creating virtual environment (if needed)...
if not exist ".venv\Scripts\python.exe" (
    %PY_CMD% -m venv .venv
    if %errorlevel% neq 0 (
        echo Failed to create virtual environment.
        pause
        exit /b 1
    )
)

echo [3/5] Installing build dependencies...
call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller pywin32
if %errorlevel% neq 0 (
    echo Failed to install build dependencies.
    pause
    exit /b 1
)

echo [4/5] Building executable...
pyinstaller --noconfirm --clean --windowed --name "ProSlideViewer" viewer.py
if %errorlevel% neq 0 (
    echo Build failed.
    pause
    exit /b 1
)

echo [5/5] Build complete.
echo EXE path: dist\ProSlideViewer\ProSlideViewer.exe
pause
exit /b 0
