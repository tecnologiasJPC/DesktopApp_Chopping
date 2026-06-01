@echo off
setlocal EnableDelayedExpansion

:: Navigate to the repository root where this script lives
cd /d "%~dp0"

if not exist requirements.txt (
    echo ERROR: requirements.txt not found in %CD%
    exit /b 1
)

:: Ensure Python is available on PATH
where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not available on PATH. Install Python and try again.
    exit /b 1
)

:: Create virtual environment if it does not exist
if exist ".venv\Scripts\python.exe" (
    echo Virtual environment already exists.
) else (
    echo Creating virtual environment in .venv...
    python -m venv .venv
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment.
        exit /b 1
    )
)

:: Activate virtual environment
call ".venv\Scripts\activate"
if errorlevel 1 (
    echo ERROR: Failed to activate virtual environment.
    exit /b 1
)

echo Installing dependencies from requirements.txt...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install dependencies from requirements.txt.
    exit /b 1
)

echo.
echo Setup complete.
echo To use the virtual environment, run:
echo     .venv\Scripts\activate
endlocal
pause