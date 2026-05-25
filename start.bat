@echo off
setlocal

REM =====================================================
REM  Python Virtual Environment Setup Script
REM =====================================================

REM Get the directory where this .bat file is located
set "PROJECT_DIR=%~dp0"

echo ==========================================
echo Setting up Python virtual environment...
echo ==========================================

REM Change to the project directory
cd /d "%PROJECT_DIR%"

REM Check if Python is installed and available in PATH
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not added to PATH.
    pause
    exit /b 1
)

REM Create virtual environment if it does not exist
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
) else (
    echo Virtual environment already exists.
)

REM Activate the virtual environment
call venv\Scripts\activate.bat

REM Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip

REM Install dependencies from requirements.txt
if exist "requirements.txt" (
    echo Installing dependencies from requirements.txt...
    pip install -r requirements.txt
) else (
    echo WARNING: requirements.txt was not found.
)

echo ==========================================
echo Environment setup completed successfully.
echo ==========================================

pause