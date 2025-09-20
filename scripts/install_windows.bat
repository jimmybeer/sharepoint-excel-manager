@echo off
echo SharePoint Excel Manager - Windows Installation Script
echo ====================================================

:: Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or later from https://python.org
    pause
    exit /b 1
)

:: Display Python version
echo Python version:
python --version

:: Check if we're in the project directory
if not exist "pyproject.toml" (
    echo ERROR: Please run this script from the project root directory
    echo Expected to find pyproject.toml in current directory
    pause
    exit /b 1
)

:: Create virtual environment
echo.
echo Creating virtual environment...
python -m venv venv
if errorlevel 1 (
    echo ERROR: Failed to create virtual environment
    pause
    exit /b 1
)

:: Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat

:: Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip

:: Install the project in development mode
echo.
echo Installing project dependencies...
pip install -e .
if errorlevel 1 (
    echo ERROR: Failed to install project dependencies
    pause
    exit /b 1
)

:: Install development dependencies
echo.
echo Installing development dependencies...
pip install -e ".[dev]"

echo.
echo ====================================================
echo Installation completed successfully!
echo.
echo To run the application:
echo   1. Activate the virtual environment: venv\Scripts\activate.bat
echo   2. Run the application: python -m sharepoint_excel_manager.main
echo.
echo Or use the run script: scripts\run_windows.bat
echo ====================================================
pause