@echo off
echo SharePoint Excel Manager - Windows Runner
echo ========================================

:: Check if virtual environment exists
if not exist "venv\Scripts\activate.bat" (
    echo ERROR: Virtual environment not found
    echo Please run install_windows.bat first
    pause
    exit /b 1
)

:: Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat

:: Check if package is installed
python -c "import sharepoint_excel_manager" >nul 2>&1
if errorlevel 1 (
    echo ERROR: SharePoint Excel Manager not installed
    echo Please run install_windows.bat first
    pause
    exit /b 1
)

:: Run the application
echo Starting SharePoint Excel Manager...
echo.
python -m sharepoint_excel_manager.main

:: Keep window open if there was an error
if errorlevel 1 (
    echo.
    echo Application exited with an error
    pause
)