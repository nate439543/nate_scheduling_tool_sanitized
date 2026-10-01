@echo off
setlocal

:: Always run from this script's directory
set "PROJECT_ROOT=%~dp0"
pushd "%PROJECT_ROOT%"

echo ===== Nate Code Setup =====
echo.

:: Check if Python is installed
python --version > nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Python is not installed or not in PATH.
    echo Please install Python 3.8 or later from https://www.python.org/downloads/
    echo Be sure to check "Add Python to PATH" during installation.
    echo.
    echo Press any key to exit...
    pause > nul
    exit /b 1
)

:: Create virtual environment if it doesn't exist
if not exist ".venv" (
    echo Creating virtual environment...
    python -m venv .venv
)

:: Activate the virtual environment
call .venv\Scripts\activate.bat

:: Install required packages
echo Installing required packages...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo.
echo Setup completed successfully!


echo.
echo Setup completed. You can now run Nate's Scheduling Tool.bat.
echo.

:: Edit config instructions
echo ===== CONFIGURATION =====
echo.
echo To customize email settings, edit the config.ini file with your text editor.

echo Press any key to exit...
pause > nul

popd
endlocal