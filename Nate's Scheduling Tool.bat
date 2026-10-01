@echo off
REM Launch Pipeline Dashboard
REM Double-click this file to run the app.

:: Keep setup artifacts in the repo root, even though this launcher lives in a subfolder
set "PROJECT_ROOT=%~dp0"

pushd "%PROJECT_ROOT%"

:: Check if virtual environment exists, if not create one
if not exist ".venv" (
    call :cecho Virtual environment not found. Setting up...
    python -m venv .venv
    call :cecho Virtual environment created.
)

@REM :: Activate the virtual environment
call "%PROJECT_ROOT%\.venv\Scripts\activate.bat"

@REM :: Keep dependencies in sync every launch
@REM python -m pip install -r requirements.txt

python app/nate_scheduling_tool_gui.py

REM Keep window open if there's an error, so it doesn't just vanish
if errorlevel 1 (
    echo.
    echo Something went wrong. Press any key to close.
    pause >nul
)