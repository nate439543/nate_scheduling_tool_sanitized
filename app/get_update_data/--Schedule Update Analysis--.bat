@echo off
call :cecho ===== --Update Analysis Data Transfer-- =====
echo.

:: Keep setup artifacts in the repo root, even though this launcher lives in a subfolder
set "PROJECT_ROOT=%~dp0.."
set "SCRIPT_DIR=%~dp0"

pushd "%PROJECT_ROOT%"

:: Check if virtual environment exists, if not create one
if not exist ".venv" (
    call :cecho Virtual environment not found. Setting up...
    python -m venv .venv
    call "%PROJECT_ROOT%\.venv\Scripts\activate.bat"
    pip install -r requirements.txt
    call :cecho Setup complete!
    echo.
) else (
    :: Activate the virtual environment
    call "%PROJECT_ROOT%\.venv\Scripts\activate.bat"
)

popd



:DATE_PROMPT
call :cecho Enter the Update's OLD Data Date (MMDDYY format, e.g., 091525):
set /p DATE_INPUT=

if "%DATE_INPUT%"=="" (
    call :cecho You gotta enter a date, dumb-dumb
    goto :DATE_PROMPT
) else if "%DATE_INPUT:~5,1%"=="" (
    call :cecho Date must be exactly 6 characters in MMDDYY format.
    goto :DATE_PROMPT
) else if not "%DATE_INPUT:~6,1%"=="" (
    call :cecho Date must be exactly 6 characters in MMDDYY format.
    goto :DATE_PROMPT
) else (
    call :cecho Proceeding...
)


call :cecho Getting schedule data from schedules with old data date %DATE_INPUT%
python "%SCRIPT_DIR%get_update_data_main.py" --old_data_date %DATE_INPUT% 

call :end_process

goto :eof

:end_process
echo.
call :cecho Process completed.

:: Deactivate the virtual environment
call "%PROJECT_ROOT%\.venv\Scripts\deactivate.bat"

echo.
call :cecho Press any key to exit...
pause > nul

goto :eof



:: Prints a line in a random color using PowerShell
:cecho
setlocal EnableDelayedExpansion
set "TEXT=%*"
:: Pick a random non-Black console color inside PowerShell and print the message from env var to avoid quoting issues
set "CECHO_TEXT=!TEXT!"
powershell -NoProfile -Command "$m=$env:CECHO_TEXT; $colors=[enum]::GetNames([System.ConsoleColor]) | Where-Object { $_ -ne 'Black' }; $f=$colors | Get-Random; Write-Host -ForegroundColor $f $m"
endlocal & goto :eof