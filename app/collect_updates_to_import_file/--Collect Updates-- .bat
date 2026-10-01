@echo off
call :cecho ===== --Collect Updates Into Import Sheet-- =====
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
call :cecho Enter the OLD data date for the project from which you would like to collect update data (MMDDYY format, e.g., 091525):
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

:RUN_SCRIPT
call :cecho Collecting data from data date %DATE_INPUT% 
python "%SCRIPT_DIR%collect_updates_to_import_file.py" --old_data_date %DATE_INPUT% 

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