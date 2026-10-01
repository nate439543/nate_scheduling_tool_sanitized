@echo off
call :cecho ===== Rename Files =====
echo.

:: Change to the script directory
cd %~dp0

:: Check if virtual environment exists, if not create one
if not exist ".venv" (
    call :cecho Virtual environment not found. Setting up...
    python -m venv .venv
    call .venv\Scripts\activate.bat
    pip install -r requirements.txt
    call :cecho Setup complete!
    echo.
) else (
    :: Activate the virtual environment
    call .venv\Scripts\activate.bat
)

:DATE_PROMPT
call :cecho Enter date (MMDDYY format, e.g., 091525) or press Enter for 090525 (testing date):
set /p DATE_INPUT=




call :cecho ------------------------------------
call :cecho ------------------------------------
call :cecho ------------------------------------
call :cecho ------------------------------------
call :cecho ------------------------------------
call :cecho Processing...

if "%DATE_INPUT%"=="" (
    python rename_to_onedrive_main.py --date "090525" 
) else (
    python rename_to_onedrive_main.py --date %DATE_INPUT%
)


echo.
call :cecho Process completed.

:: Deactivate the virtual environment
call .venv\Scripts\deactivate.bat

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