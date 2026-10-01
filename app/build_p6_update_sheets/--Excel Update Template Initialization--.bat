@echo off
call :cecho ===== --Excel Update Template Initialization-- =====
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
call :cecho Enter the Current Data Date (Date should NOT be in the future) (MMDDYY format, e.g., 091525):
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
) 

for /f %%i in ('python "%PROJECT_ROOT%\utilities\date_calculations.py" --old_data_date %DATE_INPUT%') do set "DEFAULT_NEW_DATA_DATE=%%i"
call :cecho Default next data date: %DEFAULT_NEW_DATA_DATE%

:DEFAULT_NEW_DD_PROMPT

call :cecho Would you like to use the default next data date of %DEFAULT_NEW_DATA_DATE%? (Y/N, default Y):
set /p USE_DEFAULT_NEW_DATA_DATE_INPUT=

if /I "%USE_DEFAULT_NEW_DATA_DATE_INPUT%"=="Y" (
    set NEW_DATA_DATE=%DEFAULT_NEW_DATA_DATE%
    goto :BUILD_SHEETS
) else if /I "%USE_DEFAULT_NEW_DATA_DATE_INPUT%"=="N" (
    goto :MONTHLY_BOOL_PROMPT

    
) else if "%USE_DEFAULT_NEW_DATA_DATE_INPUT%"=="" (
    set NEW_DATA_DATE=%DEFAULT_NEW_DATA_DATE%
    goto :BUILD_SHEETS
) else (
    call :cecho Please enter Y or N for the default next data date question.
    goto :DEFAULT_NEW_DD_PROMPT
)


call :cecho ------------------------------------

:MONTHLY_BOOL_PROMPT
    call :cecho is this a monthly update? (Y/N, default N):
    set /p MONTHLY_BOOL_INPUT=


if /I "%MONTHLY_BOOL_INPUT%"=="Y" (
    set MONTHLY_BOOL=true

) else if /I "%MONTHLY_BOOL_INPUT%"=="N" (

    set MONTHLY_BOOL=false
) else if "%MONTHLY_BOOL_INPUT%"=="" (
    set MONTHLY_BOOL=false
) else (
    call :cecho Please enter Y or N for the monthly update question.
    goto :MONTHLY_BOOL_PROMPT
)


if "%MONTHLY_BOOL%"=="true" (
    call :cecho Monthly update selected. Next data date will be calculated based on the current data date.
    for /f %%i in ('python "%PROJECT_ROOT%\utilities\date_calculations.py" --old_data_date %DATE_INPUT% --monthly_bool %MONTHLY_BOOL%') do set "NEW_DATA_DATE=%%i"
    call :cecho Selected next data date: %NEW_DATA_DATE%
    goto :BUILD_SHEETS
)

:CUSTOM_NEW_DD_PROMPT
    call :cecho Enter the Next Data Date (MMDDYY format, e.g., 092525):
    set /p NEW_DATA_DATE=

:BUILD_SHEETS
call :cecho Building update sheets with current data date %DATE_INPUT% and next data date %NEW_DATA_DATE%
python "%SCRIPT_DIR%build_p6_update_sheet.py" --old_data_date %DATE_INPUT% --new_data_date %NEW_DATA_DATE% 

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