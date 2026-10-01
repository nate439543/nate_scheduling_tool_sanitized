@echo off
REM Prompt user for the data date in MMDDYY format
set /p datadate=Enter date (MMDDYY format, e.g., 090525): 

REM Activate the virtual environment and run the script
call "%~dp0venv\Scripts\activate"
python "%~dp0generate_schedules.py" %datadate%
pause
