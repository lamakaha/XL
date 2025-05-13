@echo off
echo xlwings System Information Collector
echo ================================
echo.
echo This script will collect system information to help diagnose COM communication errors
echo between Excel and Python when using xlwings.
echo.
echo The information will be saved to a CSV file in the current directory.
echo The CSV file will have three columns:
echo   - Group: The category of the parameter (e.g., python, excel, xlwings)
echo   - Parameter: The specific parameter name
echo   - Value: The value of the parameter
echo.
echo This format makes it easy to compare specific categories between machines.
echo.

REM Check if Python is installed and in PATH
where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo Python not found in PATH. Please install Python or add it to your PATH.
    echo.
    pause
    exit /b 1
)

REM Run the Python script
echo Running system information collector...
python xlwings_system_info_complete.py

echo.
echo If the script completed successfully, you should see a CSV file in this directory.
echo Please run this script on both working and problematic machines, then compare the results.
echo.
echo TIP: Open the CSV file in Excel and use the Filter feature to focus on specific groups.
echo      This makes it easier to identify differences between machines.
echo.
pause
