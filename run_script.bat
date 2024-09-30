@echo off
set PYTHON_CMD=python

:: Check if Python is installed and available in the PATH
%PYTHON_CMD% --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Python is not installed or not in PATH.
    pause
    exit /b 1
)

:: Run main.py from the current directory
%PYTHON_CMD% "%~dp0main.py"

pause
