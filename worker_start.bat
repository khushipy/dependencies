@echo off
:: Change to the directory where this batch file is located
cd /d "%~dp0"

:: Check if required files exist
if not exist "main.exe" (
    echo Error: main.exe not found in %CD%
    pause
    exit /b 1
)

if not exist "ASAMD.exe" (
    echo Error: ASAMD.exe not found in %CD%
    pause
    exit /b 1
)

if not exist "input.txt" (
    echo Error: input.txt not found in %CD%
    pause
    exit /b 1
)

if not exist "input_file.xlsx" (
    echo Error: input_file.xlsx not found in %CD%
    pause
    exit /b 1
)

echo Starting worker process at %TIME%
start "" "main.exe" "input.txt" "input_file.xlsx"
echo Worker started at %TIME%