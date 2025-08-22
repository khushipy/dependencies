@echo off
echo Installing PyInstaller if needed...
pip install pyinstaller

echo Cleaning up previous builds...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "__pycache__" rmdir /s /q __pycache__

echo Building executable...
pyinstaller build.spec --clean

if %ERRORLEVEL% NEQ 0 (
    echo Build failed with error code %ERRORLEVEL%
    pause
    exit /b %ERRORLEVEL%
)

echo Copying required files to dist folder...
copy /Y ASAMD.exe dist\
copy /Y input_file.xlsx dist\
copy /Y input.txt dist\

echo Build complete! The executable is in the 'dist' folder.

rem Open the dist folder in explorer
explorer "%~dp0dist"

pause
