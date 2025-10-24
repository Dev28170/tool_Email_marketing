@echo off
echo ========================================
echo Email Mass Sender - Executable Builder
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8+ and try again
    pause
    exit /b 1
)

echo Python found. Installing dependencies...
echo.

REM Install PyInstaller if not already installed
python -m pip install pyinstaller

echo.
echo Building executable...
echo This may take several minutes...
echo.

REM Run the build script
python build_exe.py

echo.
echo Build process completed!
echo Check the dist/EmailMassSender folder for your executable
echo.
pause

