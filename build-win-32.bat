@echo off
REM Build Windows 32-bit installer
REM Requires Python 32-bit installed

echo ========================================
echo Building UB PDF for Windows 32-bit
echo ========================================
echo.

REM Check if Node.js is installed
where node >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Node.js is not installed!
    echo Please install Node.js from https://nodejs.org/
    pause
    exit /b 1
)

REM Install npm dependencies
echo Installing npm dependencies...
call npm install

REM Build Python converter for 32-bit
echo.
echo Building Python converter (32-bit)...
cd python_converter
call build_windows.bat
cd ..

REM Build Electron app for Windows 32-bit
echo.
echo Building Electron app for Windows 32-bit...
call npm run build -- --win --ia32

echo.
echo ========================================
echo Build complete!
echo ========================================
echo Installer location: release\
dir release\*.exe
pause
