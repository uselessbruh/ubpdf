@echo off
REM Build Windows 64-bit installer

echo ========================================
echo Building UB PDF for Windows 64-bit
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

REM Build Python converter
echo.
echo Building Python converter...
cd python_converter
call build_windows.bat
cd ..

REM Build Electron app for Windows 64-bit
echo.
echo Building Electron app for Windows 64-bit...
call npm run build -- --win --x64

echo.
echo ========================================
echo Build complete!
echo ========================================
echo Installer location: release\
dir release\*.exe
pause
