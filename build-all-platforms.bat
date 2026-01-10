@echo off
REM Build for all Windows platforms (32-bit and 64-bit)

echo ========================================
echo Building UB PDF for All Windows Platforms
echo ========================================
echo.

REM Install npm dependencies
echo Installing npm dependencies...
call npm install

REM Build Python converter
echo.
echo Building Python converter...
cd python_converter
call build_windows.bat
cd ..

REM Build Electron app for all Windows architectures
echo.
echo Building Electron app for Windows (32-bit and 64-bit)...
call npm run build -- --win --ia32 --x64

echo.
echo ========================================
echo Build complete!
echo ========================================
echo Installers location: release\
dir release\*.exe
pause
