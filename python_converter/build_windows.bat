@echo off
REM Build Python converter for Windows (64-bit and 32-bit)

echo Building Windows converter executable...

REM Install dependencies
pip install -r requirements_lite.txt
pip install pyinstaller

REM Build converter_lite for Windows 64-bit
echo Building 64-bit version...
pyinstaller --onefile ^
    --name converter ^
    --icon=..\assets\logo.ico ^
    --add-data "converter_lite.py;." ^
    --hidden-import=pdf2docx ^
    --hidden-import=docx ^
    --hidden-import=reportlab ^
    --hidden-import=openpyxl ^
    --hidden-import=pptx ^
    --hidden-import=pdfplumber ^
    --hidden-import=PIL ^
    --collect-all pdf2docx ^
    --collect-all reportlab ^
    --collect-all pdfplumber ^
    --noconfirm ^
    converter_lite.py

REM Move executable
if not exist "..\executables\bin" mkdir "..\executables\bin"
copy dist\converter.exe ..\executables\bin\converter.exe

echo.
echo ✓ Windows converter built successfully!
echo Binary location: ..\executables\bin\converter.exe
pause
