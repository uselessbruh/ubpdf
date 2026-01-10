#!/bin/bash
# Build Python converter for Linux
# Run this script on a Linux machine or WSL

echo "Building Linux converter executable..."

# Install dependencies
pip install -r requirements_lite.txt
pip install pyinstaller

# Build converter_lite (lightweight version)
pyinstaller --onefile \
    --name converter_linux \
    --add-data "converter_lite.py:." \
    --hidden-import=pdf2docx \
    --hidden-import=docx \
    --hidden-import=reportlab \
    --hidden-import=openpyxl \
    --hidden-import=pptx \
    --hidden-import=pdfplumber \
    --hidden-import=PIL \
    --collect-all pdf2docx \
    --collect-all reportlab \
    --collect-all pdfplumber \
    converter_lite.py

# Move executable to parent executables directory
mkdir -p ../executables/bin
cp dist/converter_linux ../executables/bin/

echo "✓ Linux converter built successfully!"
echo "Binary location: ../executables/bin/converter_linux"
