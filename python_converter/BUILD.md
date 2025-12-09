# Document Converter - Build Instructions

## Prerequisites
- Python 3.10 or higher
- pip (Python package manager)

## Setup Steps

### 1. Install Dependencies
```powershell
cd python_converter
pip install -r requirements.txt
pip install pyinstaller
```

### 2. Build EXE
```powershell
pyinstaller --onefile --noconsole --name converter converter.py
```

### 3. Copy EXE to Executables
After build completes, find the EXE at:
```
python_converter\dist\converter.exe
```

Copy it to:
```
ubpdfs\executables\converter.exe
```

## Usage Examples

From Node.js/Electron:
```javascript
// PDF to Word
execFile('executables/converter.exe', ['pdf-to-word', 'input.pdf', 'output.docx'])

// Word to PDF
execFile('executables/converter.exe', ['word-to-pdf', 'input.docx', 'output.pdf'])

// PDF to Excel
execFile('executables/converter.exe', ['pdf-to-excel', 'input.pdf', 'output.xlsx'])

// Excel to PDF
execFile('executables/converter.exe', ['excel-to-pdf', 'input.xlsx', 'output.pdf'])

// HTML to PDF
execFile('executables/converter.exe', ['html-to-pdf', 'input.html', 'output.pdf'])

// PowerPoint to PDF
execFile('executables/converter.exe', ['ppt-to-pdf', 'input.pptx', 'output.pdf'])

// PDF to HTML
execFile('executables/converter.exe', ['pdf-to-html', 'input.pdf', 'output.html'])

// PDF to PowerPoint
execFile('executables/converter.exe', ['pdf-to-ppt', 'input.pdf', 'output.pptx'])
```

## Conversion Types

- `pdf-to-word` - Convert PDF to Word DOCX
- `word-to-pdf` - Convert Word to PDF
- `pdf-to-excel` - Extract tables from PDF to Excel
- `excel-to-pdf` - Convert Excel to PDF
- `html-to-pdf` - Convert HTML to PDF
- `ppt-to-pdf` - Convert PowerPoint to PDF
- `pdf-to-html` - Convert PDF to HTML
- `pdf-to-ppt` - Convert PDF to PowerPoint (basic)

## Expected EXE Size
~40-50MB (includes all libraries)

## Testing
```powershell
# Test PDF to Word
.\converter.exe pdf-to-word test.pdf output.docx

# Test Word to PDF
.\converter.exe word-to-pdf test.docx output.pdf
```
