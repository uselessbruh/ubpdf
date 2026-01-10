# UB PDF Build Guide

## Build Instructions for All Platforms

### 📦 Prerequisites
- **Node.js** 16+ (https://nodejs.org/)
- **Python** 3.8+ (https://www.python.org/)
- **Git** (for cloning the repository)

---

## 🪟 Windows Builds

### Windows 64-bit Only
```bash
build-win-64.bat
```

### Windows 32-bit Only
```bash
build-win-32.bat
```

### Both 32-bit and 64-bit
```bash
build-all-platforms.bat
```

Or manually:
```bash
npm install
cd python_converter
build_windows.bat
cd ..
npm run build:win
```

**Output:** `release/UB PDF Setup x.x.x.exe` (32-bit and 64-bit versions)

---

## 🐧 Linux Builds

### AppImage (Universal Linux Package)
```bash
chmod +x build-linux-appimage.sh
./build-linux-appimage.sh
```

Or manually:
```bash
npm install
cd python_converter
chmod +x build_linux.sh
./build_linux.sh
cd ..
npm run build:linux
```

**Output:** `release/UB-PDF-x.x.x.AppImage`

### .deb Package (Debian/Ubuntu)
```bash
chmod +x build-linux-deb.sh
./build-linux-deb.sh
```

**Output:** `release/ub-pdf_x.x.x_amd64.deb`

---

## 📋 Build Scripts Overview

| Script | Platform | Architecture | Output |
|--------|----------|--------------|--------|
| `build-win-32.bat` | Windows | 32-bit (ia32) | `.exe` installer |
| `build-win-64.bat` | Windows | 64-bit (x64) | `.exe` installer |
| `build-all-platforms.bat` | Windows | Both | Two `.exe` installers |
| `build-linux-appimage.sh` | Linux | 64-bit | `.AppImage` |
| `build-linux-deb.sh` | Linux | 64-bit | `.deb` package |

---

## 🚀 Quick Start Development

### Start Development Server
```bash
npm install
npm start
```

### Build for Current Platform
```bash
npm run build
```

### Build for Specific Platform
```bash
# Windows 32-bit
npm run build:win32

# Windows 64-bit
npm run build:win64

# Windows (both)
npm run build:win

# Linux (AppImage + deb)
npm run build:linux

# All platforms
npm run build:all
```

---

## 📁 Project Structure

```
ubpdfs/
├── main.js                    # Electron main process
├── package.json               # Project configuration
├── build-win-32.bat          # Windows 32-bit build script
├── build-win-64.bat          # Windows 64-bit build script
├── build-linux-appimage.sh   # Linux AppImage build script
├── build-linux-deb.sh        # Linux .deb build script
├── python_converter/
│   ├── converter_lite.py     # Lightweight converter
│   ├── build_windows.bat     # Python build for Windows
│   └── build_linux.sh        # Python build for Linux
├── executables/
│   └── bin/
│       ├── converter.exe     # Windows converter binary
│       └── converter_linux   # Linux converter binary
└── release/                   # Build output folder
```

---

## 🔧 Troubleshooting

### Windows Build Issues

**Python not found:**
```bash
# Add Python to PATH or specify full path
C:\Python39\python.exe -m pip install -r requirements_lite.txt
```

**Node.js not found:**
- Install from https://nodejs.org/
- Restart terminal after installation

### Linux Build Issues

**Permission denied:**
```bash
chmod +x build-linux-appimage.sh
chmod +x python_converter/build_linux.sh
```

**Missing dependencies:**
```bash
sudo apt-get update
sudo apt-get install python3 python3-pip nodejs npm
```

---

## 📦 Dependencies

### Python Libraries (Lightweight)
- pdf2docx - PDF to Word conversion
- python-docx - Word document handling
- reportlab - PDF generation
- openpyxl - Excel file handling
- python-pptx - PowerPoint handling
- pdfplumber - PDF reading and extraction
- Pillow - Image processing

### Node.js Packages
- electron - Desktop app framework
- electron-builder - Multi-platform builder
- pdf-lib - PDF manipulation

---

## ✅ Post-Build

After building, installers/packages are located in the `release/` folder:

**Windows:**
- `UB PDF Setup 1.0.0.exe` (64-bit)
- `UB PDF Setup 1.0.0-ia32.exe` (32-bit)

**Linux:**
- `UB-PDF-1.0.0.AppImage`
- `ub-pdf_1.0.0_amd64.deb`

## 🎯 Distribution

- **Windows:** Run the `.exe` installer
- **Linux AppImage:** Make executable and run: `chmod +x *.AppImage && ./UB-PDF*.AppImage`
- **Linux .deb:** Install with: `sudo dpkg -i *.deb`
