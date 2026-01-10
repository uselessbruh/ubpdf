#!/bin/bash
# Build AppImage for Linux
# Must be run on Linux (Ubuntu/Debian recommended)

set -e

echo "🚀 Building UB PDF AppImage for Linux..."

# Install required tools
echo "📦 Installing dependencies..."
sudo apt-get update
sudo apt-get install -y python3 python3-pip nodejs npm

# Install Electron builder globally
npm install

# Build Python converter
echo "🐍 Building Python converter..."
cd python_converter
chmod +x build_linux.sh
./build_linux.sh
cd ..

# Build Electron app for Linux
echo "⚡ Building Electron app..."
npm run build -- --linux AppImage

echo "✅ Build complete!"
echo "📍 AppImage location: release/"
ls -lh release/*.AppImage
