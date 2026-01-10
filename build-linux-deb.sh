#!/bin/bash
# Build .deb package for Debian/Ubuntu
# Must be run on Linux

set -e

echo "🚀 Building UB PDF .deb package for Linux..."

# Install required tools
echo "📦 Installing dependencies..."
sudo apt-get update
sudo apt-get install -y python3 python3-pip nodejs npm

# Install project dependencies
npm install

# Build Python converter
echo "🐍 Building Python converter..."
cd python_converter
chmod +x build_linux.sh
./build_linux.sh
cd ..

# Build Electron app as .deb
echo "⚡ Building Electron app..."
npm run build -- --linux deb

echo "✅ Build complete!"
echo "📍 .deb package location: release/"
ls -lh release/*.deb
