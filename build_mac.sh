#!/bin/bash
# MD2DOCX - Build Mac Application
# Run this script on macOS to create MD2DOCX.app

echo "==============================================="
echo "MD2DOCX - Build Mac Application"
echo "==============================================="
echo

# Check if running on macOS
if [[ "$OSTYPE" != "darwin"* ]]; then
    echo "Error: This script must be run on macOS"
    exit 1
fi

# Check Python
if ! command -v python3 &> /dev/null; then
    echo "Error: Python 3 is required"
    exit 1
fi

echo "Installing dependencies..."
pip3 install python-docx lxml latex2mathml pyinstaller

echo
echo "Building Mac application..."
echo

pyinstaller --onefile --windowed --name "MD2DOCX" \
    --osx-bundle-identifier "com.proecheng.md2docx" \
    md2docx.py

echo
echo "==============================================="
if [ -f "dist/MD2DOCX.app" ] || [ -d "dist/MD2DOCX.app" ]; then
    echo "Build successful!"
    echo
    echo "Application location: dist/MD2DOCX.app"
    echo
    echo "Usage:"
    echo "  1. Double-click MD2DOCX.app to open file dialog"
    echo "  2. Drag .md file onto MD2DOCX.app icon"
    echo "  3. Terminal: ./dist/MD2DOCX.app/Contents/MacOS/MD2DOCX file.md"
    echo
    echo "To install: drag MD2DOCX.app to /Applications"
else
    echo "Build failed. Please check errors above."
fi
echo "==============================================="
