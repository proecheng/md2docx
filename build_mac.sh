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

# Get latex2mathml data file path
LATEX2MATHML_PATH=$(python3 -c "import latex2mathml; import os; print(os.path.dirname(latex2mathml.__file__))")

echo
echo "Building Mac application..."
echo

pyinstaller --onefile --windowed --name "MD2DOCX" \
    --osx-bundle-identifier "com.proecheng.md2docx" \
    --add-data "${LATEX2MATHML_PATH}/unimathsymbols.txt:latex2mathml" \
    md2docx.py

echo
echo "==============================================="
if [ -d "dist/MD2DOCX.app" ]; then
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
