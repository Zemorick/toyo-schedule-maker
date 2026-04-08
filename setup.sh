#!/bin/bash
# Toyo Schedule Maker - Setup Script
# Run this once to install dependencies

VENV_DIR="$HOME/.toyo_scheduler/venv"

echo "Setting up Toyo Schedule Maker..."

if [ ! -d "$VENV_DIR" ]; then
    echo "Creating virtual environment at $VENV_DIR..."
    mkdir -p "$HOME/.toyo_scheduler"
    python3 -m venv "$VENV_DIR"
fi

echo "Installing core dependencies..."
"$VENV_DIR/bin/pip" install --quiet openpyxl fpdf2

echo ""
echo "Would you like to install OCR support (Import from Photo)?"
echo "This allows reading handwritten sign-up sheets from photos using EasyOCR."
echo "Note: This downloads several GB (PyTorch + EasyOCR + OpenCV)."
read -p "Install OCR? [y/N] " ocr_choice
if [[ "$ocr_choice" =~ ^[Yy]$ ]]; then
    echo "Installing OCR libraries (this may take a while)..."
    "$VENV_DIR/bin/pip" install easyocr opencv-python-headless
    echo "OCR support installed!"
else
    echo "Skipping OCR. You can install it later by re-running setup.sh."
fi

echo ""
echo "Setup complete! Run the app with:"
echo "  ./run.sh"
