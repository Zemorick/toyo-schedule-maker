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
echo "This allows reading handwritten sign-up sheets from photos using TrOCR."
echo "Note: This downloads ~4GB (PyTorch + Transformers)."
read -p "Install OCR? [y/N] " ocr_choice
if [[ "$ocr_choice" =~ ^[Yy]$ ]]; then
    echo "Installing CPU-only PyTorch (this may take a while)..."
    "$VENV_DIR/bin/pip" install torch --index-url https://download.pytorch.org/whl/cpu
    echo "Installing TrOCR dependencies..."
    "$VENV_DIR/bin/pip" install transformers pillow
    echo "OCR support installed!"
else
    echo "Skipping OCR. You can install it later by re-running setup.sh."
fi

echo ""
echo "Setup complete! Run the app with:"
echo "  ./run.sh"
