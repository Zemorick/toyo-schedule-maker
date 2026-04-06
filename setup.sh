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

echo "Installing dependencies..."
"$VENV_DIR/bin/pip" install --quiet openpyxl fpdf2

echo ""
echo "Setup complete! Run the app with:"
echo "  ./run.sh"
