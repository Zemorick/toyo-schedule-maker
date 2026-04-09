#!/bin/bash
# Toyo Schedule Maker - Setup Script
# Run this once to install dependencies

VENV_DIR="$HOME/.toyo_scheduler/venv"
OCR_VENV_DIR="$HOME/.toyo_scheduler/ocr_venv"

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
echo "This allows reading handwritten sign-up sheets from photos using PaddleOCR."
echo "Note: This downloads ~1.5 GB (PaddlePaddle + PaddleOCR + OpenCV)."
read -p "Install OCR? [y/N] " ocr_choice
if [[ "$ocr_choice" =~ ^[Yy]$ ]]; then
    echo "Installing OpenCV in main environment..."
    "$VENV_DIR/bin/pip" install --quiet opencv-python-headless

    # PaddleOCR needs Python <=3.12; create a separate venv
    if [ ! -d "$OCR_VENV_DIR" ]; then
        echo "Creating OCR environment (Python 3.12 required for PaddlePaddle)..."
        if command -v uv &>/dev/null; then
            uv venv "$OCR_VENV_DIR" --python 3.12
        elif command -v "$HOME/.local/bin/uv" &>/dev/null; then
            "$HOME/.local/bin/uv" venv "$OCR_VENV_DIR" --python 3.12
        elif command -v python3.12 &>/dev/null; then
            python3.12 -m venv "$OCR_VENV_DIR"
        else
            echo "ERROR: Python 3.12 is required for PaddleOCR."
            echo "Install it with: curl -LsSf https://astral.sh/uv/install.sh | sh && uv python install 3.12"
            exit 1
        fi
    fi

    echo "Installing PaddleOCR (this may take a while)..."
    if command -v uv &>/dev/null; then
        uv pip install --python "$OCR_VENV_DIR/bin/python3" "paddlepaddle>=3.0,<3.1" "paddleocr>=3.0"
    elif command -v "$HOME/.local/bin/uv" &>/dev/null; then
        "$HOME/.local/bin/uv" pip install --python "$OCR_VENV_DIR/bin/python3" "paddlepaddle>=3.0,<3.1" "paddleocr>=3.0"
    else
        "$OCR_VENV_DIR/bin/pip" install "paddlepaddle>=3.0,<3.1" "paddleocr>=3.0"
    fi

    echo "Downloading OCR models..."
    PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK=True "$OCR_VENV_DIR/bin/python3" -c "
from paddleocr import PaddleOCR
PaddleOCR(lang='en', device='cpu',
          use_doc_orientation_classify=False,
          use_doc_unwarping=False,
          use_textline_orientation=False)
print('Models downloaded successfully.')
"
    echo "OCR support installed!"
else
    echo "Skipping OCR. You can install it later by re-running setup.sh."
fi

echo ""
echo "Setup complete! Run the app with:"
echo "  ./run.sh"
