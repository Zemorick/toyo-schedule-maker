#!/bin/bash
# Toyo Schedule Maker - Launcher
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV_DIR="$HOME/.toyo_scheduler/venv"

if [ ! -d "$VENV_DIR" ]; then
    echo "First run detected. Running setup..."
    bash "$SCRIPT_DIR/setup.sh"
fi

"$VENV_DIR/bin/python3" "$SCRIPT_DIR/toyo_scheduler.py"
