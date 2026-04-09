#!/usr/bin/env python3
"""Toyo Schedule Maker - GUI Installer

Checks for dependencies and installs them with a Tkinter progress window.
Launched automatically by run.sh on first run or when deps are missing.
"""

import os
import sys
import subprocess
import threading
import tkinter as tk
from tkinter import ttk
from pathlib import Path

VENV_DIR = Path.home() / ".toyo_scheduler" / "venv"
PYTHON = VENV_DIR / "bin" / "python3"
PIP = VENV_DIR / "bin" / "pip"

OCR_VENV_DIR = Path.home() / ".toyo_scheduler" / "ocr_venv"
OCR_PYTHON = OCR_VENV_DIR / "bin" / "python3"


def check_venv():
    return PYTHON.exists()


def check_core_deps():
    """Check if openpyxl and fpdf2 are importable."""
    if not check_venv():
        return False
    try:
        result = subprocess.run(
            [str(PYTHON), "-c", "import openpyxl; from fpdf import FPDF"],
            capture_output=True, timeout=10)
        return result.returncode == 0
    except Exception:
        return False


def check_ocr_deps():
    """Check if PaddleOCR dependencies are installed in the OCR venv."""
    if not OCR_PYTHON.exists():
        return False
    try:
        result = subprocess.run(
            [str(OCR_PYTHON), "-c", "from paddleocr import PaddleOCR"],
            capture_output=True, timeout=30)
        # Also check that opencv is in the main venv
        if result.returncode != 0:
            return False
        result2 = subprocess.run(
            [str(PYTHON), "-c", "import cv2"],
            capture_output=True, timeout=10)
        return result2.returncode == 0
    except Exception:
        return False


class InstallerWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Toyo Schedule Maker - Setup")
        self.root.geometry("520x400")
        self.root.resizable(False, False)

        # Center on screen
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - 520) // 2
        y = (self.root.winfo_screenheight() - 400) // 2
        self.root.geometry(f"+{x}+{y}")

        self.cancelled = False
        self.install_success = False

        self._build_ui()
        self._check_status()

    def _build_ui(self):
        # Title
        ttk.Label(self.root, text="Toyo Schedule Maker",
                  font=("Helvetica", 16, "bold")).pack(pady=(20, 5))
        ttk.Label(self.root, text="First-Time Setup",
                  font=("Helvetica", 11)).pack(pady=(0, 15))

        # Status frame
        status_frame = ttk.LabelFrame(self.root, text="Components", padding=10)
        status_frame.pack(fill=tk.X, padx=20, pady=5)

        # Venv status
        row1 = ttk.Frame(status_frame)
        row1.pack(fill=tk.X, pady=2)
        ttk.Label(row1, text="Python Environment", width=25, anchor="w").pack(side=tk.LEFT)
        self.venv_status = ttk.Label(row1, text="Checking...", width=15)
        self.venv_status.pack(side=tk.RIGHT)

        # Core deps status
        row2 = ttk.Frame(status_frame)
        row2.pack(fill=tk.X, pady=2)
        ttk.Label(row2, text="Core Dependencies", width=25, anchor="w").pack(side=tk.LEFT)
        self.core_status = ttk.Label(row2, text="Checking...", width=15)
        self.core_status.pack(side=tk.RIGHT)

        # OCR status
        row3 = ttk.Frame(status_frame)
        row3.pack(fill=tk.X, pady=2)
        ttk.Label(row3, text="OCR (Photo Import)", width=25, anchor="w").pack(side=tk.LEFT)
        self.ocr_status = ttk.Label(row3, text="Checking...", width=15)
        self.ocr_status.pack(side=tk.RIGHT)

        # OCR checkbox
        self.install_ocr_var = tk.BooleanVar(value=True)
        self.ocr_check = ttk.Checkbutton(
            self.root,
            text="Install OCR support (reads schedules from photos, ~1.5 GB download)",
            variable=self.install_ocr_var)
        self.ocr_check.pack(pady=(10, 5), padx=20, anchor="w")

        # Progress
        self.progress_label = ttk.Label(self.root, text="", font=("Helvetica", 10))
        self.progress_label.pack(pady=(10, 3), padx=20, anchor="w")

        self.progress = ttk.Progressbar(self.root, mode="indeterminate", length=460)
        self.progress.pack(padx=20, pady=3)

        # Buttons
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill=tk.X, padx=20, pady=(15, 10))

        self.cancel_btn = ttk.Button(btn_frame, text="Cancel", command=self._cancel)
        self.cancel_btn.pack(side=tk.RIGHT, padx=5)

        self.install_btn = ttk.Button(btn_frame, text="Install", command=self._start_install)
        self.install_btn.pack(side=tk.RIGHT, padx=5)

        self.skip_btn = ttk.Button(btn_frame, text="Skip OCR", command=self._skip_ocr)
        self.skip_btn.pack(side=tk.RIGHT, padx=5)
        self.skip_btn.pack_forget()  # hidden initially

    def _set_status(self, label, installed):
        if installed:
            label.configure(text="Installed", foreground="green")
        else:
            label.configure(text="Not installed", foreground="red")

    def _check_status(self):
        has_venv = check_venv()
        has_core = check_core_deps()
        has_ocr = check_ocr_deps()

        self._set_status(self.venv_status, has_venv)
        self._set_status(self.core_status, has_core)
        self._set_status(self.ocr_status, has_ocr)

        if has_venv and has_core and has_ocr:
            self.progress_label.configure(text="Everything is installed!")
            self.install_btn.configure(text="Launch App", command=self._launch)
            self.ocr_check.pack_forget()
            self.install_success = True
        elif has_venv and has_core and not has_ocr:
            self.progress_label.configure(text="Core app is ready. OCR is optional.")
            self.skip_btn.pack(side=tk.RIGHT, padx=5)
            self.install_ocr_var.set(True)

    def _cancel(self):
        self.cancelled = True
        self.root.destroy()

    def _skip_ocr(self):
        self.install_success = True
        self.root.destroy()

    def _launch(self):
        self.install_success = True
        self.root.destroy()

    def _start_install(self):
        self.install_btn.configure(state="disabled")
        self.cancel_btn.configure(state="disabled")
        self.skip_btn.pack_forget()
        self.ocr_check.configure(state="disabled")
        self.progress.start(15)

        thread = threading.Thread(target=self._install_thread, daemon=True)
        thread.start()

    def _install_thread(self):
        try:
            # Step 1: Create venv if needed
            if not check_venv():
                self._update_status("Creating Python environment...")
                VENV_DIR.parent.mkdir(parents=True, exist_ok=True)
                result = subprocess.run(
                    [sys.executable, "-m", "venv", str(VENV_DIR)],
                    capture_output=True, text=True)
                if result.returncode != 0:
                    self._show_error(f"Failed to create venv:\n{result.stderr}")
                    return
                self._update_label(self.venv_status, "Installed", "green")

            # Step 2: Install core deps
            if not check_core_deps():
                self._update_status("Installing core dependencies...")
                result = subprocess.run(
                    [str(PIP), "install", "--quiet", "openpyxl", "fpdf2"],
                    capture_output=True, text=True, timeout=300)
                if result.returncode != 0:
                    self._show_error(f"Failed to install core deps:\n{result.stderr}")
                    return
                self._update_label(self.core_status, "Installed", "green")

            # Step 3: Install OCR if selected
            if self.install_ocr_var.get() and not check_ocr_deps():
                # Install opencv in main venv
                self._update_status("Installing OpenCV...")
                result = subprocess.run(
                    [str(PIP), "install", "--quiet", "opencv-python-headless"],
                    capture_output=True, text=True, timeout=300)
                if result.returncode != 0:
                    self._show_error(f"Failed to install OpenCV:\n{result.stderr}")
                    return

                # Create OCR venv with Python 3.12 (PaddlePaddle needs <=3.12)
                self._update_status("Setting up OCR environment...")
                if not OCR_PYTHON.exists():
                    # Try uv first, fall back to python3.12 directly
                    uv_bin = Path.home() / ".local" / "bin" / "uv"
                    if uv_bin.exists():
                        result = subprocess.run(
                            [str(uv_bin), "venv", str(OCR_VENV_DIR), "--python", "3.12"],
                            capture_output=True, text=True)
                    else:
                        result = subprocess.run(
                            ["python3.12", "-m", "venv", str(OCR_VENV_DIR)],
                            capture_output=True, text=True)
                    if result.returncode != 0:
                        self._show_error(
                            "Failed to create OCR environment. Python 3.12 is required "
                            "for PaddleOCR.\n\nInstall it with: curl -LsSf "
                            "https://astral.sh/uv/install.sh | sh && uv python install 3.12"
                            f"\n\n{result.stderr}")
                        return

                self._update_status("Installing PaddleOCR (this may take a while)...")
                # Use uv pip if available, otherwise fall back to pip
                uv_bin = Path.home() / ".local" / "bin" / "uv"
                ocr_pip = OCR_VENV_DIR / "bin" / "pip"
                if uv_bin.exists():
                    result = subprocess.run(
                        [str(uv_bin), "pip", "install",
                         "--python", str(OCR_PYTHON),
                         "paddlepaddle>=3.0,<3.1", "paddleocr>=3.0"],
                        capture_output=True, text=True, timeout=600)
                elif ocr_pip.exists():
                    result = subprocess.run(
                        [str(ocr_pip), "install",
                         "paddlepaddle>=3.0,<3.1", "paddleocr>=3.0"],
                        capture_output=True, text=True, timeout=600)
                else:
                    result = subprocess.run(
                        [str(OCR_PYTHON), "-m", "pip", "install",
                         "paddlepaddle>=3.0,<3.1", "paddleocr>=3.0"],
                        capture_output=True, text=True, timeout=600)
                if result.returncode != 0:
                    self._show_error(f"Failed to install PaddleOCR:\n{result.stderr}")
                    return

                # Pre-download OCR models so first use is instant
                self._update_status("Downloading OCR models...")
                result = subprocess.run(
                    [str(OCR_PYTHON), "-c",
                     "import os; os.environ['PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK']='True'; "
                     "from paddleocr import PaddleOCR; "
                     "PaddleOCR(lang='en', device='cpu', "
                     "use_doc_orientation_classify=False, "
                     "use_doc_unwarping=False, "
                     "use_textline_orientation=False)"],
                    capture_output=True, text=True, timeout=300)
                if result.returncode != 0:
                    self._show_error(f"Failed to download OCR models:\n{result.stderr}")
                    return
                self._update_label(self.ocr_status, "Installed", "green")

            self._finish_success()

        except subprocess.TimeoutExpired:
            self._show_error("Installation timed out. Check your internet connection.")
        except Exception as e:
            self._show_error(str(e))

    def _update_status(self, text):
        self.root.after(0, lambda: self.progress_label.configure(text=text))

    def _update_label(self, label, text, color):
        self.root.after(0, lambda: label.configure(text=text, foreground=color))

    def _show_error(self, msg):
        def _do():
            self.progress.stop()
            self.progress_label.configure(text="Installation failed!", foreground="red")
            self.install_btn.configure(state="normal", text="Retry", command=self._start_install)
            self.cancel_btn.configure(state="normal")
            from tkinter import messagebox
            messagebox.showerror("Installation Error", msg)
        self.root.after(0, _do)

    def _finish_success(self):
        def _do():
            self.progress.stop()
            self.progress_label.configure(text="Installation complete!", foreground="green")
            self.install_btn.configure(state="normal", text="Launch App", command=self._launch)
            self.cancel_btn.configure(state="normal")
            self.install_success = True
        self.root.after(0, _do)

    def run(self):
        self.root.mainloop()
        return self.install_success


def needs_install():
    """Return True if installer should run."""
    return not check_core_deps()


if __name__ == "__main__":
    # Can be run directly for testing
    app = InstallerWindow()
    if app.run():
        print("Ready to launch app.")
    else:
        print("Installation cancelled.")
