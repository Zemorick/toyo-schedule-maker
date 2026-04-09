#!/usr/bin/env python3
"""PaddleOCR worker — runs in a separate Python 3.12 venv.

Called by toyo_scheduler.py as a subprocess because PaddlePaddle requires
Python <=3.12, while the main app uses system Python for proper Tk fonts.

Protocol: receives image path on stdin, writes JSON results to stdout.
"""

import json
import os
import sys

os.environ['PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK'] = 'True'

from paddleocr import PaddleOCR

_reader = None


def get_reader():
    global _reader
    if _reader is None:
        _reader = PaddleOCR(
            lang='en', device='cpu',
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_textline_orientation=False)
    return _reader


def readtext(img_path):
    """Run OCR on an image, return list of (bbox, text, confidence)."""
    reader = get_reader()
    out = []
    for r in reader.predict(img_path):
        texts = r['rec_texts']
        scores = r['rec_scores']
        polys = r['dt_polys']
        for poly, text, score in zip(polys, texts, scores):
            out.append([poly.tolist(), text, float(score)])
    return out


def main():
    # Signal readiness
    print(json.dumps({"status": "ready"}), flush=True)

    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            request = json.loads(line)
        except json.JSONDecodeError:
            print(json.dumps({"error": "invalid JSON"}), flush=True)
            continue

        cmd = request.get("cmd")
        if cmd == "quit":
            break
        elif cmd == "ocr":
            img_path = request.get("path", "")
            if not os.path.exists(img_path):
                print(json.dumps({"error": f"file not found: {img_path}"}), flush=True)
                continue
            try:
                results = readtext(img_path)
                print(json.dumps({"results": results}), flush=True)
            except Exception as e:
                print(json.dumps({"error": str(e)}), flush=True)
        else:
            print(json.dumps({"error": f"unknown cmd: {cmd}"}), flush=True)


if __name__ == "__main__":
    main()
