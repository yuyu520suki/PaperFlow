"""Local Windows OCR for a single image; never treats OCR as formula recovery."""
from __future__ import annotations

import json
import os
from pathlib import Path
import re
import subprocess
from typing import Any


SUPPORTED_IMAGES = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"}


class OcrError(RuntimeError):
    pass


def _powershell() -> Path:
    if os.name != "nt":
        raise OcrError("Windows.Media.Ocr requires Windows")
    executable = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32/WindowsPowerShell/v1.0/powershell.exe"
    if not executable.is_file():
        raise OcrError("Windows PowerShell 5.1 is unavailable; PowerShell 7 is not a WinRT replacement")
    return executable


def _invoke(arguments: list[str], timeout: float) -> dict[str, Any]:
    script = Path(__file__).resolve().parents[1] / "scripts/windows_ocr.ps1"
    if not script.is_file():
        raise OcrError("Packaged Windows OCR script is missing")
    command = [str(_powershell()), "-NoProfile", "-NonInteractive", "-File", str(script), *arguments]
    try:
        result = subprocess.run(command, capture_output=True, encoding="utf-8-sig", errors="replace", timeout=timeout, check=False)
    except subprocess.TimeoutExpired as exc:
        raise OcrError("Windows OCR helper timed out; no result was published") from exc
    try:
        payload = json.loads(result.stdout.strip())
    except (ValueError, TypeError) as exc:
        raise OcrError("Windows OCR helper returned invalid JSON: " + (result.stderr or result.stdout)[-1000:]) from exc
    if result.returncode or "error" in payload:
        raise OcrError(payload.get("message") or result.stderr.strip() or "Windows OCR helper failed")
    return payload


def ocr_status(*, timeout: float = 15) -> dict[str, Any]:
    """Read installed recognizer languages; never install or change a language."""
    return _invoke(["-Status"], timeout)


def _normalize_ocr_text(text: str) -> str:
    return re.sub(r"(?<=[\u3400-\u9fff]) +(?=[\u3400-\u9fff])", "", text)


def recognize_image(source: str | Path, output: str | Path | None = None, *, language: str | None = None, timeout: float = 60) -> dict[str, Any]:
    """Return text plus pixel-space line/word boxes for one image.

    Multi-frame TIFF and oversized images fail explicitly, avoiding a silent
    first-frame or resampling loss. Caller may render PDF pages individually.
    ``output`` is an optional JSON sidecar, created exclusively.
    """
    source = Path(source).resolve(strict=True)
    if not source.is_file() or source.suffix.lower() not in SUPPORTED_IMAGES:
        raise ValueError("OCR input must be a supported single-page image")
    target = Path(output).resolve() if output is not None else None
    if target is not None and target.exists():
        raise FileExistsError(target)
    if timeout <= 0:
        raise ValueError("timeout must be positive")
    if language is not None and (not language or not all(c.isalnum() or c == "-" for c in language)):
        raise ValueError("language must be a BCP 47 language tag")
    arguments = ["-ImagePath", str(source), "-TimeoutSeconds", str(min(120, max(1, int(timeout)-5)))]
    if language:
        arguments.extend(["-Language", language])
    payload = _invoke(arguments, timeout)
    if payload.get("engine") != "Windows.Media.Ocr" or not isinstance(payload.get("text"), str) or not isinstance(payload.get("lines"), list):
        raise OcrError("Windows OCR result does not satisfy its text/geometry contract")
    payload["raw_text"] = payload["text"]
    # OcrResult.Text concatenates lines on this engine. Reconstruct explicit
    # recognized line breaks, retaining the original and boxes for review.
    payload["text"] = "\n".join(line["text"] for line in payload["lines"])
    payload["normalized_text"] = _normalize_ocr_text(payload["text"])
    payload["normalization"] = "recognized line breaks restored; normalized_text additionally removes inter-CJK spaces; raw_text retained"
    # These are invariants of the import adapter, not claims made by the engine.
    payload["requires_review"] = True
    payload["equations_reconstructed"] = False
    payload["table_structure_reconstructed"] = False
    if target is not None:
        target.parent.mkdir(parents=True, exist_ok=True)
        with target.open("x", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
            f.write("\n")
    return payload
