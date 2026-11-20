"""LibreOffice discovery and conversion, safe for concurrent Streamlit users."""
from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
import threading
from pathlib import Path

_LOCK = threading.Lock()

_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/usr/lib/libreoffice/program/soffice",
    "/usr/local/bin/soffice",
    "/opt/libreoffice/program/soffice",
]


class OfficeUnavailable(RuntimeError):
    pass


class ConversionError(RuntimeError):
    pass


def find_soffice() -> str | None:
    override = os.environ.get("SOFFICE_PATH")
    if override and os.path.exists(override):
        return override
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for path in _CANDIDATES:
        if os.path.exists(path):
            return path
    return None


def available() -> bool:
    return find_soffice() is not None


def _safe_stem(name: str) -> str:
    stem = Path(name).stem or "document"
    return "".join(c if c.isalnum() or c in "-_." else "_" for c in stem)[:60]


def convert(data: bytes, name: str, target: str = "docx", timeout: int = 180) -> bytes:
    """Convert *data* with LibreOffice. ``target`` is a LO filter name.

    Each call gets its own user profile, otherwise concurrent conversions
    corrupt each other and soffice silently returns partial files.
    """
    exe = find_soffice()
    if not exe:
        raise OfficeUnavailable(
            "LibreOffice is not installed or soffice is not on PATH. "
            "Install LibreOffice to convert legacy formats and export PDFs.")

    stem = _safe_stem(name)
    ext = target.split(":")[0]
    with tempfile.TemporaryDirectory(prefix="rebrand-") as td:
        src = os.path.join(td, f"{stem}.{Path(name).suffix.lstrip('.') or 'bin'}")
        with open(src, "wb") as fh:
            fh.write(data)
        profile = Path(td, "lo-profile").as_uri()
        cmd = [
            exe,
            f"-env:UserInstallation={profile}",
            "--headless", "--nologo", "--norestore", "--nolockcheck",
            "--nodefault", "--nofirststartwizard",
            "--convert-to", target,
            "--outdir", td, src,
        ]
        try:
            with _LOCK:
                proc = subprocess.run(cmd, capture_output=True, timeout=timeout)
        except subprocess.TimeoutExpired as exc:
            raise ConversionError(
                f"LibreOffice timed out after {timeout}s converting {name}") from exc

        out = os.path.join(td, f"{stem}.{ext}")
        if not os.path.exists(out):
            detail = (proc.stderr or proc.stdout or b"").decode("utf-8", "replace")[-500:]
            raise ConversionError(
                f"LibreOffice could not convert {name} to {ext}. {detail}".strip())
        with open(out, "rb") as fh:
            return fh.read()


def to_docx(data: bytes, name: str) -> bytes:
    return convert(data, name, "docx")


def to_pdf(data: bytes, name: str, timeout: int = 180) -> bytes:
    return convert(data, name, "pdf", timeout=timeout)
