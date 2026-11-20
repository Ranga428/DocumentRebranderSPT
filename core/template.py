"""Brand template registry and contract validation."""
from __future__ import annotations

import hashlib
import io
import json
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path

import docx
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
TPL_DIR = ROOT / "templates"
FONT_DIR = ROOT / "assets" / "fonts"
MANIFEST_PATH = TPL_DIR / "manifest.json"


class TemplateError(RuntimeError):
    pass


@dataclass(frozen=True)
class TemplateSpec:
    key: str
    file: str
    path: Path
    sha256: str

    @property
    def orientation(self) -> str:
        return self.key

    def read(self) -> bytes:
        return self.path.read_bytes()


def manifest() -> dict:
    return json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))


def available() -> dict[str, TemplateSpec]:
    out: dict[str, TemplateSpec] = {}
    for key, entry in manifest()["templates"].items():
        path = TPL_DIR / entry["file"]
        out[key] = TemplateSpec(key=key, file=entry["file"], path=path,
                                sha256=str(entry.get("sha256", "")).lower())
    return out


def get(orientation: str) -> TemplateSpec:
    specs = available()
    spec = specs.get(orientation)
    if spec is None:
        raise TemplateError(f"Unknown page orientation: {orientation}")
    if not spec.path.exists():
        raise TemplateError(
            f"Template for {orientation} pages is missing: {spec.path.name}. "
            "Add it to the templates folder.")
    return spec


def choose(model, wanted: str = "auto") -> tuple[str, str | None]:
    """Return (template key, warning). Falls back when a variant is missing."""
    specs = available()
    have = {k for k, s in specs.items() if s.path.exists()}
    if not have:
        raise TemplateError("No brand template found in templates/.")

    warning = None
    if wanted in ("portrait", "landscape"):
        key = wanted
        if key not in have:
            other = "landscape" if key == "portrait" else "portrait"
            warning = (f"{key.title()} template not found; using {other} instead.")
            key = other
    else:
        key = "landscape" if model.has_wide_content() else "portrait"
        if key not in have:
            other = "landscape" if key == "portrait" else "portrait"
            warning = (f"Auto-selected {key}, but that template is missing; "
                       f"using {other} instead.")
            key = other
    return key, warning


def inspect(data: bytes) -> dict:
    """Cheap structural facts about a .docx package."""
    info: dict = {"fonts_embedded": False, "control_placeholder": False,
                  "orientation": None, "page_twips": None, "tables": 0}
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        names = z.namelist()
        info["fonts_embedded"] = any(n.startswith("word/fonts/") and n.endswith(".ttf")
                                     for n in names)
        for name in names:
            if re.fullmatch(r"word/(header|footer)\d*\.xml", name):
                try:
                    xml = z.read(name).decode("utf-8", "replace")
                except KeyError:
                    continue
                if "CONTROL#" in xml:
                    info["control_placeholder"] = True
    d = docx.Document(io.BytesIO(data))
    info["tables"] = len(d.tables)
    if d.sections:
        sec = d.sections[0]
        info["page_twips"] = (int(sec.page_width.twips), int(sec.page_height.twips))
        info["orientation"] = "landscape" if sec.page_width > sec.page_height else "portrait"
    return info


def validate(data: bytes, expected_orientation: str | None = None) -> list[str]:
    """Return a list of contract problems. An empty list means render-safe."""
    issues: list[str] = []
    try:
        d = docx.Document(io.BytesIO(data))
    except Exception as exc:  # pragma: no cover - corrupt input
        return [f"Template cannot be opened as .docx: {exc}"]

    required = manifest().get("required_styles", [])
    style_ids = {s.style_id for s in d.styles}
    for style in required:
        if style not in style_ids:
            issues.append(f"Missing paragraph/table style: {style}")

    if not d.sections:
        issues.append("Template has no section properties.")
    else:
        sect = d.sections[0]._sectPr
        if sect.find(qn("w:titlePg")) is None:
            issues.append("Template must set a different first-page header (titlePg).")
        if expected_orientation:
            info = inspect(data)
            if info["orientation"] and info["orientation"] != expected_orientation:
                issues.append(
                    f"Template is {info['orientation']} but was registered as "
                    f"{expected_orientation}.")

    info = inspect(data)
    if info["tables"] == 0:
        issues.append("Template has no table to use as a table prototype.")
    if not info["control_placeholder"]:
        issues.append(
            f"Template headers do not contain the control placeholder "
            f"'{manifest().get('control_placeholder', 'CONTROL#')}'.")
    if not info["fonts_embedded"]:
        issues.append(
            "Template does not embed its brand font; install Onest on the host so "
            "PDF export does not silently substitute a fallback.")

    actual = hashlib.sha256(data).hexdigest()
    spec = available().get(expected_orientation or "")
    if spec and spec.sha256 and actual != spec.sha256:
        issues.append(
            f"Template hash does not match manifest ({actual[:12]}…); run "
            "tools/update_manifest.py after an intentional edit.")
    return issues


def font_files() -> list[Path]:
    return sorted(FONT_DIR.glob("*.ttf")) + sorted(FONT_DIR.glob("*.otf"))
