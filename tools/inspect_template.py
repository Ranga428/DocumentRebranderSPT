"""Print template facts and contract problems.

    python tools/inspect_template.py
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core import template  # noqa: E402


def main() -> None:
    for key, spec in template.available().items():
        print(f"== {key}: {spec.path.name}")
        if not spec.path.exists():
            print("   MISSING")
            continue
        info = template.inspect(spec.read())
        print(f"   orientation: {info['orientation']}  page: {info['page_twips']} twips")
        print(f"   tables: {info['tables']}  embedded fonts: {info['fonts_embedded']}"
              f"  control placeholder: {info['control_placeholder']}")
        issues = template.validate(spec.read(), key)
        if issues:
            for issue in issues:
                print(f"   ISSUE: {issue}")
        else:
            print("   contract: OK")


if __name__ == "__main__":
    main()
