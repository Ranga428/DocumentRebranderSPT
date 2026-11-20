"""Recompute template hashes after an intentional template edit."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
MANIFEST = ROOT / "templates" / "manifest.json"


def main() -> None:
    data = json.loads(MANIFEST.read_text(encoding="utf-8"))
    for key, entry in data["templates"].items():
        path = ROOT / "templates" / entry["file"]
        if not path.exists():
            print(f"{key}: missing {path.name}")
            continue
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        entry["sha256"] = digest
        print(f"{key}: {digest}")
    MANIFEST.write_text(json.dumps(data, indent=2) + "\n", encoding="utf-8")
    print(f"updated {MANIFEST}")


if __name__ == "__main__":
    main()
