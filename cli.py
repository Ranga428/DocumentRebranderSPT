"""Batch/CLI entry point.

    python cli.py report.pdf -o out/
    python cli.py report.pdf notes.md policy.docx -o out/
    python cli.py notes.md --orientation landscape --no-cover
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from core import template
from core.export import build
from core.ingest import IngestError, ingest


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Rebrand one or more documents into the SPT template.")
    parser.add_argument("inputs", nargs="+", metavar="input",
                        help="Source document(s)")
    parser.add_argument("-o", "--outdir", default="out", help="Output directory")
    parser.add_argument("--orientation", default="auto",
                        choices=["auto", "portrait", "landscape"])
    parser.add_argument("--cover", dest="cover", action="store_true", default=True)
    parser.add_argument("--no-cover", dest="cover", action="store_false")
    args = parser.parse_args(argv)

    outdir = Path(args.outdir)
    failures = 0
    used_stems: set[str] = set()

    for raw in args.inputs:
        source = Path(raw)
        if not source.exists():
            print(f"error: {source} does not exist", file=sys.stderr)
            failures += 1
            continue

        try:
            model = ingest(source.name, source.read_bytes())
        except IngestError as exc:
            print(f"error: {source.name}: {exc}", file=sys.stderr)
            failures += 1
            continue

        key, warning = template.choose(model, args.orientation)
        if warning:
            print(f"warning: {source.name}: {warning}", file=sys.stderr)
        spec = template.get(key)
        issues = template.validate(spec.read(), key)
        if issues:
            print(f"template problems ({source.name}):", file=sys.stderr)
            for issue in issues:
                print(f"  - {issue}", file=sys.stderr)
            failures += 1
            continue

        result = build(model, spec.read(), cover=args.cover)

        outdir.mkdir(parents=True, exist_ok=True)
        stem = source.stem
        if stem in used_stems:
            stem = f"{stem}-{len(used_stems) + 1}"
        used_stems.add(stem)
        docx_path = outdir / f"{stem}-SPT.docx"
        docx_path.write_bytes(result.docx)
        print(f"wrote {docx_path}  ({key}, {len(model.blocks)} blocks)")
        if result.pdf:
            pdf_path = outdir / f"{stem}-SPT.pdf"
            pdf_path.write_bytes(result.pdf)
            print(f"wrote {pdf_path}  ({result.pages} pages)")
        for message in result.warnings:
            print(f"warning: {message}", file=sys.stderr)

    if len(args.inputs) > 1:
        print(f"processed {len(args.inputs) - failures} of {len(args.inputs)} file(s)")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
