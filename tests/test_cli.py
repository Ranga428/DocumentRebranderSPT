"""CLI batch mode: several inputs in one run."""
import types

from core.export import ExportResult
from core.models import Block, DocumentModel
import cli


def _fake_spec():
    return types.SimpleNamespace(read=lambda: b"template", sha256="deadbeef")


def test_cli_processes_every_input(monkeypatch, tmp_path):
    calls = []

    def fake_ingest(name, data):
        calls.append(name)
        return DocumentModel(meta={"title": name}, blocks=[Block(kind="paragraph")])

    monkeypatch.setattr(cli, "ingest", fake_ingest)
    monkeypatch.setattr(cli.template, "choose", lambda model, orientation: ("portrait", ""))
    monkeypatch.setattr(cli.template, "get", lambda key: _fake_spec())
    monkeypatch.setattr(cli.template, "validate", lambda data, key: [])
    monkeypatch.setattr(cli, "build",
                        lambda model, data, cover: ExportResult(docx=b"docx-bytes"))

    one = tmp_path / "one.txt"
    two = tmp_path / "two.txt"
    one.write_text("first")
    two.write_text("second")
    outdir = tmp_path / "out"

    assert cli.main([str(one), str(two), "-o", str(outdir)]) == 0
    assert calls == ["one.txt", "two.txt"]
    assert sorted(p.name for p in outdir.iterdir()) == ["one-SPT.docx", "two-SPT.docx"]


def test_cli_keeps_going_when_one_file_fails(monkeypatch, tmp_path):
    from core.ingest import IngestError

    def fake_ingest(name, data):
        if name == "bad.txt":
            raise IngestError("unreadable")
        return DocumentModel(meta={"title": name}, blocks=[Block(kind="paragraph")])

    monkeypatch.setattr(cli, "ingest", fake_ingest)
    monkeypatch.setattr(cli.template, "choose", lambda model, orientation: ("portrait", ""))
    monkeypatch.setattr(cli.template, "get", lambda key: _fake_spec())
    monkeypatch.setattr(cli.template, "validate", lambda data, key: [])
    monkeypatch.setattr(cli, "build",
                        lambda model, data, cover: ExportResult(docx=b"docx-bytes"))

    good = tmp_path / "good.txt"
    bad = tmp_path / "bad.txt"
    good.write_text("ok")
    bad.write_text("nope")
    outdir = tmp_path / "out"

    assert cli.main([str(bad), str(good), "-o", str(outdir)]) == 1
    assert [p.name for p in outdir.iterdir()] == ["good-SPT.docx"]
