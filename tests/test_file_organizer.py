import os
from datetime import datetime
from pathlib import Path

import pytest

from va_toolkit import file_organizer as fo


@pytest.fixture
def messy(tmp_path):
    src = tmp_path / "Downloads"
    src.mkdir()
    files = {
        "Invoice_ACME_2026-09-14.pdf": b"inv",
        "Invoice_ACME_2026-09-14 (1).pdf": b"inv",  # duplicate
        "photo.JPG": b"img",
        "budget_20260801.xlsx": b"xls",
        "notes.txt": b"txt",
        "weird.bin": b"bin",
        ".hidden": b"secret",
    }
    for name, data in files.items():
        (src / name).write_bytes(data)
    ts = datetime(2026, 5, 20).timestamp()
    for p in src.iterdir():
        os.utime(p, (ts, ts))
    return src


@pytest.mark.parametrize("name, expected", [
    ("factura_0042.pdf", "invoices"), ("Receipt-taxi.png", "invoices"), ("IMG_1.jpeg", "images"),
    ("report.DOCX", "docs"), ("data.csv", "spreadsheets"), ("backup.zip", "archives"), ("x.unknown", "other"),
])
def test_categorize(name, expected):
    assert fo.categorize(Path(name)) == expected


def test_month_from_filename_or_mtime(messy):
    assert fo.month_folder(messy / "budget_20260801.xlsx") == "2026-08"
    assert fo.month_folder(messy / "Invoice_ACME_2026-09-14.pdf") == "2026-09"
    assert fo.month_folder(messy / "notes.txt") == "2026-05"


def test_dry_run_touches_nothing(messy, tmp_path):
    before = sorted(p.name for p in messy.iterdir())
    moves = fo.organize(messy, tmp_path / "out", tmp_path / "plan.csv", dry_run=True)
    assert sorted(p.name for p in messy.iterdir()) == before
    assert not (tmp_path / "out").exists()
    assert len(moves) == 6  # hidden file ignored
    assert (tmp_path / "plan.csv").read_text().count("\n") == 7


def test_duplicates_detected_by_hash(messy, tmp_path):
    moves = fo.plan(messy, tmp_path / "out")
    dups = [m for m in moves if m.category == "duplicate"]
    assert len(dups) == 1
    assert dups[0].source.name == "Invoice_ACME_2026-09-14 (1).pdf"
    assert dups[0].duplicate_of.name == "Invoice_ACME_2026-09-14.pdf"
    assert dups[0].destination.parent.name == fo.DUPLICATES


def test_organize_moves_files_and_undo_restores(messy, tmp_path):
    dest, log = tmp_path / "out", tmp_path / "undo.csv"
    original = {p.name: p.read_bytes() for p in messy.iterdir()}
    fo.organize(messy, dest, log)
    assert (dest / "invoices/2026-09/Invoice_ACME_2026-09-14.pdf").exists()
    assert (dest / "images/2026-05/photo.JPG").exists()
    assert (dest / "spreadsheets/2026-08/budget_20260801.xlsx").exists()
    assert (dest / "other/2026-05/weird.bin").exists()
    assert sorted(p.name for p in messy.iterdir()) == [".hidden"]

    assert fo.undo(log) == 6
    assert {p.name: p.read_bytes() for p in messy.iterdir()} == original
    assert dest.exists() and not any(dest.iterdir())  # empty sub-folders cleaned, root kept
    assert not log.exists() and log.with_suffix(".undone.csv").exists()


def test_copy_mode_keeps_originals_and_undo_deletes_copies(messy, tmp_path):
    dest, log = tmp_path / "out", tmp_path / "undo.csv"
    fo.organize(messy, dest, log, copy=True)
    assert len(list(messy.iterdir())) == 7
    assert len([p for p in dest.rglob("*") if p.is_file()]) == 6
    fo.undo(log)
    assert not [p for p in dest.rglob("*") if p.is_file()]
    assert len(list(messy.iterdir())) == 7


def test_name_collision_gets_suffix(tmp_path):
    src = tmp_path / "src"
    src.mkdir()
    dest = tmp_path / "out"
    (dest / "docs/2026-09").mkdir(parents=True)
    (dest / "docs/2026-09/notes_2026-09-01.txt").write_text("existing")
    (src / "notes_2026-09-01.txt").write_text("new")
    moves = fo.plan(src, dest)
    assert moves[0].destination.name == "notes_2026-09-01 (1).txt"
