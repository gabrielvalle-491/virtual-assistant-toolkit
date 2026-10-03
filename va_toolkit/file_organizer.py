"""File organizer: tidy a messy downloads-style folder.

Each file is assigned a category (invoices, images, docs, spreadsheets, ...)
and a month folder, giving ``<dest>/<category>/<YYYY-MM>/<file>``:

  * category — keywords in the name win ("invoice", "factura", "receipt" ->
    invoices), otherwise the file extension decides;
  * month — a date in the file name (2026-09-14, 20260914, 2026_09_14) wins,
    otherwise the file's modification time is used.

Exact duplicates are found by SHA-256: the copy with the shortest name is
organized normally, the others go to ``<dest>/_duplicates/``. Every action is appended to an
undo log (CSV) so the whole run can be reverted with :func:`undo`.
``dry_run=True`` only returns/writes the plan and touches nothing.
"""

from __future__ import annotations

import csv
import hashlib
import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

CATEGORY_BY_EXT: dict[str, str] = {
    **dict.fromkeys([".jpg", ".jpeg", ".png", ".gif", ".webp", ".heic", ".svg"], "images"),
    **dict.fromkeys([".pdf", ".doc", ".docx", ".txt", ".md", ".odt", ".rtf", ".pptx"], "docs"),
    **dict.fromkeys([".xlsx", ".xls", ".csv", ".ods"], "spreadsheets"),
    **dict.fromkeys([".zip", ".rar", ".7z", ".tar", ".gz"], "archives"),
    **dict.fromkeys([".mp3", ".wav", ".mp4", ".mov", ".m4a"], "media"),
}
INVOICE_WORDS = ("invoice", "factura", "receipt", "recibo", "bill")
DATE_RE = re.compile(r"(20\d{2})[-_.]?(0[1-9]|1[0-2])[-_.]?(0[1-9]|[12]\d|3[01])")
LOG_FIELDS = ["timestamp", "action", "source", "destination", "dest_root", "category", "sha256"]
DUPLICATES = "_duplicates"


@dataclass(frozen=True)
class PlannedMove:
    source: Path
    destination: Path
    category: str
    sha256: str
    duplicate_of: Path | None = None


def categorize(path: Path) -> str:
    """Return the category folder name for a file."""
    name = path.name.lower()
    if any(word in name for word in INVOICE_WORDS):
        return "invoices"
    return CATEGORY_BY_EXT.get(path.suffix.lower(), "other")


def month_folder(path: Path) -> str:
    """``YYYY-MM`` from a date in the file name, else from the modification time."""
    match = DATE_RE.search(path.name)
    if match:
        return f"{match.group(1)}-{match.group(2)}"
    return datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m")


def file_hash(path: Path, chunk_size: int = 65536) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(chunk_size), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _unique(dest: Path, taken: set[Path]) -> Path:
    """Avoid overwriting: ``report.pdf`` -> ``report (1).pdf`` if needed."""
    candidate, n = dest, 1
    while candidate in taken or candidate.exists():
        candidate = dest.with_name(f"{dest.stem} ({n}){dest.suffix}")
        n += 1
    return candidate


def plan(source: str | Path, dest: str | Path) -> list[PlannedMove]:
    """Work out where every file in ``source`` (top level, non-hidden) should go."""
    source, dest = Path(source), Path(dest)
    files = sorted(p for p in source.iterdir() if p.is_file() and not p.name.startswith("."))
    hashes = {path: file_hash(path) for path in files}
    # Within a group of identical files keep the shortest name ("report.pdf"
    # rather than "report (1).pdf" or "report - Copy.pdf").
    originals: dict[str, Path] = {}
    for path in sorted(files, key=lambda p: (len(p.name), p.name)):
        originals.setdefault(hashes[path], path)
    taken: set[Path] = set()
    moves: list[PlannedMove] = []
    for path in files:
        digest = hashes[path]
        if originals[digest] != path:
            target = _unique(dest / DUPLICATES / path.name, taken)
            moves.append(PlannedMove(path, target, "duplicate", digest, originals[digest]))
        else:
            category = categorize(path)
            target = _unique(dest / category / month_folder(path) / path.name, taken)
            moves.append(PlannedMove(path, target, category, digest))
        taken.add(moves[-1].destination)
    return moves


def write_plan(moves: list[PlannedMove], path: str | Path) -> Path:
    """Write the plan as CSV (used for dry runs)."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh, lineterminator="\n")
        writer.writerow(["source", "destination", "category", "duplicate_of", "sha256"])
        for m in moves:
            writer.writerow([m.source.as_posix(), m.destination.as_posix(), m.category,
                             m.duplicate_of.name if m.duplicate_of else "", m.sha256])
    return path


def organize(
    source: str | Path,
    dest: str | Path,
    log_path: str | Path,
    dry_run: bool = False,
    copy: bool = False,
) -> list[PlannedMove]:
    """Organize ``source`` into ``dest``.

    With ``dry_run`` nothing is moved and the plan is written to ``log_path``
    instead. With ``copy`` the originals are left in place (handy for demos).
    """
    moves = plan(source, dest)
    if dry_run:
        write_plan(moves, log_path)
        return moves
    log_path = Path(log_path)
    log_path.parent.mkdir(parents=True, exist_ok=True)
    new_log = not log_path.exists()
    action = "copy" if copy else "move"
    with log_path.open("a", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=LOG_FIELDS, lineterminator="\n")
        if new_log:
            writer.writeheader()
        for m in moves:
            m.destination.parent.mkdir(parents=True, exist_ok=True)
            if copy:
                shutil.copy2(m.source, m.destination)
            else:
                shutil.move(str(m.source), str(m.destination))
            writer.writerow({
                "timestamp": datetime.now().isoformat(timespec="seconds"),
                "action": action,
                "source": str(m.source),
                "destination": str(m.destination),
                "dest_root": str(Path(dest)),
                "category": m.category,
                "sha256": m.sha256,
            })
    return moves


def undo(log_path: str | Path) -> int:
    """Revert every action in the log (newest first). Returns the number reverted.

    Moves are moved back; copies are deleted. Empty folders left behind are
    removed (never above the original destination root). The log is renamed
    to ``*.undone.csv`` so it cannot be applied twice.
    """
    log_path = Path(log_path)
    with log_path.open(newline="", encoding="utf-8") as fh:
        entries = list(csv.DictReader(fh))
    reverted = 0
    for entry in reversed(entries):
        src, dst = Path(entry["source"]), Path(entry["destination"])
        if not dst.exists():
            continue
        if entry["action"] == "move":
            src.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(dst), str(_unique(src, set())))
        else:
            dst.unlink()
        reverted += 1
        root = Path(entry["dest_root"]).resolve()
        parent = dst.parent.resolve()
        while parent != root and root in parent.parents and not any(parent.iterdir()):
            parent.rmdir()
            parent = parent.parent
    log_path.rename(log_path.with_suffix(".undone.csv"))
    return reverted
