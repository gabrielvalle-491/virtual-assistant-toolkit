"""Mail merge: personalized email drafts from an Excel contact list.

The template is a text file whose first line is ``Subject: ...`` followed by a
blank line and the body. Placeholders use Jinja2 syntax (``{{ first_name }}``).

Every contact is validated before rendering:
  * the email address must look valid and must not be a duplicate,
  * every field used by the template must be present and non-empty.

Valid contacts become ``.eml`` drafts (marked ``X-Unsent: 1`` so Outlook and
Thunderbird open them as editable drafts). Nothing is ever sent.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from email.message import EmailMessage
from pathlib import Path
from typing import Any

from jinja2 import Environment, StrictUndefined, meta
from openpyxl import Workbook
from openpyxl.styles import PatternFill

from .excel_utils import read_table, write_table

EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}$")
_ENV = Environment(undefined=StrictUndefined, keep_trailing_newline=True, autoescape=False)
_ENV.filters["money"] = lambda v: f"{float(v):,.2f}"
_ENV.filters["date"] = lambda v, fmt="%B %d, %Y": v.strftime(fmt) if isinstance(v, (date, datetime)) else str(v)

READY = "READY"
SKIPPED = "SKIPPED"


@dataclass
class DraftResult:
    """Outcome for one contact row."""

    row: int
    name: str
    email: str
    status: str
    issues: list[str] = field(default_factory=list)
    subject: str = ""
    file: str = ""


@dataclass
class EmailTemplate:
    """A parsed subject + body template."""

    subject: str
    body: str

    @property
    def fields(self) -> set[str]:
        """Placeholder names used anywhere in the template."""
        names: set[str] = set()
        for source in (self.subject, self.body):
            names |= meta.find_undeclared_variables(_ENV.parse(source))
        return names


def parse_template(text: str) -> EmailTemplate:
    """Split template text into subject and body. The first line must be ``Subject: ...``."""
    first, _, rest = text.partition("\n")
    if not first.lower().startswith("subject:"):
        raise ValueError("Template must start with a 'Subject:' line")
    return EmailTemplate(subject=first.split(":", 1)[1].strip(), body=rest.lstrip("\n"))


def load_template(path: str | Path) -> EmailTemplate:
    """Read and parse a template file."""
    return parse_template(Path(path).read_text(encoding="utf-8"))


def is_valid_email(value: Any) -> bool:
    """Return True when ``value`` looks like a deliverable email address."""
    return isinstance(value, str) and bool(EMAIL_RE.match(value.strip()))


def _is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def validate_contact(contact: dict[str, Any], required: set[str], seen_emails: set[str]) -> list[str]:
    """Return a list of human-readable problems for one contact (empty list = OK)."""
    issues: list[str] = []
    email = contact.get("email")
    if _is_blank(email):
        issues.append("missing email")
    elif not is_valid_email(email):
        issues.append(f"invalid email '{email}'")
    elif str(email).strip().lower() in seen_emails:
        issues.append("duplicate email")
    missing = sorted(f for f in required if f != "email" and _is_blank(contact.get(f)))
    if missing:
        issues.append("missing field(s): " + ", ".join(missing))
    return issues


def build_message(sender: str, to: str, subject: str, body: str) -> EmailMessage:
    """Build an unsent draft message."""
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = to
    msg["Subject"] = subject
    msg["X-Unsent"] = "1"
    msg.set_content(body, cte="8bit")
    return msg


def _slug(text: str) -> str:
    ascii_text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "_", ascii_text.lower()).strip("_") or "contact"


def merge(
    contacts: list[dict[str, Any]],
    template: EmailTemplate,
    out_dir: str | Path,
    sender: str,
) -> list[DraftResult]:
    """Validate every contact and write one ``.eml`` draft per valid row."""
    drafts_dir = Path(out_dir) / "drafts"
    drafts_dir.mkdir(parents=True, exist_ok=True)
    for old in drafts_dir.glob("*.eml"):  # keep the folder in sync with this run
        old.unlink()
    required = template.fields
    subject_tpl = _ENV.from_string(template.subject)
    body_tpl = _ENV.from_string(template.body)
    seen: set[str] = set()
    results: list[DraftResult] = []

    for idx, contact in enumerate(contacts, start=2):  # row 1 is the header
        name = " ".join(str(contact.get(k) or "").strip() for k in ("first_name", "last_name")).strip()
        email = str(contact.get("email") or "").strip()
        issues = validate_contact(contact, required, seen)
        if issues:
            results.append(DraftResult(idx, name, email, SKIPPED, issues))
            continue
        seen.add(email.lower())
        context = {k: (v.strip() if isinstance(v, str) else v) for k, v in contact.items()}
        subject = subject_tpl.render(**context).strip()
        body = body_tpl.render(**context)
        filename = f"{idx - 1:03d}_{_slug(name)}.eml"
        (drafts_dir / filename).write_bytes(bytes(build_message(sender, email, subject, body)))
        results.append(DraftResult(idx, name, email, READY, [], subject, f"drafts/{filename}"))
    return results


def write_summary(results: list[DraftResult], path: str | Path) -> Path:
    """Write a summary workbook: one row per contact with status and issues."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Mail merge"
    write_table(
        ws,
        ["Row", "Name", "Email", "Status", "Issues", "Subject", "Draft file"],
        ([r.row, r.name, r.email, r.status, "; ".join(r.issues), r.subject, r.file] for r in results),
    )
    red = PatternFill("solid", fgColor="F8CBAD")
    for row in ws.iter_rows(min_row=2):
        if row[3].value == SKIPPED:
            for cell in row:
                cell.fill = red
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def run(contacts_path: str | Path, template_path: str | Path, out_dir: str | Path, sender: str) -> list[DraftResult]:
    """End-to-end: read contacts, render drafts, write ``mail_merge_summary.xlsx``."""
    results = merge(read_table(contacts_path), load_template(template_path), out_dir, sender)
    write_summary(results, Path(out_dir) / "mail_merge_summary.xlsx")
    return results
