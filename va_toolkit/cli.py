"""Command-line interface: ``python -m va_toolkit <command> [options]``."""

from __future__ import annotations

import argparse
import shutil
import sys
from collections import Counter
from datetime import date
from pathlib import Path

from . import __version__, daily_brief, expense_report, file_organizer, mail_merge, scheduler

SAMPLES = Path("samples")
OUTPUT = Path("output")
DEMO_DATE = date(2026, 10, 5)
DEMO_MONTH = "2026-09"
DEMO_SENDER = "Gabriel Valle <assistant@example.com>"
DEMO_ORGANIZER = "assistant@example.com"


def cmd_mail_merge(args: argparse.Namespace) -> int:
    results = mail_merge.run(args.contacts, args.template, args.out, args.sender)
    ready = [r for r in results if r.status == mail_merge.READY]
    print(f"Mail merge: {len(ready)} drafts written, {len(results) - len(ready)} contacts skipped -> {args.out}")
    for r in results:
        if r.issues:
            print(f"  row {r.row:>3} {r.name or '(no name)':<22} SKIPPED: {'; '.join(r.issues)}")
    return 0


def cmd_schedule(args: argparse.Namespace) -> int:
    meetings, unscheduled = scheduler.run(args.requests, args.availability, args.out, args.organizer, args.tz)
    tz = scheduler.ZoneInfo(args.tz)
    print(f"Scheduler: {len(meetings)} meetings booked, {len(unscheduled)} unscheduled -> {args.out}")
    for m in meetings:
        start, end = m.local(tz)
        a_start, _ = m.local(m.request.tz)
        print(f"  {start:%a %d %b %H:%M}-{end:%H:%M}  {m.request.attendee:<18} {m.request.topic:<28}"
              f" (their time {a_start:%a %H:%M} {m.request.tz.key})")
    for u in unscheduled:
        print(f"  NOT BOOKED  {u.request.attendee:<18} {u.request.topic:<28} {u.reason}")
    return 0


def cmd_organize(args: argparse.Namespace) -> int:
    log = Path(args.log)
    moves = file_organizer.organize(args.source, args.dest, log, dry_run=args.dry_run, copy=args.copy)
    mode = "DRY RUN (nothing moved)" if args.dry_run else ("copied" if args.copy else "moved")
    counts = Counter(m.category for m in moves)
    print(f"File organizer: {len(moves)} files {mode} -> {args.dest}")
    for category, n in sorted(counts.items()):
        print(f"  {category:<13} {n}")
    for m in moves:
        if m.duplicate_of:
            print(f"  duplicate: {m.source.name}  (same content as {m.duplicate_of.name})")
    print(f"  {'plan' if args.dry_run else 'undo log'}: {log}")
    return 0


def cmd_undo(args: argparse.Namespace) -> int:
    n = file_organizer.undo(args.log)
    print(f"Undo: {n} actions reverted")
    return 0


def cmd_expenses(args: argparse.Namespace) -> int:
    report = expense_report.run(args.receipts, args.month, args.out)
    print(f"Expense report {report.month}: {len(report.expenses)} receipts, total USD {report.total_usd:,.2f}")
    for cat, total in report.by_category.items():
        print(f"  {cat:<16} USD {total:>9,.2f}")
    flagged = report.flagged
    print(f"  {len(flagged)} items flagged:")
    for r in flagged.itertuples():
        print(f"    {r.date} {r.employee:<14} {r.category:<15} USD {r.amount_usd:>8,.2f}  {r.flags}")
    return 0


def cmd_brief(args: argparse.Namespace) -> int:
    today = date.fromisoformat(args.today) if args.today else date.today()
    brief = daily_brief.run(args.tasks, args.out, today, args.manager)
    print(f"Daily brief {today}: {brief.open_count} open, {len(brief.overdue)} overdue, "
          f"{len(brief.due_today)} due today, {len(brief.this_week)} this week, "
          f"{len(brief.blocked)} blocked -> {args.out}")
    return 0


def cmd_demo(args: argparse.Namespace) -> int:
    """Run every module on the bundled samples and write to output/."""
    fo_out = OUTPUT / "file_organizer"
    if fo_out.exists():
        shutil.rmtree(fo_out)
    steps = [
        (cmd_mail_merge, dict(contacts=SAMPLES / "mail_merge/contacts.xlsx",
                              template=SAMPLES / "mail_merge/template.txt",
                              out=OUTPUT / "mail_merge", sender=DEMO_SENDER)),
        (cmd_schedule, dict(requests=SAMPLES / "scheduler/meeting_requests.xlsx",
                            availability=SAMPLES / "scheduler/availability.xlsx",
                            out=OUTPUT / "scheduler", organizer=DEMO_ORGANIZER, tz=scheduler.DEFAULT_TZ)),
        (cmd_organize, dict(source=SAMPLES / "messy_downloads", dest=fo_out / "dry_run_preview", dry_run=True,
                            copy=True, log=fo_out / "dry_run_plan.csv")),
        (cmd_organize, dict(source=SAMPLES / "messy_downloads", dest=fo_out / "organized", dry_run=False,
                            copy=True, log=fo_out / "undo_log.csv")),
        (cmd_expenses, dict(receipts=SAMPLES / "expenses/receipts.csv", month=DEMO_MONTH,
                            out=OUTPUT / "expense_report")),
        (cmd_brief, dict(tasks=SAMPLES / "daily_brief/tasks.xlsx", out=OUTPUT / "daily_brief",
                         today=DEMO_DATE.isoformat(), manager="Laura")),
    ]
    for func, kwargs in steps:
        func(argparse.Namespace(**kwargs))
        print()
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="va_toolkit", description="Automations for everyday virtual-assistant tasks.")
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    sub = p.add_subparsers(dest="command", required=True)

    s = sub.add_parser("mail-merge", help="Create personalized .eml drafts from an Excel contact list (never sends).")
    s.add_argument("contacts", type=Path)
    s.add_argument("template", type=Path)
    s.add_argument("--out", type=Path, default=OUTPUT / "mail_merge")
    s.add_argument("--sender", default=DEMO_SENDER)
    s.set_defaults(func=cmd_mail_merge)

    s = sub.add_parser("schedule", help="Propose meeting slots and export .ics invites + Excel agenda.")
    s.add_argument("requests", type=Path)
    s.add_argument("availability", type=Path)
    s.add_argument("--out", type=Path, default=OUTPUT / "scheduler")
    s.add_argument("--organizer", default=DEMO_ORGANIZER)
    s.add_argument("--tz", default=scheduler.DEFAULT_TZ, help="Timezone of the availability sheet")
    s.set_defaults(func=cmd_schedule)

    s = sub.add_parser("organize", help="Sort a messy folder into category/month folders.")
    s.add_argument("source", type=Path)
    s.add_argument("--dest", type=Path, required=True)
    s.add_argument("--log", type=Path, default=Path("organize_log.csv"), help="Undo log (or plan, with --dry-run)")
    s.add_argument("--dry-run", action="store_true", help="Only write the plan; touch nothing")
    s.add_argument("--copy", action="store_true", help="Copy instead of move (keep originals)")
    s.set_defaults(func=cmd_organize)

    s = sub.add_parser("undo", help="Revert an organize run using its undo log.")
    s.add_argument("log", type=Path)
    s.set_defaults(func=cmd_undo)

    s = sub.add_parser("expenses", help="Monthly expense report workbook with policy flags.")
    s.add_argument("receipts", type=Path)
    s.add_argument("--month", required=True, help="YYYY-MM")
    s.add_argument("--out", type=Path, default=OUTPUT / "expense_report")
    s.set_defaults(func=cmd_expenses)

    s = sub.add_parser("brief", help="Daily brief (Markdown + HTML) from tasks.xlsx.")
    s.add_argument("tasks", type=Path)
    s.add_argument("--out", type=Path, default=OUTPUT / "daily_brief")
    s.add_argument("--today", help="YYYY-MM-DD (default: today)")
    s.add_argument("--manager", default="Manager")
    s.set_defaults(func=cmd_brief)

    s = sub.add_parser("demo", help="Run every module on the bundled samples.")
    s.set_defaults(func=cmd_demo)
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
