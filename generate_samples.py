"""Generate the synthetic sample data in samples/ (deterministic, no real people or companies).

Usage:  python generate_samples.py [--out samples]
"""

from __future__ import annotations

import argparse
import csv
import os
import random
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook

from va_toolkit.excel_utils import write_table

DEMO_MONDAY = date(2026, 10, 5)


def _save(path: Path, headers: list[str], rows: list[list]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    write_table(wb.active, headers, rows)
    wb.save(path)


# ----------------------------------------------------------------------------- mail merge
CONTACTS = [
    ["Ana", "Torres", "ana.torres@example.com", "Torres Design Studio", "INV-2026-101", 1250.00, date(2026, 10, 15)],
    ["Mark", "Hughes", "mark.hughes@example.org", "Hughes Logistics", "INV-2026-102", 480.50, date(2026, 10, 12)],
    ["Lucía", "Fernández", "lucia.fernandez@example.com", "Café Andino", "INV-2026-103", 320.00, date(2026, 10, 10)],
    ["Priya", "Shah", "priya.shah@example.net", "Shah Consulting", "INV-2026-104", 2100.00, date(2026, 10, 20)],
    ["Tom", "Becker", "tom.becker@example", "Becker Bikes", "INV-2026-105", 95.00, date(2026, 10, 9)],
    ["", "Nakamura", "k.nakamura@example.com", "Nakamura Imports", "INV-2026-106", 760.00, date(2026, 10, 18)],
    ["Sofía", "Ruiz", "sofia.ruiz@example.com", "Ruiz & Asociados", "INV-2026-107", None, date(2026, 10, 14)],
    ["Daniel", "Okafor", "daniel.okafor@example.com", "Okafor Fitness", "INV-2026-108", 640.00, date(2026, 10, 16)],
    ["Ana", "Torres", "Ana.Torres@example.com", "Torres Design Studio", "INV-2026-109", 300.00, date(2026, 10, 22)],
    ["Emma", "Wilson", "emma.wilson@example.co.uk", "Wilson Bakery", "INV-2026-110", 210.75, date(2026, 10, 11)],
    ["Javier", "Morales", "", "Morales Travel", "INV-2026-111", 1530.00, date(2026, 10, 25)],
    ["Chloé", "Martin", "chloe.martin@example.fr", "Atelier Martin", "INV-2026-112", 880.00, date(2026, 10, 19)],
]

TEMPLATE = """Subject: Friendly reminder: invoice {{ invoice_number }} for {{ company }}

Hi {{ first_name }},

I hope you are having a good week. This is a friendly reminder that invoice
{{ invoice_number }} for USD {{ amount_due | money }} is due on {{ due_date | date }}.

If the payment has already been sent, please disregard this message and thank you!
Otherwise, you can reply to this email with any questions or if you need a copy
of the invoice.

Kind regards,
Gabriel Valle
Virtual Assistant
"""


def mail_merge_samples(out: Path) -> None:
    _save(out / "mail_merge/contacts.xlsx",
          ["first_name", "last_name", "email", "company", "invoice_number", "amount_due", "due_date"], CONTACTS)
    (out / "mail_merge/template.txt").write_text(TEMPLATE, encoding="utf-8")


# ----------------------------------------------------------------------------- scheduler
REQUESTS = [
    ["REQ-01", "Olivia Carter", "olivia.carter@example.com", "Quarterly review", 60, "Mon, Tue", "America/New_York", "09:00", "17:00"],
    ["REQ-02", "James Patel", "james.patel@example.co.uk", "Supplier onboarding call", 45, "Tue, Wed", "Europe/London", "09:00", "17:30"],
    ["REQ-03", "Marta López", "marta.lopez@example.es", "Marketing sync", 30, "Mon, Wed, Fri", "Europe/Madrid", "10:00", "18:00"],
    ["REQ-04", "Ryan Brooks", "ryan.brooks@example.com", "Website redesign kickoff", 90, "Thu", "America/Los_Angeles", "08:00", "16:00"],
    ["REQ-05", "Valentina Rossi", "valentina.rossi@example.com", "Invoice reconciliation", 30, "Mon", "America/Argentina/Buenos_Aires", "09:00", "18:00"],
    ["REQ-06", "Hannah Lee", "hannah.lee@example.com.au", "Partnership intro", 30, "Tue, Thu", "Australia/Sydney", "09:00", "17:00"],
    ["REQ-07", "Carlos Méndez", "carlos.mendez@example.mx", "Weekly 1:1", 30, "Mon, Tue, Wed, Thu, Fri", "America/Mexico_City", "09:00", "17:00"],
    ["REQ-08", "Olivia Carter", "olivia.carter@example.com", "Budget follow-up", 30, "Mon", "America/New_York", "09:00", "12:00"],
]


def scheduler_samples(out: Path) -> None:
    _save(out / "scheduler/meeting_requests.xlsx",
          ["request_id", "attendee", "email", "topic", "duration_min", "preferred_days", "timezone", "earliest", "latest"],
          REQUESTS)
    rows = []
    for offset in range(5):  # Mon 5 Oct - Fri 9 Oct 2026, organizer time (Buenos Aires)
        day = DEMO_MONDAY + timedelta(days=offset)
        rows.append([day, "09:00", "12:00"])
        rows.append([day, "14:00", "18:00"])
    rows[0] = [DEMO_MONDAY, "10:00", "12:00"]  # Monday morning starts late (team stand-up)
    _save(out / "scheduler/availability.xlsx", ["date", "start", "end"], rows)


# ----------------------------------------------------------------------------- file organizer
MESSY_FILES = [
    "Invoice_ACME_2026-09-14.pdf", "factura_0042_20260902.pdf", "receipt-taxi-2026_09_21.pdf",
    "IMG_20260915_101233.jpg", "Screenshot 2026-09-30 at 11.02.png", "company_logo.png",
    "Q3_budget_2026-09-01.xlsx", "contacts_export_2026-08-28.csv", "meeting_notes_2026-09-22.docx",
    "project_brief.pdf", "archive_2026-07-10.zip", "voice_memo_20260918.m4a", "setup_notes.txt",
    "mystery_file.xyz",
]
DUPLICATES = {
    "Invoice_ACME_2026-09-14 (1).pdf": "Invoice_ACME_2026-09-14.pdf",
    "IMG_20260915_101233 - Copy.jpg": "IMG_20260915_101233.jpg",
}


def messy_folder_samples(out: Path) -> None:
    folder = out / "messy_downloads"
    folder.mkdir(parents=True, exist_ok=True)
    for name in MESSY_FILES:
        (folder / name).write_bytes(f"Synthetic placeholder file for the va_toolkit demo: {name}\n".encode())
    for copy_name, original in DUPLICATES.items():
        (folder / copy_name).write_bytes((folder / original).read_bytes())
    stamp = datetime(2026, 9, 10, 12, 0).timestamp()  # used by files without a date in the name
    for path in folder.iterdir():
        os.utime(path, (stamp, stamp))


# ----------------------------------------------------------------------------- expenses
EMPLOYEES = ["Laura Gómez", "Peter Novak", "Aisha Bello", "Martín Sosa"]
VENDORS = {
    "Meals": ["Green Bowl Café", "Deli 24", "Sushi Go", "La Parrilla"],
    "Transport": ["Uber", "City Taxi", "Metro card"],
    "Office Supplies": ["Staples", "OfficeMax"],
    "Software": ["Canva Pro", "Zoom", "Notion"],
    "Lodging": ["Hotel Plaza"],
    "Travel": ["Aerolíneas Argentinas", "Delta"],
}


def expense_samples(out: Path) -> None:
    rng = random.Random(42)
    rows: list[list] = []
    n = 1000
    month_days = [date(2026, 9, d) for d in range(1, 31) if date(2026, 9, d).weekday() < 5]

    def add(day: date, emp: str, cat: str, vendor: str, desc: str, amount: float, cur: str = "USD", rid: str | None = None):
        nonlocal n
        n += 1
        rows.append([day.isoformat(), emp, cat, vendor, desc, f"{amount:.2f}", cur, f"R-{n}" if rid is None else rid])

    meal_days = rng.sample([(e, d) for e in EMPLOYEES for d in month_days], 14)
    for emp, day in sorted(meal_days, key=lambda x: (x[1], x[0])):
        add(day, emp, "Meals", rng.choice(VENDORS["Meals"]), "Lunch with client", round(rng.uniform(8, 18), 2))
    for _ in range(10):
        add(rng.choice(month_days), rng.choice(EMPLOYEES), "Transport", rng.choice(VENDORS["Transport"]),
            "Ride to client office", round(rng.uniform(6, 30), 2))
    for _ in range(4):
        add(rng.choice(month_days), rng.choice(EMPLOYEES), "Office Supplies", rng.choice(VENDORS["Office Supplies"]),
            "Printer paper and pens", round(rng.uniform(15, 60), 2))
    add(date(2026, 9, 1), "Laura Gómez", "Software", "Canva Pro", "Monthly subscription", 14.99)
    add(date(2026, 9, 1), "Peter Novak", "Software", "Zoom", "Monthly subscription", 15.99)

    # Deliberate policy cases
    add(date(2026, 9, 16), "Peter Novak", "Meals", "La Parrilla", "Team dinner", 21.40)
    add(date(2026, 9, 16), "Peter Novak", "Meals", "Deli 24", "Coffee and snacks", 9.80)  # > 25/day together
    add(date(2026, 9, 22), "Aisha Bello", "Travel", "Delta", "Flight to Miami trade show", 612.00)  # > 500
    add(date(2026, 9, 23), "Aisha Bello", "Lodging", "Hotel Plaza", "Trade show hotel, 2 nights", 368.00)
    add(date(2026, 9, 23), "Aisha Bello", "Meals", "Sushi Go", "Dinner, trade show", 31.00, rid="")  # cap + no receipt
    add(date(2026, 9, 10), "Martín Sosa", "Transport", "City Taxi", "Airport transfer", 24000.00, cur="ARS")
    add(date(2026, 9, 18), "Laura Gómez", "Office Supplies", "OfficeMax", "Ink cartridges", 42.30)
    add(date(2026, 9, 18), "Laura Gómez", "Office Supplies", "OfficeMax", "Ink cartridges", 42.30)  # duplicate
    add(date(2026, 9, 25), "Martín Sosa", "Software", "Notion", "Team plan", 18.00, cur="EUR")
    # Outside the demo month (filtered out by --month 2026-09)
    add(date(2026, 8, 29), "Laura Gómez", "Meals", "Deli 24", "Lunch", 12.00)
    add(date(2026, 10, 1), "Peter Novak", "Transport", "Uber", "Ride to client", 11.50)

    rows.sort(key=lambda r: r[0])
    path = out / "expenses/receipts.csv"
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh, lineterminator="\n")
        writer.writerow(["date", "employee", "category", "vendor", "description", "amount", "currency", "receipt_id"])
        writer.writerows(rows)


# ----------------------------------------------------------------------------- daily brief
TASKS = [  # (title, owner, priority, status, due offset from Monday 5 Oct 2026, notes)
    ("Send September invoices to clients", "Gabriel", "High", "In Progress", -3, "3 of 12 still pending"),
    ("Renew domain example-shop.com", "Laura", "High", "Open", -1, "Expires Oct 6"),
    ("Book flights for Madrid trip", "Gabriel", "Medium", "Open", -2, "Waiting for final dates"),
    ("Prepare weekly KPI report", "Gabriel", "High", "Open", 0, "Use dashboard export"),
    ("Confirm catering for Thursday event", "Gabriel", "Medium", "Open", 0, ""),
    ("Reply to supplier quote (Hughes Logistics)", "Laura", "Medium", "In Progress", 0, ""),
    ("Update CRM with new leads", "Gabriel", "Low", "Open", 1, "From Friday's webinar"),
    ("Schedule Q4 planning meeting", "Laura", "High", "Open", 2, "Needs 4 attendees across 3 timezones"),
    ("Reconcile credit-card statement", "Gabriel", "Medium", "Open", 3, ""),
    ("Draft newsletter for October", "Peter", "Low", "Open", 4, "Topics in shared doc"),
    ("Archive Q3 project files", "Gabriel", "Low", "Open", 6, ""),
    ("Collect W-8BEN forms from contractors", "Laura", "High", "Blocked", 1, "Waiting on 2 contractors"),
    ("Order new office chairs", "Peter", "Low", "Blocked", 9, "Budget approval pending"),
    ("Onboard new client: Café Andino", "Gabriel", "Medium", "Open", 10, ""),
    ("Quarterly tax documents to accountant", "Laura", "High", "Open", 12, ""),
    ("Clean up shared drive permissions", "Gabriel", "Low", "Open", None, "Someday"),
    ("Send August expense report", "Gabriel", "High", "Done", -5, ""),
    ("Set up Zoom webinar", "Peter", "Medium", "Done", -1, ""),
    ("Book meeting room for Friday", "Gabriel", "Low", "Done", 0, ""),
]


def task_samples(out: Path) -> None:
    rows = []
    for i, (title, owner, prio, status, offset, notes) in enumerate(TASKS, start=1):
        due = DEMO_MONDAY + timedelta(days=offset) if offset is not None else None
        rows.append([f"T-{i:03d}", title, owner, prio, status, due, notes])
    _save(out / "daily_brief/tasks.xlsx", ["task_id", "title", "owner", "priority", "status", "due_date", "notes"], rows)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", type=Path, default=Path("samples"))
    out = parser.parse_args().out
    mail_merge_samples(out)
    scheduler_samples(out)
    messy_folder_samples(out)
    expense_samples(out)
    task_samples(out)
    print(f"Sample data written to {out}/")


if __name__ == "__main__":
    main()
