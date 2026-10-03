# Virtual Assistant Toolkit

[![CI](https://github.com/gabrielvalle-491/virtual-assistant-toolkit/actions/workflows/ci.yml/badge.svg)](https://github.com/gabrielvalle-491/virtual-assistant-toolkit/actions/workflows/ci.yml)
![Python](https://img.shields.io/badge/python-3.11%20%7C%203.12-blue)
![License](https://img.shields.io/badge/license-MIT-green)

**English** · [Español](README.es.md)

A Python toolkit that automates five everyday virtual-assistant tasks: email drafts, meeting
scheduling across timezones, file organization, expense reports and a daily task brief.
Everything runs from one command line (`python -m va_toolkit <command>`), reads the Excel/CSV
files a VA already works with, and writes files a manager can open straight away.

> Portfolio project. All names, companies, emails and amounts are **synthetic** and were produced
> by `generate_samples.py`. This is not client work.

## The business problem

A large part of a virtual assistant's week is repetitive admin work that is slow and easy to get wrong
by hand:

- writing the same payment reminder to a dozen clients, changing the name, amount and date each time;
- finding meeting times when the manager is in Buenos Aires and the attendees are in New York, London and Sydney;
- cleaning up a Downloads folder full of invoices, screenshots and "file (1).pdf" duplicates;
- turning a month of receipts into an expense report and checking each line against the expense policy;
- telling the manager every morning what is overdue, what is due today and what is blocked.

This toolkit does each of those in seconds, and every module **validates its input first** (invalid
emails, missing fields, impossible time slots, out-of-policy expenses), so the assistant reviews the
exceptions instead of re-checking everything.

## Features

| Module | Input | Output | Highlights |
|---|---|---|---|
| `mail_merge` | Contact list (`.xlsx`) + Jinja2 template | One `.eml` draft per contact + `mail_merge_summary.xlsx` | Flags invalid/duplicate emails and missing fields; drafts open as editable drafts (`X-Unsent: 1`); **never sends email** |
| `scheduler` | Meeting requests (`.xlsx`) + availability (`.xlsx`) | `.ics` invite per meeting, combined `agenda.ics`, `agenda.xlsx` | Timezone-aware with `zoneinfo` (incl. DST), preferred weekdays and local hours checked in the attendee's timezone, no overlaps + 15-min buffer, explains why a request could not be booked |
| `file_organizer` | A messy folder | `<category>/<YYYY-MM>/` folders, undo log (`.csv`) | Invoices detected by name (invoice/factura/receipt), month from file name or modified date, SHA-256 duplicate detection, `--dry-run`, `--copy`, full `undo` |
| `expense_report` | Receipts (`.csv`) | `expense_report_YYYY-MM.xlsx` (3 sheets) | Live `SUMIF`/`COUNTIFS` formulas, bar chart, currency conversion, flags meals > USD 25/day, items > USD 500, missing receipts and duplicates |
| `daily_brief` | Tasks (`.xlsx`) | `daily_brief_YYYY-MM-DD.md` + `.html` | Overdue / due today / later this week / blocked, sorted by priority then due date |

## Quick start

```bash
git clone https://github.com/gabrielvalle-491/virtual-assistant-toolkit.git
cd virtual-assistant-toolkit
pip install -r requirements.txt

python generate_samples.py          # (optional) rebuild the synthetic data in samples/
python -m va_toolkit demo           # run all five modules on samples/ -> output/
pytest                              # 53 tests
```

Run a single module on your own files:

```bash
python -m va_toolkit mail-merge contacts.xlsx template.txt --out output/mail_merge --sender "Me <me@example.com>"
python -m va_toolkit schedule meeting_requests.xlsx availability.xlsx --tz America/Argentina/Buenos_Aires
python -m va_toolkit organize ~/Downloads --dest ~/Organized --dry-run --log plan.csv   # preview only
python -m va_toolkit organize ~/Downloads --dest ~/Organized --log undo_log.csv        # do it
python -m va_toolkit undo undo_log.csv                                                 # changed your mind
python -m va_toolkit expenses receipts.csv --month 2026-09
python -m va_toolkit brief tasks.xlsx --today 2026-10-05 --manager Laura
```

## Real sample output

Everything below was copied from `python -m va_toolkit demo` on the bundled sample data. The generated
files are committed in [`output/`](output/).

### 1. Mail merge — payment reminders

```text
Mail merge: 7 drafts written, 5 contacts skipped -> output/mail_merge
  row   6 Tom Becker             SKIPPED: invalid email 'tom.becker@example'
  row   7 Nakamura               SKIPPED: missing field(s): first_name
  row   8 Sofía Ruiz             SKIPPED: missing field(s): amount_due
  row  10 Ana Torres             SKIPPED: duplicate email
  row  12 Javier Morales         SKIPPED: missing email
```

One of the drafts ([`output/mail_merge/drafts/001_ana_torres.eml`](output/mail_merge/drafts/001_ana_torres.eml)):

```text
From: Gabriel Valle <assistant@example.com>
To: ana.torres@example.com
Subject: Friendly reminder: invoice INV-2026-101 for Torres Design Studio
X-Unsent: 1

Hi Ana,

I hope you are having a good week. This is a friendly reminder that invoice
INV-2026-101 for USD 1,250.00 is due on October 15, 2026.
...
```

### 2. Scheduler — 8 requests, 6 timezones

Availability is Mon 5 – Fri 9 Oct 2026, 09:00–12:00 and 14:00–18:00 Buenos Aires time.

```text
Scheduler: 6 meetings booked, 2 unscheduled -> output/scheduler
  Mon 05 Oct 10:00-11:00  Olivia Carter      Quarterly review             (their time Mon 09:00 America/New_York)
  Mon 05 Oct 11:15-11:45  Marta López        Marketing sync               (their time Mon 16:15 Europe/Madrid)
  Mon 05 Oct 14:00-14:30  Valentina Rossi    Invoice reconciliation       (their time Mon 14:00 America/Argentina/Buenos_Aires)
  Mon 05 Oct 14:45-15:15  Carlos Méndez      Weekly 1:1                   (their time Mon 11:45 America/Mexico_City)
  Tue 06 Oct 09:00-09:45  James Patel        Supplier onboarding call     (their time Tue 13:00 Europe/London)
  Thu 08 Oct 14:00-15:30  Ryan Brooks        Website redesign kickoff     (their time Thu 10:00 America/Los_Angeles)
  NOT BOOKED  Hannah Lee         Partnership intro            no availability overlaps the attendee's preferred days/hours
  NOT BOOKED  Olivia Carter      Budget follow-up             all matching slots already booked
```

Hannah is in Sydney: her 09:00–17:00 is 19:00–03:00 in Buenos Aires, so the tool reports it instead of
booking her at 3 a.m. Each booked meeting also has an `.ics` invite in
[`output/scheduler/invites/`](output/scheduler/invites/) that imports into Google Calendar or Outlook.

### 3. File organizer — 16 messy files

```text
File organizer: 16 files copied -> output/file_organizer/organized
  archives      1
  docs          3
  duplicate     2
  images        3
  invoices      3
  media         1
  other         1
  spreadsheets  2
  duplicate: IMG_20260915_101233 - Copy.jpg  (same content as IMG_20260915_101233.jpg)
  duplicate: Invoice_ACME_2026-09-14 (1).pdf  (same content as Invoice_ACME_2026-09-14.pdf)
  undo log: output/file_organizer/undo_log.csv
```

Resulting tree:

```text
organized/
├── _duplicates/       IMG_20260915_101233 - Copy.jpg, Invoice_ACME_2026-09-14 (1).pdf
├── archives/2026-07/  archive_2026-07-10.zip
├── docs/2026-09/      meeting_notes_2026-09-22.docx, project_brief.pdf, setup_notes.txt
├── images/2026-09/    IMG_20260915_101233.jpg, Screenshot 2026-09-30 at 11.02.png, company_logo.png
├── invoices/2026-09/  Invoice_ACME_2026-09-14.pdf, factura_0042_20260902.pdf, receipt-taxi-2026_09_21.pdf
├── media/2026-09/     voice_memo_20260918.m4a
├── other/2026-09/     mystery_file.xyz
├── spreadsheets/2026-08/ contacts_export_2026-08-28.csv
└── spreadsheets/2026-09/ Q3_budget_2026-09-01.xlsx
```

(The demo uses `--copy` so `samples/` stays intact; the default is to move. Files without a date in their
name use the modified date, which `generate_samples.py` sets to September 2026.)

### 4. Expense report — September 2026

```text
Expense report 2026-09: 39 receipts, total USD 1,712.52
  Travel           USD    612.00
  Lodging          USD    368.00
  Meals            USD    248.07
  Office Supplies  USD    231.73
  Transport        USD    202.30
  Software         USD     50.42
  5 items flagged:
    2026-09-16 Peter Novak    Meals           USD    21.40  meals USD 31.20 that day > USD 25/day cap
    2026-09-16 Peter Novak    Meals           USD     9.80  meals USD 31.20 that day > USD 25/day cap
    2026-09-18 Laura Gómez    Office Supplies USD    42.30  possible duplicate
    2026-09-22 Aisha Bello    Travel          USD   612.00  over USD 500: needs manager approval
    2026-09-23 Aisha Bello    Meals           USD    31.00  missing receipt; meals USD 31.00 that day > USD 25/day cap
```

The workbook ([`output/expense_report/expense_report_2026-09.xlsx`](output/expense_report/expense_report_2026-09.xlsx))
has a **Summary** sheet with `SUMIF`/`COUNTIFS` formulas and a bar chart, an **Expenses** sheet with a
`=ROUND(Amount*Rate,2)` USD column and flagged rows in red, and a **Policy flags** sheet. ARS and EUR receipts
are converted with fixed, illustrative rates that you can change in `ExpensePolicy`.

### 5. Daily brief — Monday 5 October 2026

```text
Daily brief 2026-10-05: 16 open, 3 overdue, 3 due today, 6 this week, 2 blocked -> output/daily_brief
```

HTML version ([`output/daily_brief/daily_brief_2026-10-05.html`](output/daily_brief/daily_brief_2026-10-05.html)),
screenshot taken with headless Chromium:

![Daily brief screenshot](docs/daily_brief.png)

Markdown version (excerpt of [`daily_brief_2026-10-05.md`](output/daily_brief/daily_brief_2026-10-05.md)):

```markdown
## Overdue (3)

| Priority | Task | Owner | Due | Notes |
|---|---|---|---|---|
| High | T-001 Send September invoices to clients | Gabriel | 3 days late (due Fri 02 Oct) | 3 of 12 still pending |
| High | T-002 Renew domain example-shop.com | Laura | 1 day late (due Sun 04 Oct) | Expires Oct 6 |
| Medium | T-003 Book flights for Madrid trip | Gabriel | 2 days late (due Sat 03 Oct) | Waiting for final dates |
```

## Tests

```text
$ pytest -q
.....................................................                    [100%]
53 passed in 0.62s
```

The tests cover email validation, template parsing, DST-correct timezone conversion, preferred
weekdays checked in the attendee's own timezone, overlap/buffer rules, ICS formatting (escaping,
75-octet line folding), dry-run safety, duplicate detection, move/copy + undo round trips, the meals
daily cap, approval threshold, currency conversion, the workbook formulas and chart, the brief's date
buckets and HTML escaping, and the CLI. CI runs `ruff` and `pytest` on Python 3.11 and 3.12.

## Project structure

```text
virtual-assistant-toolkit/
├── va_toolkit/
│   ├── cli.py              # python -m va_toolkit <command>
│   ├── excel_utils.py      # read sheets as dicts, header styling, auto-fit
│   ├── mail_merge.py
│   ├── scheduler.py
│   ├── file_organizer.py
│   ├── expense_report.py
│   └── daily_brief.py
├── samples/                # synthetic input data (generate_samples.py)
├── output/                 # real output from `python -m va_toolkit demo`
├── tests/                  # pytest suite (53 tests)
├── docs/daily_brief.png
├── generate_samples.py
├── requirements.txt
└── .github/workflows/ci.yml
```

## Tech stack

Python 3.11+, openpyxl (Excel read/write, formulas, charts), pandas (expense aggregation), Jinja2
(email templates), `zoneinfo` (timezones), `email` (RFC 5322 drafts), hand-written RFC 5545 `.ics`
export, pytest, ruff, GitHub Actions. Built with the help of Claude Code.

## Author

Gabriel Valle — Virtual Assistant & Automation · Villa Mercedes, Argentina · Remote

Licensed under the [MIT License](LICENSE).
