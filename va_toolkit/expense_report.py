"""Expense report: monthly Excel report from a CSV of receipts, with policy checks.

CSV columns: date, employee, category, vendor, description, amount, currency, receipt_id

Policy checks (see :class:`ExpensePolicy`):
  * meals above the daily cap per employee (default USD 25/day),
  * any single item above the approval threshold (default USD 500),
  * missing receipt number,
  * possible duplicates (same employee, date, vendor and amount),
  * unsupported currency.

The workbook has three sheets: *Summary* (live SUMIF/COUNTIF formulas plus a bar
chart), *Expenses* (every line with a USD formula column, flagged rows in red)
and *Policy flags*.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill

from .excel_utils import autofit_columns, style_header, write_table

USD_FMT = '"$"#,##0.00'
FLAG_FILL = PatternFill("solid", fgColor="F8CBAD")
TOTAL_FONT = Font(bold=True)


@dataclass
class ExpensePolicy:
    """Company expense policy. FX rates are fixed, illustrative values."""

    meals_daily_cap_usd: float = 25.0
    approval_threshold_usd: float = 500.0
    meal_category: str = "Meals"
    fx_to_usd: dict[str, float] = field(default_factory=lambda: {"USD": 1.0, "EUR": 1.08, "ARS": 0.001})


@dataclass
class ExpenseReport:
    month: str
    expenses: pd.DataFrame  # one row per receipt, with amount_usd and flags
    by_category: pd.Series
    total_usd: float

    @property
    def flagged(self) -> pd.DataFrame:
        return self.expenses[self.expenses["flags"].str.len() > 0]


def load_receipts(path: str | Path) -> pd.DataFrame:
    """Read the receipts CSV and normalise types."""
    df = pd.read_csv(path, dtype={"receipt_id": "string"})
    df.columns = [c.strip().lower() for c in df.columns]
    df["date"] = pd.to_datetime(df["date"]).dt.date
    df["currency"] = df["currency"].fillna("USD").str.strip().str.upper()
    df["receipt_id"] = df["receipt_id"].fillna("").str.strip()
    for col in ("employee", "category", "vendor", "description"):
        df[col] = df[col].fillna("").astype(str).str.strip()
    df["amount"] = pd.to_numeric(df["amount"], errors="raise").astype(float)
    return df


def apply_policy(df: pd.DataFrame, policy: ExpensePolicy | None = None) -> pd.DataFrame:
    """Return a copy with ``rate``, ``amount_usd`` and a list of ``flags`` per row."""
    policy = policy or ExpensePolicy()
    df = df.copy().reset_index(drop=True)
    df["rate"] = df["currency"].map(policy.fx_to_usd)
    df["amount_usd"] = (df["amount"] * df["rate"].fillna(0)).round(2)
    flags: list[list[str]] = [[] for _ in range(len(df))]

    for i, row in df.iterrows():
        if pd.isna(row["rate"]):
            flags[i].append(f"unsupported currency {row['currency']}")
        if not row["receipt_id"]:
            flags[i].append("missing receipt")
        if row["amount_usd"] > policy.approval_threshold_usd:
            flags[i].append(f"over USD {policy.approval_threshold_usd:,.0f}: needs manager approval")

    meals = df[df["category"].str.lower() == policy.meal_category.lower()]
    daily = meals.groupby(["employee", "date"])["amount_usd"].transform("sum")
    for i, day_total in daily.items():
        if day_total > policy.meals_daily_cap_usd + 1e-9:
            flags[i].append(f"meals USD {day_total:,.2f} that day > USD {policy.meals_daily_cap_usd:,.0f}/day cap")

    dup_keys = ["employee", "date", "vendor", "amount"]
    for i in df.index[df.duplicated(dup_keys, keep="first")]:
        flags[i].append("possible duplicate")

    df["flags"] = ["; ".join(f) for f in flags]
    return df


def build_report(df: pd.DataFrame, month: str, policy: ExpensePolicy | None = None) -> ExpenseReport:
    """Filter receipts to ``month`` (``YYYY-MM``) and apply the policy."""
    in_month = df[[d.strftime("%Y-%m") == month for d in df["date"]]]
    if in_month.empty:
        raise ValueError(f"No receipts found for {month}")
    checked = apply_policy(in_month.sort_values(["date", "employee"]), policy)
    by_category = checked.groupby("category")["amount_usd"].sum().round(2).sort_values(ascending=False)
    return ExpenseReport(month, checked, by_category, round(float(checked["amount_usd"].sum()), 2))


def write_workbook(report: ExpenseReport, path: str | Path) -> Path:
    """Write the formatted three-sheet workbook."""
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    ws = wb.create_sheet("Expenses")
    flags_ws = wb.create_sheet("Policy flags")

    # --- Expenses sheet (data + per-row USD formula)
    headers = ["Date", "Employee", "Category", "Vendor", "Description", "Amount", "Currency",
               "Rate to USD", "Amount USD", "Receipt ID", "Status", "Policy notes"]
    ws.append(headers)
    for n, row in enumerate(report.expenses.itertuples(index=False), start=2):
        ws.append([row.date, row.employee, row.category, row.vendor, row.description, row.amount,
                   row.currency, None if pd.isna(row.rate) else row.rate, f"=ROUND(F{n}*H{n},2)",
                   row.receipt_id, "FLAGGED" if row.flags else "OK", row.flags])
        if row.flags:
            for cell in ws[n]:
                cell.fill = FLAG_FILL
    last = ws.max_row
    ws.append(["TOTAL", None, None, None, None, None, None, None, f"=SUM(I2:I{last})"])
    for cell in ws[ws.max_row]:
        cell.font = TOTAL_FONT
    for r in range(2, ws.max_row + 1):
        ws[f"A{r}"].number_format = "yyyy-mm-dd"
        ws[f"F{r}"].number_format = "#,##0.00"
        ws[f"I{r}"].number_format = USD_FMT
    style_header(ws)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:L{last}"
    autofit_columns(ws)

    # --- Summary sheet (formulas reference the Expenses sheet so edits flow through)
    def col(letter: str) -> str:
        return f"Expenses!${letter}$2:${letter}${last}"

    summary["A1"] = f"Expense report — {report.month}"
    summary["A1"].font = Font(bold=True, size=14)
    summary["A2"] = "Synthetic demo data · amounts in USD · formulas recalculate in Excel"
    summary["A2"].font = Font(italic=True, color="7F7F7F")
    summary.append([])
    summary.append(["Category", "Items", "Total USD", "Flagged items"])
    style_header(summary, row=4)
    first_cat = summary.max_row + 1
    for cat in report.by_category.index:
        r = summary.max_row + 1
        summary.append([
            cat,
            f'=COUNTIF({col("C")},A{r})',
            f'=SUMIF({col("C")},A{r},{col("I")})',
            f'=COUNTIFS({col("C")},A{r},{col("K")},"FLAGGED")',
        ])
    last_cat = summary.max_row
    summary.append(["TOTAL", f"=SUM(B{first_cat}:B{last_cat})", f"=SUM(C{first_cat}:C{last_cat})",
                    f"=SUM(D{first_cat}:D{last_cat})"])
    for cell in summary[summary.max_row]:
        cell.font = TOTAL_FONT
    for r in range(first_cat, summary.max_row + 1):
        summary[f"C{r}"].number_format = USD_FMT

    summary.append([])
    summary.append(["Employee", "Items", "Total USD", "Flagged items"])
    style_header(summary, row=summary.max_row)
    for emp in sorted(report.expenses["employee"].unique()):
        r = summary.max_row + 1
        summary.append([
            emp,
            f'=COUNTIF({col("B")},A{r})',
            f'=SUMIF({col("B")},A{r},{col("I")})',
            f'=COUNTIFS({col("B")},A{r},{col("K")},"FLAGGED")',
        ])
        summary[f"C{r}"].number_format = USD_FMT
    autofit_columns(summary)
    summary.column_dimensions["A"].width = 22

    chart = BarChart()
    chart.type = "bar"
    chart.title = f"Spend by category — {report.month}"
    chart.x_axis.title = None
    chart.y_axis.title = "USD"
    chart.legend = None
    chart.add_data(Reference(summary, min_col=3, min_row=4, max_row=last_cat), titles_from_data=True)
    chart.set_categories(Reference(summary, min_col=1, min_row=first_cat, max_row=last_cat))
    chart.height, chart.width = 8, 16
    summary.add_chart(chart, "F4")

    # --- Policy flags sheet
    flagged = report.flagged
    write_table(
        flags_ws,
        ["Date", "Employee", "Category", "Vendor", "Amount USD", "Issue(s)"],
        ([r.date, r.employee, r.category, r.vendor, r.amount_usd, r.flags] for r in flagged.itertuples()),
    )
    for r in range(2, flags_ws.max_row + 1):
        flags_ws[f"A{r}"].number_format = "yyyy-mm-dd"
        flags_ws[f"E{r}"].number_format = USD_FMT

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def run(csv_path: str | Path, month: str, out_dir: str | Path, policy: ExpensePolicy | None = None) -> ExpenseReport:
    """End-to-end: CSV -> ``expense_report_<month>.xlsx``."""
    report = build_report(load_receipts(csv_path), month, policy)
    write_workbook(report, Path(out_dir) / f"expense_report_{month}.xlsx")
    return report
