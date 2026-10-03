import pytest

from va_toolkit import cli


def test_help_lists_all_commands(capsys):
    with pytest.raises(SystemExit):
        cli.main(["--help"])
    out = capsys.readouterr().out
    for command in ("mail-merge", "schedule", "organize", "undo", "expenses", "brief", "demo"):
        assert command in out


def test_cli_end_to_end(tmp_path, samples, capsys):
    cli.main(["expenses", str(samples / "expenses/receipts.csv"), "--month", "2026-09", "--out", str(tmp_path)])
    cli.main(["brief", str(samples / "daily_brief/tasks.xlsx"), "--today", "2026-10-05", "--out", str(tmp_path)])
    cli.main(["organize", str(samples / "messy_downloads"), "--dest", str(tmp_path / "org"),
              "--log", str(tmp_path / "plan.csv"), "--dry-run"])
    out = capsys.readouterr().out
    assert "Expense report 2026-09" in out
    assert "3 overdue" in out
    assert "DRY RUN" in out
    assert (tmp_path / "expense_report_2026-09.xlsx").exists()
