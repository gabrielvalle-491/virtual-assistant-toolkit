from datetime import date

from va_toolkit import daily_brief as db

TODAY = date(2026, 10, 7)  # Wednesday


def t(tid, due, status="Open", priority="Medium", title="Task"):
    return db.Task(tid, title, "Gabriel", priority, status, due)


def test_buckets_relative_to_today():
    brief = db.build_brief([
        t("late", date(2026, 10, 6)),
        t("today", TODAY),
        t("sunday", date(2026, 10, 11)),
        t("next-week", date(2026, 10, 12)),
        t("done", date(2026, 10, 1), status="Done"),
        t("nodate", None),
    ], TODAY)
    assert [x.task_id for x in brief.overdue] == ["late"]
    assert [x.task_id for x in brief.due_today] == ["today"]
    assert [x.task_id for x in brief.this_week] == ["sunday"]
    assert brief.open_count == 5 and brief.done_count == 1


def test_sorted_by_priority_then_due():
    brief = db.build_brief([
        t("low", date(2026, 10, 1), priority="Low"),
        t("high-later", date(2026, 10, 5), priority="High"),
        t("high-earlier", date(2026, 10, 2), priority="High"),
    ], TODAY)
    assert [x.task_id for x in brief.overdue] == ["high-earlier", "high-later", "low"]


def test_blocked_section():
    brief = db.build_brief([t("b", date(2026, 10, 20), status="Blocked")], TODAY)
    assert [x.task_id for x in brief.blocked] == ["b"]


def test_html_escapes_user_text():
    brief = db.build_brief([t("x", TODAY, title="<script>alert(1)</script>")], TODAY)
    out = db.to_html(brief)
    assert "<script>alert" not in out and "&lt;script&gt;" in out


def test_markdown_contains_counts_and_tables():
    brief = db.build_brief([t("a", date(2026, 10, 5), priority="High", title="Pay | rent")], TODAY)
    md = db.to_markdown(brief, "Laura")
    assert "Good morning Laura" in md
    assert "## Overdue (1)" in md and "2 days late" in md
    assert "Pay \\| rent" in md


def test_run_on_samples(tmp_path, samples):
    brief = db.run(samples / "daily_brief/tasks.xlsx", tmp_path, date(2026, 10, 5), "Laura")
    assert (len(brief.overdue), len(brief.due_today), len(brief.this_week), len(brief.blocked)) == (3, 3, 6, 2)
    assert brief.done_count == 3
    assert (tmp_path / "daily_brief_2026-10-05.md").exists()
    assert (tmp_path / "daily_brief_2026-10-05.html").read_text().startswith("<!doctype html>")
