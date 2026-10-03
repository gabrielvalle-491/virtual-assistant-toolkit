"""Scheduler: propose non-overlapping meeting slots across timezones.

Inputs
  * meeting requests (Excel): request_id, attendee, email, topic, duration_min,
    preferred_days (e.g. "Mon, Wed"), timezone (IANA name), earliest, latest
    (the attendee's acceptable local hours, "HH:MM").
  * availability (Excel): date, start, end — the assistant's/manager's free
    blocks, expressed in the organizer timezone.

The planner is greedy: requests are handled longest-first (they are hardest to
fit), and each one gets the earliest slot that
  1. sits completely inside an availability block,
  2. falls on one of the attendee's preferred weekdays *in their timezone*,
  3. is inside the attendee's local working hours,
  4. does not overlap another meeting (plus a buffer).

Outputs: one ``.ics`` invite per meeting, a combined ``agenda.ics`` and an
Excel agenda with an "Unscheduled" sheet explaining what could not be placed.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import UTC, date, datetime, time, timedelta
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

from openpyxl import Workbook

from .excel_utils import read_table, write_table

DAY_NAMES = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]
DEFAULT_TZ = "America/Argentina/Buenos_Aires"


@dataclass(frozen=True)
class MeetingRequest:
    """One meeting someone asked for."""

    request_id: str
    attendee: str
    email: str
    topic: str
    duration: timedelta
    preferred_days: frozenset[int]  # 0 = Monday
    tz: ZoneInfo
    earliest: time
    latest: time


@dataclass(frozen=True)
class Block:
    """A free block in the organizer's calendar (aware datetimes)."""

    start: datetime
    end: datetime


@dataclass(frozen=True)
class Meeting:
    """A request that was given a slot."""

    request: MeetingRequest
    start: datetime  # aware, UTC
    end: datetime

    def local(self, tz: ZoneInfo) -> tuple[datetime, datetime]:
        return self.start.astimezone(tz), self.end.astimezone(tz)


@dataclass(frozen=True)
class Unscheduled:
    request: MeetingRequest
    reason: str


def parse_time(value: Any) -> time:
    """Accept ``time`` objects, ``datetime`` objects or ``"HH:MM"`` strings."""
    if isinstance(value, datetime):
        return value.time()
    if isinstance(value, time):
        return value
    hours, minutes = str(value).strip().split(":")[:2]
    return time(int(hours), int(minutes))


def parse_date(value: Any) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value).strip())


def parse_days(value: Any) -> frozenset[int]:
    """Turn ``"Mon, Wed"`` (or ``"Monday;Friday"``) into weekday numbers. Empty = any weekday."""
    text = str(value or "").replace(";", ",").replace("/", ",")
    days = set()
    for part in text.split(","):
        key = part.strip().lower()[:3]
        if not key:
            continue
        if key not in DAY_NAMES:
            raise ValueError(f"Unknown weekday '{part.strip()}'")
        days.add(DAY_NAMES.index(key))
    return frozenset(days or range(5))


def request_from_row(row: dict[str, Any]) -> MeetingRequest:
    return MeetingRequest(
        request_id=str(row["request_id"]).strip(),
        attendee=str(row["attendee"]).strip(),
        email=str(row.get("email") or "").strip(),
        topic=str(row.get("topic") or "Meeting").strip(),
        duration=timedelta(minutes=int(row["duration_min"])),
        preferred_days=parse_days(row.get("preferred_days")),
        tz=ZoneInfo(str(row["timezone"]).strip()),
        earliest=parse_time(row.get("earliest") or "09:00"),
        latest=parse_time(row.get("latest") or "17:00"),
    )


def block_from_row(row: dict[str, Any], tz: ZoneInfo) -> Block:
    day = parse_date(row["date"])
    start = datetime.combine(day, parse_time(row["start"]), tzinfo=tz)
    end = datetime.combine(day, parse_time(row["end"]), tzinfo=tz)
    if end <= start:
        raise ValueError(f"Availability block on {day} ends before it starts")
    return Block(start, end)


def _fits_attendee(req: MeetingRequest, start: datetime, end: datetime) -> bool:
    local_start, local_end = start.astimezone(req.tz), end.astimezone(req.tz)
    return (
        local_start.weekday() in req.preferred_days
        and local_start.date() == local_end.date()
        and local_start.time() >= req.earliest
        and local_end.time() <= req.latest
    )


def _overlaps(start: datetime, end: datetime, booked: list[Meeting], buffer: timedelta) -> bool:
    return any(start < m.end + buffer and m.start - buffer < end for m in booked)


def plan(
    requests: list[MeetingRequest],
    blocks: list[Block],
    step_minutes: int = 15,
    buffer_minutes: int = 15,
) -> tuple[list[Meeting], list[Unscheduled]]:
    """Assign each request the earliest valid slot. Returns (scheduled, unscheduled)."""
    step, buffer = timedelta(minutes=step_minutes), timedelta(minutes=buffer_minutes)
    ordered_blocks = sorted(blocks, key=lambda b: b.start)
    booked: list[Meeting] = []
    unscheduled: list[Unscheduled] = []

    # Longest meetings first; ties keep the original (priority) order.
    for req in sorted(requests, key=lambda r: -r.duration):
        slot = None
        fits_window = False
        for block in ordered_blocks:
            start = block.start.astimezone(UTC)
            block_end = block.end.astimezone(UTC)
            while start + req.duration <= block_end:
                end = start + req.duration
                if _fits_attendee(req, start, end):
                    fits_window = True
                    if not _overlaps(start, end, booked, buffer):
                        slot = Meeting(req, start, end)
                        break
                start += step
            if slot:
                break
        if slot:
            booked.append(slot)
        else:
            reason = (
                "all matching slots already booked"
                if fits_window
                else "no availability overlaps the attendee's preferred days/hours"
            )
            unscheduled.append(Unscheduled(req, reason))
    booked.sort(key=lambda m: m.start)
    return booked, unscheduled


# --------------------------------------------------------------------------- ICS


def _ics_escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")


def _fold(line: str) -> str:
    """Fold content lines longer than 75 octets (RFC 5545 §3.1)."""
    out, current = [], line
    while len(current.encode()) > 75:
        cut = 75
        while len(current[:cut].encode()) > 75:
            cut -= 1
        out.append(current[:cut])
        current = " " + current[cut:]
    out.append(current)
    return "\r\n".join(out)


def _fmt(dt: datetime) -> str:
    return dt.astimezone(UTC).strftime("%Y%m%dT%H%M%SZ")


def vevent(meeting: Meeting, organizer: str, stamp: datetime) -> list[str]:
    req = meeting.request
    local_start, _ = meeting.local(req.tz)
    desc = f"{req.topic} with {req.attendee}.\nYour local time: {local_start:%a %d %b %H:%M} ({req.tz.key})."
    lines = [
        "BEGIN:VEVENT",
        f"UID:{req.request_id}@va-toolkit",
        f"DTSTAMP:{_fmt(stamp)}",
        f"DTSTART:{_fmt(meeting.start)}",
        f"DTEND:{_fmt(meeting.end)}",
        f"SUMMARY:{_ics_escape(req.topic)}",
        f"DESCRIPTION:{_ics_escape(desc)}",
        f"ORGANIZER:mailto:{organizer}",
    ]
    if req.email:
        lines.append(f"ATTENDEE;CN={_ics_escape(req.attendee)};RSVP=TRUE:mailto:{req.email}")
    lines.append("END:VEVENT")
    return lines


def to_ics(meetings: list[Meeting], organizer: str, stamp: datetime | None = None) -> str:
    """Serialize meetings as an iCalendar (RFC 5545) invitation."""
    stamp = stamp or datetime.now(UTC)
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//va-toolkit//scheduler//EN", "METHOD:REQUEST"]
    for meeting in meetings:
        lines += vevent(meeting, organizer, stamp)
    lines.append("END:VCALENDAR")
    return "\r\n".join(_fold(line) for line in lines) + "\r\n"


# --------------------------------------------------------------------------- outputs


def write_agenda(meetings: list[Meeting], unscheduled: list[Unscheduled], tz: ZoneInfo, path: str | Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Agenda"
    rows = []
    for m in meetings:
        start, end = m.local(tz)
        a_start, a_end = m.local(m.request.tz)
        rows.append([
            start.strftime("%Y-%m-%d"), start.strftime("%a"), start.strftime("%H:%M"), end.strftime("%H:%M"),
            m.request.attendee, m.request.topic, int(m.request.duration.total_seconds() // 60),
            f"{a_start:%a %H:%M}-{a_end:%H:%M}", m.request.tz.key, m.request.request_id,
        ])
    write_table(
        ws,
        ["Date", "Day", f"Start ({tz.key})", "End", "Attendee", "Topic", "Minutes",
         "Attendee local time", "Attendee timezone", "Request ID"],
        rows,
    )
    ws2 = wb.create_sheet("Unscheduled")
    write_table(
        ws2,
        ["Request ID", "Attendee", "Topic", "Minutes", "Preferred days", "Timezone", "Reason"],
        ([u.request.request_id, u.request.attendee, u.request.topic,
          int(u.request.duration.total_seconds() // 60),
          ", ".join(DAY_NAMES[d].title() for d in sorted(u.request.preferred_days)),
          u.request.tz.key, u.reason] for u in unscheduled),
    )
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def run(
    requests_path: str | Path,
    availability_path: str | Path,
    out_dir: str | Path,
    organizer: str,
    tz_name: str = DEFAULT_TZ,
) -> tuple[list[Meeting], list[Unscheduled]]:
    """End-to-end: read both workbooks, plan, write invites and the agenda."""
    tz = ZoneInfo(tz_name)
    requests = [request_from_row(r) for r in read_table(requests_path)]
    blocks = [block_from_row(r, tz) for r in read_table(availability_path)]
    meetings, unscheduled = plan(requests, blocks)

    out = Path(out_dir)
    invites = out / "invites"
    invites.mkdir(parents=True, exist_ok=True)
    for old in invites.glob("*.ics"):
        old.unlink()
    stamp = datetime.now(UTC).replace(microsecond=0)
    for m in meetings:
        invite = to_ics([m], organizer, stamp)
        (invites / f"{m.request.request_id}.ics").write_text(invite, encoding="utf-8", newline="")
    (out / "agenda.ics").write_text(to_ics(meetings, organizer, stamp), encoding="utf-8", newline="")
    write_agenda(meetings, unscheduled, tz, out / "agenda.xlsx")
    return meetings, unscheduled
