import datetime as dt
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
import timesheet as ts  # noqa: E402

CFG = ts.load_config()
DAYS = ts.week_days(CFG, "2026-11-09")
TZ = ts.tz_of(CFG)

# Thu 12 / Fri 13 Nov 2026 as read from the calendar.
EVENTS = [
    {"date": "2026-11-12", "start": "09:15", "end": "09:30", "title": "Netzon Infrastructure Sync"},
    {"date": "2026-11-12", "start": "10:30", "end": "10:45", "title": "AWS Partner - Weekly"},
    {"date": "2026-11-12", "start": "12:00", "end": "14:00", "title": "Focus time"},
    {"date": "2026-11-12", "start": "12:00", "end": "12:30", "title": "Infrastructure"},
    {"date": "2026-11-12", "start": "14:45", "end": "15:00", "title": "Take a break"},
    {"date": "2026-11-12", "start": "15:00", "end": "16:00", "title": "Data lake structure"},
    {"date": "2026-11-12", "start": "15:30", "end": "16:00", "title": "Catch up"},
    {"date": "2026-11-12", "start": "16:00", "end": "16:45", "title": "Netzon Data Foundation"},
    {"date": "2026-11-13", "start": "09:15", "end": "09:30", "title": "Netzon Infrastructure Sync"},
    {"date": "2026-11-13", "start": "09:30", "end": "11:00", "title": "Learning time"},
    {"date": "2026-11-13", "start": "09:30", "end": "10:00", "title": "Netzon"},
    {"date": "2026-11-13", "start": "12:30", "end": "13:00", "title": "Uninstall ESET AV"},
    {"date": "2026-11-13", "start": "14:30", "end": "16:30", "title": "Focus time"},
    {"date": "2026-11-13", "start": "14:45", "end": "15:00", "title": "Take a break"},
    {"date": "2026-11-13", "start": "16:00", "end": "16:30", "title": "Netzon"},
    {"date": "2026-11-13", "start": "16:30", "end": "17:00", "title": "Catch up on messages"},
]


def full_day(date):
    d = dt.date.fromisoformat(date)
    return [{"start": ts.at(d, "09:00", TZ), "end": ts.at(d, "17:00", TZ), "hours": 8.0}]


STATUS = {"2026-11-09": full_day("2026-11-09"), "2026-11-10": full_day("2026-11-10"),
          "2026-11-11": full_day("2026-11-11"), "2026-11-12": [], "2026-11-13": []}


def fake_resolve(ref, day):
    if isinstance(ref, dict):
        name = ref["name"].format(month=day.strftime("%B"), year=day.year)
        return ("z8znm1kc3d" if ref["list"] == "trainings" else f"{ref['list']}/{name}", name)
    return (ref, ref)


def rows(plan, day):
    return [(e["start"], e["end"], e["task_id"], e["description"]) for e in plan["days"][day]["entries"]]


class PlanTest(unittest.TestCase):
    def setUp(self):
        focus = ts.parse_focus(["thu=PMO integration", "fri=Data Foundation planning"], DAYS)
        self.plan = ts.build_plan(CFG, DAYS, STATUS, EVENTS, focus, fake_resolve)

    def test_full_days_skipped(self):
        self.assertEqual([s["date"] for s in self.plan["skipped"]], ["2026-11-09", "2026-11-10", "2026-11-11"])

    def test_thursday_matches_hand_logged(self):
        self.assertEqual(rows(self.plan, "2026-11-12"), [
            ("09:00", "09:15", "z8znm1nk9u", "PMO integration"),
            ("09:15", "09:30", "z8znm1pg7e", "Netzon Infra meeting"),
            ("09:30", "10:30", "z8znm1nk9u", "PMO integration"),
            ("10:30", "10:45", "z8znm1qrde", "AWS Partner meeting"),
            ("10:45", "15:00", "z8znm1nk9u", "PMO integration"),
            ("15:00", "16:00", "z8znm1mxep", "Data lake structure"),
            ("16:00", "16:45", "z8znm1mxep", "Netzon Data Foundation"),
            ("16:45", "17:00", "z8znm1nk9u", "PMO integration"),
        ])
        self.assertEqual(self.plan["days"]["2026-11-12"]["total"], 8.0)

    def test_friday_matches_hand_logged(self):
        self.assertEqual(rows(self.plan, "2026-11-13"), [
            ("09:00", "09:15", "data_foundation/Project Planning", "Data Foundation planning"),
            ("09:15", "09:30", "z8znm1pg7e", "Netzon Infra meeting"),
            ("09:30", "11:00", "z8znm1kc3d", "Learning time"),
            ("11:00", "12:30", "data_foundation/Project Planning", "Data Foundation planning"),
            ("12:30", "13:00", "z8znm1qrdf", "Uninstall ESET AV"),
            ("13:00", "16:30", "data_foundation/Project Planning", "Data Foundation planning"),
            ("16:30", "17:00", "z8znm1qrdf", "Catch up on messages"),
        ])
        self.assertTrue(all(e["flag"] is None for e in self.plan["days"]["2026-11-13"]["entries"]))

    def test_missing_focus_note_is_flagged(self):
        plan = ts.build_plan(CFG, DAYS, STATUS, EVENTS, {}, fake_resolve)
        flags = {e["flag"] for e in plan["days"]["2026-11-12"]["entries"]}
        self.assertIn("focus block: what did you work on?", flags)
        self.assertIn("open time: which task?", flags)

    def test_partial_day_only_fills_remainder(self):
        st = dict(STATUS)
        d = dt.date(2026, 11, 12)
        st["2026-11-12"] = [{"start": ts.at(d, "09:00", TZ), "end": ts.at(d, "12:00", TZ), "hours": 3.0}]
        focus = ts.parse_focus(["thu=PMO integration"], DAYS)
        plan = ts.build_plan(CFG, DAYS, st, EVENTS, focus, fake_resolve)
        self.assertEqual(plan["days"]["2026-11-12"]["total"], 8.0)
        self.assertEqual(plan["days"]["2026-11-12"]["entries"][0]["start"], "12:00")


ICS = """BEGIN:VCALENDAR
BEGIN:VEVENT
UID:infra
SUMMARY:Netzon Infra
DTSTART;TZID=Singapore Standard Time:20261005T091500
DTEND;TZID=Singapore Standard Time:20261005T093000
RRULE:FREQ=WEEKLY;INTERVAL=1;BYDAY=MO,TU,WE,TH,FR
EXDATE;TZID=Singapore Standard Time:20261110T091500
END:VEVENT
BEGIN:VEVENT
UID:infra
RECURRENCE-ID;TZID=Singapore Standard Time:20261112T091500
SUMMARY:Netzon Infra (moved)
DTSTART;TZID=Singapore Standard Time:20261112T100000
DTEND;TZID=Singapore Standard Time:20261112T101500
END:VEVENT
BEGIN:VEVENT
UID:df
SUMMARY:Netzon Data Foundation\\, weekly
DTSTART:20261112T080000Z
DTEND:20261112T084500Z
END:VEVENT
BEGIN:VEVENT
UID:allday
SUMMARY:Holiday
DTSTART;VALUE=DATE:20261113
DTEND;VALUE=DATE:20261114
END:VEVENT
END:VCALENDAR
"""


class IcsTest(unittest.TestCase):
    def test_expand(self):
        evs = ts.ics_events(CFG, DAYS, ICS)
        got = [(e["date"], e["start"], e["end"], e["title"]) for e in evs]
        self.assertEqual(got, [
            ("2026-11-09", "09:15", "09:30", "Netzon Infra"),
            ("2026-11-11", "09:15", "09:30", "Netzon Infra"),
            ("2026-11-12", "10:00", "10:15", "Netzon Infra (moved)"),
            ("2026-11-12", "16:00", "16:45", "Netzon Data Foundation, weekly"),
            ("2026-11-13", "09:00", "17:00", "Holiday"),  # all-day, flagged all_day=True
            ("2026-11-13", "09:15", "09:30", "Netzon Infra"),
        ])


    def test_real_world_rules(self):
        """Patterns seen in a published Outlook feed: Paris-time event, weekly series with EXDATE
        and a moved occurrence, nth-weekday monthly, yearly, multi-day and recurring all-day."""
        ics = """BEGIN:VCALENDAR
BEGIN:VEVENT
UID:standup
SUMMARY:Team Standup
DTSTART;TZID=Singapore Standard Time:20261001T101500
DTEND;TZID=Singapore Standard Time:20261001T103000
RRULE:FREQ=WEEKLY;UNTIL=20261218T021500Z;INTERVAL=1;BYDAY=MO,TU,TH,FR;WKST=SU
EXDATE;TZID=Singapore Standard Time:20261103T101500,20261110T101500
X-MICROSOFT-CDO-BUSYSTATUS:BUSY
END:VEVENT
BEGIN:VEVENT
UID:standup
RECURRENCE-ID;TZID=Singapore Standard Time:20261113T101500
SUMMARY:Team Standup
DTSTART;TZID=Singapore Standard Time:20261113T110000
DTEND;TZID=Singapore Standard Time:20261113T111500
END:VEVENT
BEGIN:VEVENT
UID:paris
SUMMARY:Project sync
DTSTART;TZID=Romance Standard Time:20261112T090000
DTEND;TZID=Romance Standard Time:20261112T100000
END:VEVENT
BEGIN:VEVENT
UID:monthly
SUMMARY:Monthly review
DTSTART;TZID=Singapore Standard Time:20260908T150000
DTEND;TZID=Singapore Standard Time:20260908T160000
RRULE:FREQ=MONTHLY;UNTIL=20271231T000000Z;INTERVAL=1;BYDAY=2TU
END:VEVENT
BEGIN:VEVENT
UID:yearly
SUMMARY:Renewal
DTSTART;VALUE=DATE:20251110
DTEND;VALUE=DATE:20251111
RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=2MO;BYMONTH=11
X-MICROSOFT-CDO-BUSYSTATUS:FREE
END:VEVENT
BEGIN:VEVENT
UID:nomeet
SUMMARY:No-meeting day
DTSTART;VALUE=DATE:20261014
DTEND;VALUE=DATE:20261015
RRULE:FREQ=WEEKLY;UNTIL=20261231T000000Z;INTERVAL=2;BYDAY=WE;WKST=SU
X-MICROSOFT-CDO-BUSYSTATUS:FREE
END:VEVENT
BEGIN:VEVENT
UID:sick
SUMMARY:OOO - Sick Leave
DTSTART;VALUE=DATE:20261112
DTEND;VALUE=DATE:20261114
X-MICROSOFT-CDO-BUSYSTATUS:OOF
END:VEVENT
END:VCALENDAR
"""
        evs = ts.ics_events(CFG, DAYS, ics)
        got = [(e["date"], e["start"], e["end"], e["title"], e.get("all_day", False)) for e in evs]
        self.assertEqual(got, [
            ("2026-11-09", "09:00", "17:00", "Renewal", True),
            ("2026-11-09", "10:15", "10:30", "Team Standup", False),
            ("2026-11-10", "15:00", "16:00", "Monthly review", False),
            ("2026-11-11", "09:00", "17:00", "No-meeting day", True),
            ("2026-11-12", "09:00", "17:00", "OOO - Sick Leave", True),
            ("2026-11-12", "10:15", "10:30", "Team Standup", False),
            ("2026-11-12", "16:00", "17:00", "Project sync", False),
            ("2026-11-13", "09:00", "17:00", "OOO - Sick Leave", True),
            ("2026-11-13", "11:00", "11:15", "Team Standup", False),
        ])
        # all-day OOO days are not planned; other all-day events don't become work blocks
        st = {k: [] for k in STATUS}
        plan = ts.build_plan(CFG, DAYS, st, evs, {}, fake_resolve)
        self.assertEqual({s["date"] for s in plan["skipped"]}, {"2026-11-12", "2026-11-13"})
        self.assertEqual([e["source"] for e in plan["days"]["2026-11-11"]["entries"]], ["(open time)"])


if __name__ == "__main__":
    unittest.main()
