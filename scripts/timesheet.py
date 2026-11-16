#!/usr/bin/env python3
"""ClickUp timesheet helper.

Checks a week's ClickUp time entries against a daily target, reads calendar
events (Microsoft Graph, an .ics feed, or a hand-made JSON file), drafts a plan
that maps events and open time to ClickUp tasks, and logs the plan once it has
been confirmed.

Standard library only. Run `timesheet.py -h` for the subcommands.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import subprocess
import sys
import time
import tomllib
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

ROOT = Path(__file__).resolve().parent.parent
CLICKUP_API = "https://api.clickup.com/api/v2"
GRAPH_API = "https://graph.microsoft.com/v1.0"
# Microsoft's own public client for "Microsoft Graph Command Line Tools".
# It lets a user consent to delegated scopes like Calendars.Read without an app registration.
DEFAULT_MS_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
MS_SCOPE = "offline_access Calendars.Read"
TOKEN_FILE = Path(os.environ.get("TIMESHEET_TOKEN_FILE", Path.home() / ".config/clickup-timesheet/ms_token.json"))
DAY_NAMES = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]
EPS = 1e-6


# --------------------------------------------------------------------------- config

def load_config(path: str | None = None) -> dict:
    cfg = tomllib.loads(Path(path or ROOT / "config.toml").read_text())
    maps = ROOT / "mappings.toml"
    cfg["mappings"] = tomllib.loads(maps.read_text()) if maps.exists() else {}
    return cfg


def tz_of(cfg) -> ZoneInfo:
    return ZoneInfo(cfg["workday"]["timezone"])


def at(day: dt.date, hhmm: str, tz) -> dt.datetime:
    h, m = map(int, hhmm.split(":"))
    return dt.datetime.combine(day, dt.time(h, m), tz)


def ms(d: dt.datetime) -> int:
    return int(d.timestamp() * 1000)


def hours(a: dt.datetime, b: dt.datetime) -> float:
    return (b - a).total_seconds() / 3600


def week_days(cfg, spec: str | None) -> list[dt.date]:
    today = dt.datetime.now(tz_of(cfg)).date()
    if spec in (None, "", "last"):
        ref = today - dt.timedelta(days=7)
    elif spec == "this":
        ref = today
    else:
        ref = dt.date.fromisoformat(spec)
    monday = ref - dt.timedelta(days=ref.weekday())
    return [monday + dt.timedelta(days=d) for d in cfg["workday"].get("days", [0, 1, 2, 3, 4])]


# --------------------------------------------------------------------------- http

def http_json(url, method="GET", headers=None, body=None, form=None, timeout=60):
    data = None
    headers = dict(headers or {})
    if body is not None:
        data = json.dumps(body).encode()
        headers["Content-Type"] = "application/json"
    elif form is not None:
        data = urllib.parse.urlencode(form).encode()
        headers["Content-Type"] = "application/x-www-form-urlencoded"
    req = urllib.request.Request(url, data=data, method=method, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.load(r)


# --------------------------------------------------------------------------- clickup

class ClickUp:
    def __init__(self, cfg):
        self.key = os.environ.get("CLICKUP_API_KEY")
        if not self.key:
            sys.exit("CLICKUP_API_KEY is not set (add it as an environment secret).")
        self.team = str(cfg["clickup"]["team_id"])
        self.user = int(cfg["clickup"]["user_id"])
        self.billable = bool(cfg["clickup"].get("billable", False))
        self._list_cache: dict[str, list] = {}

    def req(self, method, path, params=None, body=None):
        url = CLICKUP_API + path + ("?" + urllib.parse.urlencode(params, doseq=True) if params else "")
        try:
            return http_json(url, method, {"Authorization": self.key}, body=body)
        except urllib.error.HTTPError as e:
            sys.exit(f"ClickUp {method} {path} failed: HTTP {e.code} {e.read().decode()[:300]}")

    def entries(self, start: dt.datetime, end: dt.datetime) -> list:
        r = self.req("GET", f"/team/{self.team}/time_entries",
                     {"start_date": ms(start), "end_date": ms(end), "assignee": self.user})
        return r.get("data", [])

    def list_tasks(self, list_id: str) -> list:
        if list_id not in self._list_cache:
            out, page = [], 0
            while True:
                r = self.req("GET", f"/list/{list_id}/task",
                             {"include_closed": "true", "subtasks": "true", "page": page})
                out += r.get("tasks", [])
                if r.get("last_page", True) or not r.get("tasks"):
                    break
                page += 1
            self._list_cache[list_id] = out
        return self._list_cache[list_id]

    def task(self, tid: str) -> dict:
        return self.req("GET", f"/task/{tid}")

    def create_task(self, list_id: str, name: str) -> dict:
        self._list_cache.pop(list_id, None)
        return self.req("POST", f"/list/{list_id}/task", body={"name": name, "assignees": [self.user]})

    def create_entry(self, tid: str, start: dt.datetime, hrs: float, desc: str) -> dict:
        return self.req("POST", f"/team/{self.team}/time_entries", body={
            "tid": tid, "start": ms(start), "duration": int(round(hrs * 3600000)),
            "description": desc, "billable": self.billable, "assignee": self.user})


def list_id(cfg, alias_or_id: str) -> str:
    return str(cfg.get("lists", {}).get(alias_or_id, alias_or_id))


def fetch_status(cu: ClickUp, cfg, days: list[dt.date]) -> dict:
    tz = tz_of(cfg)
    start = dt.datetime.combine(days[0], dt.time(), tz)
    end = dt.datetime.combine(days[-1] + dt.timedelta(days=1), dt.time(), tz)
    by = {d.isoformat(): [] for d in days}
    for e in cu.entries(start, end):
        dur = int(e.get("duration") or 0)
        if dur <= 0:  # running timer
            continue
        s = dt.datetime.fromtimestamp(int(e["start"]) / 1000, tz)
        task = e.get("task") or {}
        if s.date().isoformat() in by:
            by[s.date().isoformat()].append({
                "start": s, "end": s + dt.timedelta(milliseconds=dur), "hours": dur / 3.6e6,
                "task_id": task.get("id"), "task": task.get("name"),
                "description": e.get("description") or ""})
    for v in by.values():
        v.sort(key=lambda x: x["start"])
    return by


def print_status(cfg, status: dict, verbose=False):
    target = cfg["workday"]["target_hours"]
    for day, ents in status.items():
        tot = sum(e["hours"] for e in ents)
        mark = "OK " if tot >= target - EPS else "LOW"
        name = dt.date.fromisoformat(day).strftime("%a")
        print(f"{mark} {name} {day}  {tot:5.2f}h / {target}h")
        if verbose:
            for e in ents:
                print(f"      {e['start']:%H:%M}  {e['hours']:4.2f}h  {e['task']}  | {e['description']}")


# --------------------------------------------------------------------------- calendar: graph

def ms_device_login(client_id: str, tenant: str):
    base = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0"
    dc = http_json(f"{base}/devicecode", "POST", form={"client_id": client_id, "scope": MS_SCOPE})
    print(dc["message"], flush=True)
    deadline = time.time() + int(dc.get("expires_in", 900))
    interval = int(dc.get("interval", 5))
    while time.time() < deadline:
        time.sleep(interval)
        try:
            tok = http_json(f"{base}/token", "POST", form={
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                "client_id": client_id, "device_code": dc["device_code"]})
        except urllib.error.HTTPError as e:
            err = json.loads(e.read().decode() or "{}").get("error")
            if err == "authorization_pending":
                continue
            if err == "slow_down":
                interval += 5
                continue
            sys.exit(f"Sign-in failed: {err}")
        save_token(tok, client_id, tenant)
        print(f"Signed in. Refresh token saved to {TOKEN_FILE}")
        return
    sys.exit("Sign-in timed out.")


def save_token(tok: dict, client_id: str, tenant: str):
    TOKEN_FILE.parent.mkdir(parents=True, exist_ok=True)
    TOKEN_FILE.write_text(json.dumps({
        "client_id": client_id, "tenant": tenant, "refresh_token": tok.get("refresh_token"),
        "access_token": tok.get("access_token"),
        "expires_at": time.time() + int(tok.get("expires_in", 0)) - 60}))
    TOKEN_FILE.chmod(0o600)


def graph_token() -> str:
    if os.environ.get("MS_GRAPH_TOKEN"):
        return os.environ["MS_GRAPH_TOKEN"]
    stored = json.loads(TOKEN_FILE.read_text()) if TOKEN_FILE.exists() else {}
    if stored.get("access_token") and stored.get("expires_at", 0) > time.time():
        return stored["access_token"]
    refresh = stored.get("refresh_token") or os.environ.get("MS_REFRESH_TOKEN")
    if refresh:
        client = stored.get("client_id") or os.environ.get("MS_CLIENT_ID", DEFAULT_MS_CLIENT_ID)
        tenant = stored.get("tenant") or os.environ.get("MS_TENANT_ID", "organizations")
        try:
            tok = http_json(f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token", "POST", form={
                "grant_type": "refresh_token", "client_id": client, "refresh_token": refresh, "scope": MS_SCOPE})
        except urllib.error.HTTPError as e:
            sys.exit(f"Microsoft token refresh failed: HTTP {e.code}. Run `timesheet.py ms-login` again.")
        save_token(tok, client, tenant)
        return tok["access_token"]
    try:
        return subprocess.run(["az", "account", "get-access-token", "--resource-type", "ms-graph",
                               "--query", "accessToken", "-o", "tsv"],
                              check=True, capture_output=True, text=True).stdout.strip()
    except (FileNotFoundError, subprocess.CalledProcessError):
        sys.exit("No Microsoft credentials. Run `timesheet.py ms-login`, or set MS_REFRESH_TOKEN.")


def graph_events(cfg, days: list[dt.date]) -> list[dict]:
    tz = tz_of(cfg)
    start = dt.datetime.combine(days[0], dt.time(), tz).astimezone(dt.timezone.utc)
    end = dt.datetime.combine(days[-1] + dt.timedelta(days=1), dt.time(), tz).astimezone(dt.timezone.utc)
    url = f"{GRAPH_API}/me/calendarView?" + urllib.parse.urlencode({
        "startDateTime": start.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "endDateTime": end.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "$select": "subject,start,end,showAs,isAllDay,isCancelled,responseStatus",
        "$top": "200", "$orderby": "start/dateTime"})
    headers = {"Authorization": f"Bearer {graph_token()}", "Prefer": 'outlook.timezone="UTC"'}
    out = []
    while url:
        try:
            r = http_json(url, headers=headers)
        except urllib.error.HTTPError as e:
            sys.exit(f"Graph calendarView failed: HTTP {e.code} {e.read().decode()[:300]}")
        for ev in r.get("value", []):
            if ev.get("isCancelled") or ev.get("isAllDay"):
                continue
            if (ev.get("responseStatus") or {}).get("response") == "declined":
                continue
            s = parse_graph_dt(ev["start"]["dateTime"]).astimezone(tz)
            e = parse_graph_dt(ev["end"]["dateTime"]).astimezone(tz)
            out.append({"date": s.date().isoformat(), "start": f"{s:%H:%M}", "end": f"{e:%H:%M}",
                        "title": ev.get("subject") or "(no title)", "show_as": ev.get("showAs")})
        url = r.get("@odata.nextLink")
    return out


def parse_graph_dt(s: str) -> dt.datetime:
    s = s.rstrip("Z")
    if "." in s:
        s = s[: s.index(".") + 7]
    return dt.datetime.fromisoformat(s).replace(tzinfo=dt.timezone.utc)


# --------------------------------------------------------------------------- calendar: ics

WINDOWS_TZ = {
    "Singapore Standard Time": "Asia/Singapore", "China Standard Time": "Asia/Shanghai",
    "Taipei Standard Time": "Asia/Taipei", "Tokyo Standard Time": "Asia/Tokyo",
    "India Standard Time": "Asia/Kolkata", "AUS Eastern Standard Time": "Australia/Sydney",
    "GMT Standard Time": "Europe/London", "W. Europe Standard Time": "Europe/Berlin",
    "Romance Standard Time": "Europe/Paris", "Central Europe Standard Time": "Europe/Budapest",
    "Central European Standard Time": "Europe/Warsaw", "E. Europe Standard Time": "Europe/Bucharest",
    "Eastern Standard Time": "America/New_York", "Central Standard Time": "America/Chicago",
    "Mountain Standard Time": "America/Denver", "Pacific Standard Time": "America/Los_Angeles",
    "UTC": "UTC", "Coordinated Universal Time": "UTC",
}


def ics_tz(tzid: str | None, default):
    if not tzid:
        return default
    tzid = tzid.strip('"')
    for name in (tzid, WINDOWS_TZ.get(tzid)):
        if name:
            try:
                return ZoneInfo(name)
            except (ZoneInfoNotFoundError, ValueError):
                pass
    return default


def ics_parse_dt(value: str, params: dict, default_tz):
    if params.get("VALUE") == "DATE" or re.fullmatch(r"\d{8}", value):
        return dt.datetime.strptime(value, "%Y%m%d").date()
    if value.endswith("Z"):
        return dt.datetime.strptime(value, "%Y%m%dT%H%M%SZ").replace(tzinfo=dt.timezone.utc)
    return dt.datetime.strptime(value, "%Y%m%dT%H%M%S").replace(tzinfo=ics_tz(params.get("TZID"), default_tz))


def ics_duration(v: str) -> dt.timedelta:
    m = re.fullmatch(r"-?P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?", v)
    if not m:
        return dt.timedelta()
    w, d, h, mi, s = (int(x or 0) for x in m.groups())
    return dt.timedelta(weeks=w, days=d, hours=h, minutes=mi, seconds=s)


def ics_read_events(text: str, default_tz) -> list[dict]:
    lines = []
    for raw in text.splitlines():
        if raw[:1] in (" ", "\t") and lines:
            lines[-1] += raw[1:]
        else:
            lines.append(raw)
    events, cur = [], None
    for line in lines:
        if line == "BEGIN:VEVENT":
            cur = {"exdate": []}
        elif line == "END:VEVENT":
            if cur is not None:
                events.append(cur)
            cur = None
        elif cur is not None and ":" in line:
            head, value = line.split(":", 1)
            name, *ps = head.split(";")
            params = dict(p.split("=", 1) for p in ps if "=" in p)
            name = name.upper()
            if name in ("DTSTART", "DTEND", "RECURRENCE-ID"):
                cur[name] = ics_parse_dt(value, params, default_tz)
            elif name == "EXDATE":
                cur["exdate"] += [ics_parse_dt(v, params, default_tz) for v in value.split(",")]
            elif name == "SUMMARY":
                cur["SUMMARY"] = value.replace("\\,", ",").replace("\\;", ";").replace("\\n", " ")
            else:
                cur[name] = value
    return events


def nth_weekday(year: int, month: int, n: int, wd: int):
    """n-th (1-based, negative from the end) weekday wd of a month, or None."""
    first = dt.date(year, month, 1)
    days = [first + dt.timedelta(days=i) for i in range(31)
            if (first + dt.timedelta(days=i)).month == month and (first + dt.timedelta(days=i)).weekday() == wd]
    try:
        return days[n - 1] if n > 0 else days[n]
    except IndexError:
        return None


def rrule_starts(start: dt.datetime, rule: str, until_window: dt.datetime):
    """Yield occurrence starts for the RRULE subset Outlook emits:
    DAILY, WEEKLY (BYDAY), MONTHLY and YEARLY (by month day or n-th weekday, e.g. BYDAY=2TU / -1FR)."""
    r = dict(p.split("=", 1) for p in rule.split(";") if "=" in p)
    freq, interval = r.get("FREQ"), int(r.get("INTERVAL", 1))
    count = int(r["COUNT"]) if "COUNT" in r else None
    until = None
    if "UNTIL" in r:
        u = ics_parse_dt(r["UNTIL"], {}, start.tzinfo)
        until = dt.datetime.combine(u, dt.time(23, 59), start.tzinfo) if isinstance(u, dt.date) and not isinstance(u, dt.datetime) else u
    ical_days = {"MO": 0, "TU": 1, "WE": 2, "TH": 3, "FR": 4, "SA": 5, "SU": 6}
    byday_raw = [d.upper() for d in r.get("BYDAY", "").split(",") if d[-2:].upper() in ical_days]
    byday = [ical_days[d[-2:]] for d in byday_raw]
    nth = [(int(d[:-2] or r.get("BYSETPOS", 0) or 0), ical_days[d[-2:]]) for d in byday_raw]
    bymonthday = [int(x) for x in r.get("BYMONTHDAY", "").split(",") if x]
    bymonth = [int(x) for x in r.get("BYMONTH", "").split(",") if x]
    n = 0

    def emit(s):
        nonlocal n
        if until and s > until:
            return False
        if count is not None and n >= count:
            return False
        n += 1
        return True

    def month_days(y, m):
        if nth and any(k for k, _ in nth):
            ds = [nth_weekday(y, m, k, w) for k, w in nth]
        else:
            ds = []
            for md in (bymonthday or [start.day]):
                try:
                    ds.append(dt.date(y, m, md))
                except ValueError:
                    pass
        return sorted(d for d in ds if d)

    for i in range(5000):
        if freq == "DAILY":
            cands = [start + dt.timedelta(days=i * interval)]
        elif freq == "WEEKLY":
            week0 = start - dt.timedelta(days=start.weekday())
            wk = week0 + dt.timedelta(weeks=i * interval)
            cands = sorted(wk + dt.timedelta(days=d) for d in (byday or [start.weekday()]))
        elif freq in ("MONTHLY", "YEARLY"):
            if freq == "MONTHLY":
                y, m = divmod(start.month - 1 + i * interval, 12)
                months = [(start.year + y, m + 1)]
            else:
                months = [(start.year + i * interval, m) for m in (bymonth or [start.month])]
            cands = [dt.datetime.combine(d, start.time(), start.tzinfo) for y, m in months for d in month_days(y, m)]
        else:
            cands = [start] if i == 0 else []
        for c in cands:
            if c < start:
                continue
            # keep wall-clock time across DST changes
            c = c.replace(tzinfo=start.tzinfo)
            if c > until_window or not emit(c):
                return
            yield c
        if freq not in ("DAILY", "WEEKLY", "MONTHLY", "YEARLY"):
            return


def ics_events(cfg, days: list[dt.date], text: str) -> list[dict]:
    """Events in the week, in local time. All-day events (holidays, OOO, sick leave) come out once
    per covered day with all_day=True and the workday as start/end; the planner decides what they mean."""
    tz = tz_of(cfg)
    wd = cfg["workday"]
    win_start = dt.datetime.combine(days[0], dt.time(), tz)
    win_end = dt.datetime.combine(days[-1] + dt.timedelta(days=1), dt.time(), tz)
    raw = ics_read_events(text, tz)

    def as_dt(x):  # all-day values are dates; expand them as local midnight
        return x if isinstance(x, dt.datetime) else dt.datetime.combine(x, dt.time(), tz)

    overrides = {(e.get("UID"), as_dt(e["RECURRENCE-ID"])) for e in raw if "RECURRENCE-ID" in e}
    out = []
    for e in raw:
        s0 = e.get("DTSTART")
        if not isinstance(s0, dt.date):
            continue  # malformed
        if e.get("STATUS", "").upper() == "CANCELLED":
            continue
        all_day = not isinstance(s0, dt.datetime)
        end0 = e.get("DTEND")
        if isinstance(end0, dt.date):
            length = as_dt(end0) - as_dt(s0)
        else:
            length = ics_duration(e.get("DURATION", "P1D" if all_day else "PT0M"))
        s0 = as_dt(s0)
        starts = [s0]
        if "RRULE" in e and "RECURRENCE-ID" not in e:
            ex = {as_dt(x) for x in e["exdate"]}
            starts = [s for s in rrule_starts(s0, e["RRULE"], win_end)
                      if s not in ex and (e.get("UID"), s) not in overrides]
        busy = e.get("X-MICROSOFT-CDO-BUSYSTATUS", "").lower() or None
        title = e.get("SUMMARY", "(no title)")
        for s in starts:
            en = s + length
            if en <= win_start or s >= win_end:
                continue
            if all_day:
                d = s.date()
                while d < en.date() or d == s.date():
                    if days[0] <= d <= days[-1]:
                        out.append({"date": d.isoformat(), "start": wd["start"], "end": wd["end"],
                                    "title": title, "show_as": busy, "all_day": True})
                    d += dt.timedelta(days=1)
                continue
            ls, le = s.astimezone(tz), en.astimezone(tz)
            out.append({"date": ls.date().isoformat(), "start": f"{ls:%H:%M}", "end": f"{le:%H:%M}",
                        "title": title, "show_as": busy})
    out.sort(key=lambda x: (x["date"], not x.get("all_day", False), x["start"]))
    return out


def load_ics(src: str) -> str:
    if re.match(r"https?://", src):
        req = urllib.request.Request(src, headers={"User-Agent": "clickup-timesheet"})
        with urllib.request.urlopen(req, timeout=60) as r:
            return r.read().decode("utf-8", "replace")
    return Path(src).read_text()


# --------------------------------------------------------------------------- planning

def subtract(iv, occupied):
    """Parts of interval iv=(a,b) not covered by any interval in occupied."""
    parts = [iv]
    for oa, ob in occupied:
        nxt = []
        for a, b in parts:
            if ob <= a or oa >= b:
                nxt.append((a, b))
                continue
            if a < oa:
                nxt.append((a, oa))
            if ob < b:
                nxt.append((ob, b))
        parts = nxt
    return [(a, b) for a, b in parts if b > a]


def parse_focus(items: list[str], days: list[dt.date]) -> dict[str, list[str]]:
    """'thu=PMO integration' or '2026-11-12=PMO integration' -> {date: [notes in order]}."""
    out: dict[str, list[str]] = {}
    for it in items or []:
        key, _, note = it.partition("=")
        key = key.strip().lower()
        day = None
        if key[:3] in DAY_NAMES:
            day = next((d for d in days if d.weekday() == DAY_NAMES.index(key[:3])), None)
        else:
            try:
                day = dt.date.fromisoformat(key)
            except ValueError:
                pass
        if not day:
            sys.exit(f"Can't read focus note {it!r}; use e.g. thu=PMO integration")
        out.setdefault(day.isoformat(), []).append(note.strip())
    return out


def any_match(patterns, text):
    return any(re.search(p, text) for p in patterns or [])


def focus_title_note(patterns, title):
    """'Focus time (Data Foundation Planning)' -> 'Data Foundation Planning'; plain 'Focus time' -> None."""
    for p in patterns or []:
        m = re.search(p, title)
        if m:
            rest = (title[:m.start()] + title[m.end():]).strip(" ()[]-:|")
            return rest or None
    return None


def build_plan(cfg, days, status, events, focus, resolve) -> dict:
    """resolve(task_ref, date) -> (task_id, task_name) or None."""
    tz = tz_of(cfg)
    wd, cal, maps = cfg["workday"], cfg.get("calendar", {}), cfg["mappings"]
    target = float(wd["target_hours"])
    plan = {"target": target, "days": {}, "warnings": [], "skipped": []}

    def map_rule(rules, text, day):
        for r in rules or []:
            if re.search(r["match"], text):
                res = resolve(r["task"], day)
                if res is None:
                    return None, None, f"rule task {r['task']!r} not found (create it?)", r.get("desc")
                return res[0], res[1], None, r.get("desc")
        return None, None, "unmapped", None

    for day in days:
        key = day.isoformat()
        existing = status.get(key, [])
        have = sum(e["hours"] for e in existing)
        if have >= target - EPS:
            plan["skipped"].append({"date": key, "hours": round(have, 2)})
            continue
        all_day = [ev for ev in events if ev["date"] == key and ev.get("all_day")]
        ooo = [ev for ev in all_day if (ev.get("show_as") or "").lower() == "oof"]
        if ooo:
            plan["warnings"].append(f"{key} all-day out of office ('{ooo[0]['title']}'); day not planned")
            plan["skipped"].append({"date": key, "hours": round(have, 2), "reason": "out of office"})
            continue
        ws, we = at(day, wd["start"], tz), at(day, wd["end"], tz)
        occupied = [(max(e["start"], ws), min(e["end"], we)) for e in existing if e["end"] > ws and e["start"] < we]
        segs = []
        notes = list(focus.get(key, []))
        day_events = []
        for ev in events:
            if ev["date"] != key or ev.get("all_day") or any_match(cal.get("ignore"), ev["title"]):
                continue
            s, e = max(at(day, ev["start"], tz), ws), min(at(day, ev["end"], tz), we)
            if e > s:
                day_events.append((s, e, ev))
        day_events.sort(key=lambda x: (x[0], -(x[1] - x[0])))
        for s, e, ev in day_events:
            free = subtract((s, e), occupied)
            if not free:
                plan["warnings"].append(f"{key} {ev['start']} '{ev['title']}' overlaps another block; not counted")
                continue
            if any_match(cal.get("focus"), ev["title"]):
                note = notes.pop(0) if notes else focus_title_note(cal.get("focus"), ev["title"])
                if note is None:
                    tid, tname, flag, desc = None, None, "focus block: what did you work on?", ev["title"]
                else:
                    tid, tname, flag, rdesc = map_rule(maps.get("focus"), note, day)
                    desc = rdesc or note
                kind = "focus"
            else:
                tid, tname, flag, rdesc = map_rule(maps.get("rule"), ev["title"], day)
                desc, kind = rdesc or ev["title"], "event"
            for a, b in free:
                segs.append({"start": a, "end": b, "task_id": tid, "task": tname, "description": desc,
                             "source": ev["title"], "kind": kind, "flag": flag})
                occupied.append((a, b))
        # open time
        fill = cfg.get("gaps", {}).get("fill", "focus")
        f_tid = f_name = f_desc = None
        if fill == "focus":
            by_task: dict = {}
            for sg in segs:
                if sg["kind"] == "focus" and sg["task_id"]:
                    k = (sg["task_id"], sg["task"], sg["description"])
                    by_task[k] = by_task.get(k, 0) + hours(sg["start"], sg["end"])
            if by_task:
                f_tid, f_name, f_desc = max(by_task, key=by_task.get)
        elif fill.startswith("task:"):
            res = resolve(fill[5:], day)
            if res:
                f_tid, f_name = res
                f_desc = cfg.get("gaps", {}).get("desc", "")
        for a, b in subtract((ws, we), occupied):
            segs.append({"start": a, "end": b, "task_id": f_tid, "task": f_name,
                         "description": f_desc or "", "source": "(open time)", "kind": "gap",
                         "flag": None if f_tid else "open time: which task?"})
        segs.sort(key=lambda x: x["start"])
        # cap at what's still needed (existing entries outside the workday count too)
        need = target - have
        capped = []
        for sg in segs:
            h = hours(sg["start"], sg["end"])
            if need <= EPS:
                break
            if h > need + EPS:
                sg["end"] = sg["start"] + dt.timedelta(hours=need)
                h = need
            need -= h
            capped.append(sg)
        # merge adjacent blocks with the same task and description
        merged = []
        for sg in capped:
            p = merged[-1] if merged else None
            if p and p["end"] == sg["start"] and p["task_id"] == sg["task_id"] and p["description"] == sg["description"] \
                    and p["flag"] == sg["flag"]:
                p["end"] = sg["end"]
                if sg["source"] not in p["source"]:
                    p["source"] += " + " + sg["source"]
            else:
                merged.append(dict(sg))
        plan["days"][key] = {
            "existing_hours": round(have, 2),
            "entries": [{"start": f"{m['start']:%H:%M}", "end": f"{m['end']:%H:%M}",
                         "hours": round(hours(m["start"], m["end"]), 2), "task_id": m["task_id"],
                         "task": m["task"], "description": m["description"], "source": m["source"],
                         "flag": m["flag"]} for m in merged],
        }
        plan["days"][key]["total"] = round(have + sum(e["hours"] for e in plan["days"][key]["entries"]), 2)
    return plan


def top_task(status: dict, exclude: set):
    """(task_id, name, description) of the task with the most hours in status, ignoring
    tasks in exclude (meeting/training/admin rule targets). Description is the one used
    for most of its hours."""
    hrs: dict = {}
    descs: dict = {}
    for ents in status.values():
        for e in ents:
            tid = e.get("task_id")
            if not tid or tid in exclude:
                continue
            hrs[tid] = hrs.get(tid, 0) + e["hours"]
            k = (tid, e["task"], e["description"])
            descs[k] = descs.get(k, 0) + e["hours"]
    if not hrs:
        return None
    tid = max(hrs, key=hrs.get)
    best = max((k for k in descs if k[0] == tid), key=lambda k: (bool(k[2]), descs[k]))
    return best[0], best[1], best[2] or best[1]


def autofill_plan(plan: dict, fallback) -> dict:
    """Resolve every flagged block without asking: give it the fallback task, or drop it
    when there is no fallback. Keeps what was guessed in entry["auto"]."""
    for v in plan["days"].values():
        out = []
        for e in v["entries"]:
            if e["flag"]:
                if not fallback:
                    continue
                e = dict(e, auto=e["flag"], flag=None, task_id=fallback[0], task=fallback[1])
                if not e["auto"].startswith("unmapped") or not e["description"]:
                    e["description"] = fallback[2]
            p = out[-1] if out else None
            if p and p["end"] == e["start"] and p["task_id"] == e["task_id"] and p["description"] == e["description"]:
                p["end"], p["hours"] = e["end"], round(p["hours"] + e["hours"], 2)
                if e["source"] not in p["source"]:
                    p["source"] += " + " + e["source"]
                if e.get("auto"):
                    p["auto"] = p.get("auto") or e["auto"]
            else:
                out.append(e)
        v["entries"] = out
        v["total"] = round(v["existing_hours"] + sum(e["hours"] for e in out), 2)
    return plan


def worked_days(days, status, events, cfg) -> list[dt.date]:
    """Days with some sign of work: a ClickUp entry or a timed calendar event that isn't a break.
    A day with neither (public holiday, leave not on the calendar) is left alone."""
    ign = cfg.get("calendar", {}).get("ignore")
    busy = {e["date"] for e in events if not e.get("all_day") and not any_match(ign, e["title"])}
    return [d for d in days if status.get(d.isoformat()) or d.isoformat() in busy]


def print_plan(plan):
    for s in plan["skipped"]:
        why = s.get("reason") or f"already {s['hours']}h"
        print(f"OK  {s['date']}: {why}, skipped")
    for day, d in plan["days"].items():
        name = dt.date.fromisoformat(day).strftime("%a")
        print(f"\n{name} {day}  (existing {d['existing_hours']}h -> {d['total']}h)")
        for e in d["entries"]:
            flag = f"  !! {e['flag']}" if e["flag"] else ""
            print(f"  {e['start']}-{e['end']}  {e['hours']:4.2f}h  {e['task'] or '?'} [{e['task_id'] or '-'}]"
                  f"  | {e['description']}  <- {e['source']}{flag}")
    for w in plan["warnings"]:
        print(f"note: {w}")


def make_resolver(cu: ClickUp | None, cfg):
    cache: dict = {}

    def resolve(ref, day: dt.date):
        if isinstance(ref, dict):
            name = ref["name"].format(month=day.strftime("%B"), year=day.year)
            lid = list_id(cfg, ref["list"])
            key = (lid, name.lower())
            if key not in cache:
                hit = next((t for t in cu.list_tasks(lid) if t["name"].strip().lower() == name.lower()), None)
                cache[key] = (hit["id"], hit["name"]) if hit else None
            return cache[key]
        if ref not in cache:
            try:
                cache[ref] = (ref, cu.task(ref)["name"]) if cu else (ref, ref)
            except SystemExit:
                cache[ref] = None
        return cache[ref]

    return resolve


# --------------------------------------------------------------------------- commands

def cmd_status(a, cfg):
    cu = ClickUp(cfg)
    days = week_days(cfg, a.week)
    st = fetch_status(cu, cfg, days)
    if a.json:
        print(json.dumps({k: [{**e, "start": f"{e['start']:%H:%M}", "end": f"{e['end']:%H:%M}"} for e in v]
                          for k, v in st.items()}, indent=2))
    else:
        print_status(cfg, st, a.verbose)


def cmd_calendar(a, cfg):
    evs = load_events(cfg, week_days(cfg, a.week), a.source, a.ics)
    text = json.dumps(evs, indent=2)
    if a.out:
        Path(a.out).write_text(text)
        print(f"{len(evs)} events -> {a.out}")
    else:
        print(text)


def cmd_plan(a, cfg):
    cu = ClickUp(cfg)
    days = week_days(cfg, a.week)
    events = json.loads(Path(a.events).read_text()) if a.events else []
    plan = build_plan(cfg, days, fetch_status(cu, cfg, days), events, parse_focus(a.focus, days),
                      make_resolver(cu, cfg))
    if a.out:
        Path(a.out).write_text(json.dumps(plan, indent=2))
    if a.json:
        print(json.dumps(plan, indent=2))
    else:
        print_plan(plan)
        if a.out:
            print(f"\nplan saved to {a.out}")


def cmd_apply(a, cfg):
    apply_plan(ClickUp(cfg), cfg, json.loads(Path(a.plan).read_text()), a.dry_run, a.force)


def apply_plan(cu: ClickUp, cfg, plan: dict, dry_run=False, force=False):
    tz = tz_of(cfg)
    bad = [f"{d} {e['start']} {e['source']}: {e['flag'] or 'no task'}"
           for d, v in plan["days"].items() for e in v["entries"] if not e["task_id"]]
    if bad:
        sys.exit("Plan still has unresolved blocks:\n  " + "\n  ".join(bad))
    days = sorted(dt.date.fromisoformat(d) for d in plan["days"])
    if not days:
        print("Nothing to log.")
        return
    status = fetch_status(cu, cfg, days)
    target = cfg["workday"]["target_hours"]
    for d, v in plan["days"].items():
        have = sum(e["hours"] for e in status.get(d, []))
        add = sum(e["hours"] for e in v["entries"])
        if abs(have - v["existing_hours"]) > EPS:
            print(f"SKIP {d}: ClickUp changed since the plan was made ({v['existing_hours']}h -> {have:.2f}h). Re-plan.")
            continue
        if have + add > target + EPS and not force:
            print(f"SKIP {d}: would reach {have + add:.2f}h (> {target}h). Use --force to allow.")
            continue
        for e in v["entries"]:
            start = at(dt.date.fromisoformat(d), e["start"], tz)
            line = f"{d} {e['start']} {e['hours']:4.2f}h {e['task']} | {e['description']}"
            if dry_run:
                print("DRY  " + line)
            else:
                r = cu.create_entry(e["task_id"], start, e["hours"], e["description"])
                print(f"LOG  {line}  -> {(r.get('data') or {}).get('id', r)}")
    if not dry_run:
        print()
        print_status(cfg, fetch_status(cu, cfg, days))


def load_events(cfg, days, source=None, ics=None) -> list[dict]:
    src = source or cfg.get("calendar", {}).get("source", "graph")
    if src == "graph":
        return graph_events(cfg, days)
    if src == "ics":
        loc = ics or os.environ.get(cfg.get("calendar", {}).get("ics_url_env", "OUTLOOK_ICS_URL"), "")
        if not loc:
            sys.exit("No .ics source: pass --ics PATH_OR_URL or set OUTLOOK_ICS_URL.")
        return ics_events(cfg, days, load_ics(loc))
    sys.exit(f"Unknown calendar source {src!r}")


def cmd_autofill(a, cfg):
    """Unattended fill: plan the week, resolve every flag with the week's top task, log it."""
    cu = ClickUp(cfg)
    tz = tz_of(cfg)
    today = dt.datetime.now(tz).date()
    days = [d for d in week_days(cfg, a.week) if d <= today]
    if not days:
        print("No past workdays in that week yet.")
        return
    try:
        events = load_events(cfg, days, a.source, a.ics)
    except (SystemExit, OSError, urllib.error.URLError) as e:
        sys.exit(f"Calendar unavailable ({e}); nothing logged. Fix the calendar source and re-run.")
    status = fetch_status(cu, cfg, days)
    resolve = make_resolver(cu, cfg)
    rule_ids = set()
    for r in cfg["mappings"].get("rule", []):
        res = resolve(r["task"], days[0]) if isinstance(r["task"], dict) else (r["task"], None)
        if res:
            rule_ids.add(res[0])
    fallback = top_task(status, rule_ids)
    if not fallback:  # nothing logged yet this week: use last week's main task
        prev = [d - dt.timedelta(days=7) for d in days]
        fallback = top_task(fetch_status(cu, cfg, prev), rule_ids)
    worked = worked_days(days, status, events, cfg)
    idle = [d.isoformat() for d in days if d not in worked]
    plan = build_plan(cfg, worked, status, events, parse_focus(a.focus, days), resolve)
    plan = autofill_plan(plan, fallback)
    out = Path(a.out or ROOT / ".plans" / f"autofill-{days[0].isoformat()}.json")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(plan, indent=2))
    print(f"fallback task: {fallback[1]} [{fallback[0]}] | {fallback[2]}" if fallback
          else "fallback task: none (flagged blocks left unlogged)")
    for d in idle:
        print(f"SKIP {d}: no calendar events and no ClickUp entries (holiday or leave?)")
    print_plan(plan)
    for d, v in plan["days"].items():
        for e in v["entries"]:
            if e.get("auto"):
                print(f"auto {d} {e['start']}-{e['end']} -> {e['task']} ({e['auto']}; {e['source']})")
    print()
    apply_plan(cu, cfg, plan, dry_run=a.dry_run)
    if a.dry_run:
        print()
        print_status(cfg, fetch_status(cu, cfg, days))


def cmd_tasks(a, cfg):
    cu = ClickUp(cfg)
    for alias in a.list or list(cfg.get("lists", {})):
        lid = list_id(cfg, alias)
        print(f"== {alias} ({lid})")
        for t in cu.list_tasks(lid):
            print(f"  {t['id']}  {t['name']}  [{t['status']['status']}]")


def cmd_create_task(a, cfg):
    cu = ClickUp(cfg)
    lid = list_id(cfg, a.list)
    hit = next((t for t in cu.list_tasks(lid) if t["name"].strip().lower() == a.name.strip().lower()), None)
    if hit:
        print(f"exists  {hit['id']}  {hit['name']}")
        return
    t = cu.create_task(lid, a.name)
    print(f"created {t['id']}  {t['name']}")


def cmd_ms_login(a, cfg):
    ms_device_login(a.client_id, a.tenant)


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--config", help="path to config.toml (default: repo root)")
    sub = p.add_subparsers(dest="cmd", required=True)
    wk = dict(default="last", help="'last' (default), 'this', or any YYYY-MM-DD in the week")

    s = sub.add_parser("status", help="hours per day for a week")
    s.add_argument("--week", **wk)
    s.add_argument("-v", "--verbose", action="store_true")
    s.add_argument("--json", action="store_true")
    s.set_defaults(fn=cmd_status)

    s = sub.add_parser("calendar", help="export calendar events for a week as JSON")
    s.add_argument("--week", **wk)
    s.add_argument("--source", choices=["graph", "ics"])
    s.add_argument("--ics", help=".ics file path or URL")
    s.add_argument("-o", "--out")
    s.set_defaults(fn=cmd_calendar)

    s = sub.add_parser("plan", help="draft entries for short days")
    s.add_argument("--week", **wk)
    s.add_argument("--events", help="events JSON from `calendar` (or written by hand)")
    s.add_argument("--focus", action="append", help="what a focus block was for, e.g. 'thu=PMO integration' (repeatable, in order)")
    s.add_argument("-o", "--out", help="save the plan JSON here")
    s.add_argument("--json", action="store_true")
    s.set_defaults(fn=cmd_plan)

    s = sub.add_parser("apply", help="log a confirmed plan to ClickUp")
    s.add_argument("plan")
    s.add_argument("--dry-run", action="store_true")
    s.add_argument("--force", action="store_true", help="allow a day to go over the target")
    s.set_defaults(fn=cmd_apply)

    s = sub.add_parser("autofill", help="unattended: plan the week so far, auto-resolve flags, and log it")
    s.add_argument("--week", default="this", help="'this' (default), 'last', or any YYYY-MM-DD in the week")
    s.add_argument("--source", choices=["graph", "ics"])
    s.add_argument("--ics", help=".ics file path or URL")
    s.add_argument("--focus", action="append", help="optional focus notes, as for plan")
    s.add_argument("-o", "--out", help="where to save the plan JSON (default .plans/autofill-<monday>.json)")
    s.add_argument("--dry-run", action="store_true", help="plan and report, log nothing")
    s.set_defaults(fn=cmd_autofill)

    s = sub.add_parser("tasks", help="list tasks in the configured lists")
    s.add_argument("list", nargs="*", help="list aliases from config.toml or list IDs")
    s.set_defaults(fn=cmd_tasks)

    s = sub.add_parser("create-task", help="create a task (no-op if the name already exists)")
    s.add_argument("--list", required=True, help="list alias from config.toml or a list ID")
    s.add_argument("--name", required=True)
    s.set_defaults(fn=cmd_create_task)

    s = sub.add_parser("ms-login", help="sign in to Microsoft (device code) for calendar access")
    s.add_argument("--client-id", default=os.environ.get("MS_CLIENT_ID", DEFAULT_MS_CLIENT_ID))
    s.add_argument("--tenant", default=os.environ.get("MS_TENANT_ID", "organizations"))
    s.set_defaults(fn=cmd_ms_login)

    a = p.parse_args(argv)
    a.fn(a, load_config(a.config))


if __name__ == "__main__":
    main()
