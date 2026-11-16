---
name: clickup-timesheet-autofill
description: Unattended weekly fill of J. Law's ClickUp hours to 8h/day from the Outlook calendar, logging without asking. Use when a scheduled run (Friday evening) or the user asks to "auto-fill", "autofill" or "fill my hours without asking" for a week. For an interactive check-and-confirm fill, use clickup-timesheet instead.
---

# ClickUp timesheet autofill

J. Law has authorised this skill to log hours to ClickUp **without confirmation**. It does
the same as `clickup-timesheet` but answers every question with a best-effort default, then logs.
Everything goes through `scripts/timesheet.py autofill` (from the repo root, `python3 -I`).

Needs `CLICKUP_API_KEY`, and `OUTLOOK_ICS_URL` for the calendar. Never print either value.

## Run

1. Check that both variables are set (print only "set" / "missing"). If `CLICKUP_API_KEY` is missing,
   stop and report.
2. `python3 -I scripts/timesheet.py autofill --week this` (default week is the current one,
   up to today). Add `--week <date>` if the user named another week, and `--dry-run` if they asked
   for a preview only.
3. Report (see below). Don't commit anything: the run changes no repo files except `.plans/`,
   which is ignored.

Don't run `create-task`, edit `mappings.toml` or `config.toml`, or pass `--force` in an unattended run.

## What `autofill` does

- Reads the week's ClickUp entries and calendar. **Entries already in ClickUp count**, whatever
  logged them (by hand, a timer, another tool): they count toward 8h, their time slots are kept,
  and they are never changed or deleted. Days already at 8h are skipped.
- Plans each short day like `plan`: mapped meetings to their tasks, Focus blocks from the note in
  the title ("Focus time (Data Foundation Planning)") or `--focus`, breaks and cancelled meetings ignored.
- Anything that would be a question gets the **fallback task**: the non-meeting task with the
  most hours that week (last week's if this week is empty). Unmapped meetings keep their title as
  the description. These rows are printed as `auto ...`.
- Leaves alone: days with an all-day out-of-office event, days with no calendar events and no
  ClickUp entries (likely a holiday or unlisted leave), and future days.
- Never goes over 8h a day. Re-checks ClickUp right before logging and skips a day that changed.
- If the calendar can't be read, logs nothing and exits with an error.

## Report

End with a short summary for J. Law:
- Per day: hours before → after, or why it was skipped.
- The fallback task used, and every `auto` row (date, time, task, original event), so they can fix
  any wrong guess in ClickUp.
- Suggested `mappings.toml` rules for unmapped meetings that recur, to add in an interactive session.
- Any error, quoted.
