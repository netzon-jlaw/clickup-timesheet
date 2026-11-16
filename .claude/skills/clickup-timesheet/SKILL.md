---
name: clickup-timesheet
description: Check J. Law's ClickUp hours for a week against the 8h/day target and fill short days from their Outlook calendar. Use when asked to check, fill, or update ClickUp hours or timesheets ("check my hours for last week", "fill my ClickUp hours", "/clickup-timesheet"), or when a scheduled run asks for this skill.
---

# ClickUp timesheet

Everything deterministic goes through `scripts/timesheet.py` (run from the repo root with
`python3 -I scripts/timesheet.py ...`). Your job is the judgment: reading the calendar,
resolving what the rules can't, and getting confirmation before anything is written.

Needs `CLICKUP_API_KEY` in the environment. Never print or echo it, or any Microsoft token.

## Inputs (all optional)

- **Week**: `last` (default), `this`, or any date in the week.
- **Calendar source**: Graph (default), an `.ics` file/URL, or screenshots the user attaches.
- **Focus notes**: what each "Focus time" block was for, e.g. `thu=PMO integration`.
  Several on one day are used in order.
- **Gap rule**: what to do with open time (default: that day's main focus task, from `config.toml`).
- **Dry run**: plan only.

## Flow

1. **Status.** `status --week <w> -v`. If every day is at target, say so and stop.

2. **Calendar** for the short days, first source that works:
   - `calendar --week <w> -o .plans/events.json` (Graph; needs a Microsoft sign-in, see below).
   - `calendar --source ics --ics <path-or-url> -o .plans/events.json`, or with `OUTLOOK_ICS_URL` set.
   - Screenshots: transcribe each visible event into `.plans/events.json` as
     `[{"date":"YYYY-MM-DD","start":"HH:MM","end":"HH:MM","title":"..."}]`.
     Read times off the time column. If the time column isn't visible, ask for a screenshot that
     shows it, or state the assumption (day = workday start..end in `config.toml`) in the plan.

   If Graph fails with no credentials, run `ms-login` in the background, read its output file for
   the device code, and give the user the URL and code. Continue once it reports "Signed in".

3. **Plan.** `plan --week <w> --events .plans/events.json --focus "<day>=<note>" ... -o .plans/plan.json`.
   Show it as one table per day: time, ClickUp task, hours, description. Mark flagged rows
   (`!!` in the output) with ⚠️. Mention overlap notes briefly.

4. **Resolve flags.** Ask only about ⚠️ rows, all in one message:
   - *unmapped* event: which task?
   - *focus block*: what was worked on?
   - *open time*: which task?
   - *rule task not found*: offer to create it.

   Turn each answer into a rule so it isn't asked again: add a `[[rule]]` (event title) or
   `[[focus]]` (focus note) to `mappings.toml`. Create missing tasks with
   `create-task --list <alias> --name "<name>"` (it reuses an existing task with the same name).
   Find task IDs with `tasks [alias...]`. Re-run `plan` until nothing is flagged.

5. **Confirm, then log.** Show the final plan and wait for an explicit yes.
   `apply .plans/plan.json` (or `--dry-run` first). `apply` skips any day that changed in ClickUp
   since the plan was made, and any day that would go over target unless `--force`.

6. **Verify and save.** `apply` prints the week's totals at the end. Report them. If
   `mappings.toml` changed and this is a git checkout, commit it ("Add timesheet mappings: ...")
   and push to the session's working branch.

## Scheduled / headless runs

Do steps 1–3 only. Never run `apply` without a person's confirmation in the conversation.
End with the plan and the questions from step 4. The one exception is a run that asks for the
`clickup-timesheet-autofill` skill, which J. Law has authorised to log without confirmation.

## Rules

- Never write to ClickUp before the user confirms the plan shown to them.
- Never log to a day that is already at target.
- Descriptions follow the user's style: short, plain ("PMO integration", "Netzon Infra meeting").
- Training hours go to the month's task in Trainings (`{month} Hours`). If that task doesn't
  exist yet for the month, offer to create it.
