# clickup-timesheet

Checks a week of ClickUp hours against an 8h/day target and fills short days from the
Outlook calendar. A script does the API work. A Claude Code skill
(`.claude/skills/clickup-timesheet`) handles mapping and confirmation. It never logs anything
without your OK. A second skill (`clickup-timesheet-autofill`) is the unattended Friday fill.

## Use it

In a Claude Code session with this repo checked out:

> check my ClickUp hours for last week

or run the script directly:

```bash
python3 scripts/timesheet.py status --week last -v
python3 scripts/timesheet.py calendar --week last -o .plans/events.json
python3 scripts/timesheet.py plan --week last --events .plans/events.json \
  --focus "thu=PMO integration" --focus "fri=Data Foundation planning" -o .plans/plan.json
python3 scripts/timesheet.py apply .plans/plan.json --dry-run
python3 scripts/timesheet.py apply .plans/plan.json
```

**Unattended fill.** `autofill --week this` plans the week so far and logs it without asking.
Anything the rules can't place goes to the week's main (non-meeting) task, and those rows are
reported. Hours already in ClickUp count and are never changed. It skips days with an all-day
out-of-office event, and days with no calendar events and no entries. Add `--dry-run` to
preview. The `clickup-timesheet-autofill` skill runs it from a scheduled task every Friday evening.

Other commands: `tasks [list...]` lists task IDs, `create-task --list admin --name "..."`,
and `ms-login` signs in to Microsoft.

## Files

| File | What it holds |
|---|---|
| `config.toml` | Workspace/user IDs, timezone, workday, list aliases, focus/ignore patterns, gap rule |
| `mappings.toml` | Event title → task rules, and focus note → task rules. Grows as you answer questions |
| `scripts/timesheet.py` | Standard-library Python 3.11+ CLI |
| `tests/` | `python3 -m unittest discover -s tests` |

## Credentials

Environment variables. Set them as cloud environment secrets, never in the repo:

| Variable | For |
|---|---|
| `CLICKUP_API_KEY` | ClickUp personal API token (required) |
| `MS_REFRESH_TOKEN` | Optional: Microsoft refresh token, so calendar access survives a new container |
| `OUTLOOK_ICS_URL` | Optional: published Outlook calendar `.ics` link, as an alternative to Graph |

### Calendar access

**Graph (default).** `timesheet.py ms-login` starts Microsoft's device-code sign-in with the
public *Microsoft Graph Command Line Tools* client, asking only for `Calendars.Read`.
You open the link, enter the code and approve. The token is saved to
`~/.config/clickup-timesheet/ms_token.json` (mode 600). Cloud containers are temporary, so
either sign in again each session or copy `refresh_token` from that file into the
`MS_REFRESH_TOKEN` secret. It stays valid while it's used at least every 90 days. The network
policy must allow `graph.microsoft.com` and `login.microsoftonline.com`.

If your tenant blocks user consent for that client, an admin can approve it, or register an
app with delegated `Calendars.Read` and pass `--client-id` / set `MS_CLIENT_ID` and `MS_TENANT_ID`.
`az account get-access-token --resource-type ms-graph` is tried last. The Azure CLI's token
usually lacks calendar scope, though, so don't rely on it.

**.ics feed.** In Outlook on the web: Settings → Calendar → Shared calendars → Publish a
calendar → ICS link. Set it as `OUTLOOK_ICS_URL` and set `source = "ics"` in `config.toml`.
Anyone with the link can read your calendar, so treat it as a secret. The network policy must
allow `outlook.office365.com`. The parser covers what Outlook emits: daily/weekly/monthly
recurrences, exceptions, moved occurrences and Windows time zone names.

**Screenshots.** Attach week-view screenshots that show the time column. The skill transcribes them.

## How the plan is built

For each day under target, within the workday (09:00–17:00 Manila):

1. Existing ClickUp entries are kept and count toward the target.
2. Calendar events are placed earliest first. When events overlap, the longer one wins, and
   events that are fully covered are noted and not counted.
3. "Focus time" blocks take their task from `--focus` notes. Other events use `mappings.toml`.
4. Breaks (`ignore` patterns) and empty time are open time, filled per `[gaps]`. By default
   that's the day's main focus task.
5. Neighbouring blocks with the same task and description are merged, and the day is capped at the target.

Anything unresolved is flagged. `apply` refuses to run until nothing is flagged.
