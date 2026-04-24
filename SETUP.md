# Revenue Model Agent — New Machine Setup

This guide covers what to change when moving the agent to a different computer.

---

## TL;DR

1. Copy the entire project folder (e.g. `Revenue Model/`) to the new machine.
2. Install Python 3.10+ and the dependencies (`py -m pip install -r requirements.txt`).
3. Copy `config.example.json` → `config.json` and edit the two `<REPLACE_ME>` values.
4. Confirm Outlook is installed and signed in.
5. Add the project folder to Excel Trust Center as a Trusted Location.
6. Double-click `Run Revenue Model Agent.bat`.

Nothing in the Python source needs editing. All machine-specific values live in `config.json`.

---

## Step 1 — Install Python

- Download Python 3.10 or newer from [python.org](https://www.python.org/downloads/).
- During install, tick **"Add python.exe to PATH"** and **"Install launcher for all users"**.
- Verify in a new terminal: `py --version` should print `Python 3.10.x` or newer.

## Step 2 — Install dependencies

From the project folder:

```
py -m pip install -r requirements.txt
```

Installs `pywin32`, `pywinauto`, `python-dateutil`, `matplotlib`, `numpy`. `pywinauto` drives the Prophix Analyzer pane in Step 4 (no manual click needed). The batch file refuses to run if `py` is not on PATH.

## Step 3 — Create `config.json`

```
copy config.example.json config.json
```

Then open `config.json` in a text editor and fill in the two `<REPLACE_ME>` fields.

### Config keys — which to change

| Key | Change on new machine? | Notes |
|---|---|---|
| `project_root` | **YES** — always change | Absolute path to the Revenue Model folder on the new machine. Use double backslashes in JSON (e.g. `"C:\\Users\\name\\Documents\\Revenue Model"`). |
| `archive_dir` | Optional | Leave as `null` to use `<project_root>\Archive`. Set only if you want archives somewhere else (e.g. a network drive). |
| `log_dir` | Optional | Leave as `null` to use `<project_root>\agent\logs`. Set only if you want logs somewhere else. |
| `email_recipient` | **YES** — if a different person now owns the process | Email that receives failure/variance alerts AND the exec summary draft (the draft is addressed to this mailbox, not sent). Default: `asbiei@prophix.com`. |
| `prophix_wait_seconds` | Rarely | Seconds to sleep after the Prophix refresh kicks off. Default 600 (10 min). Increase on slower machines or if Prophix runs large refreshes. |
| `file_prefix` / `file_suffix` / `file_ext` / `file_date_format` | **No** (unless the workbook naming convention changes) | Parses filenames like `Revenue Model - MMDDYYYY (Internal).xlsm`. Only touch if the company renames the file pattern. |

### Example `config.json` for a new user

```json
{
  "project_root": "C:\\Users\\jsmith\\Documents\\Revenue Model",
  "archive_dir": null,
  "log_dir": null,
  "email_recipient": "jsmith@prophix.com",
  "prophix_wait_seconds": 600,
  "file_prefix": "Revenue Model - ",
  "file_suffix": " (Internal)",
  "file_ext": ".xlsm",
  "file_date_format": "%m%d%Y"
}
```

## Step 4 — Outlook

- Outlook must be installed (desktop client, not web).
- Outlook must be signed into the mailbox that should receive the draft. The `email_recipient` in `config.json` is **who the draft is addressed to**; Outlook's default profile decides **which mailbox the draft lands in**.
- No API keys or OAuth needed — `pywin32` talks to Outlook via COM.

## Step 5 — Excel Trust Center (one-time per user)

Excel must be allowed to run the Step 3 formula-roll macro without prompting.

1. Open Excel → File → Options → Trust Center → Trust Center Settings → Trusted Locations.
2. Click **Add new location**, browse to the project folder (value of `project_root`), tick **Subfolders of this location are also trusted**, click OK.
3. Close Excel.

## Step 6 — Test the read-only path

Before a real run, validate the environment with:

```
"Run Revenue Model Agent.bat" --only-step 5
```

This touches no state (opens the workbook read-only, checks Prophix tabs are populated, closes) and will surface any misconfiguration without side effects.

## Step 7 — First real run

Double-click `Run Revenue Model Agent.bat`. Step 4 is now fully automated: the agent opens Excel, activates `Units Bookings DV`, presses `Alt+N+Y+1` itself, clicks **Refresh → All Sheets** in the Prophix Analyzer pane, then sleeps ~10 min. Do not touch the keyboard or mouse while Step 4 runs — it will steal focus on Excel and any input from you could derail the keystroke chord. When the agent finishes, check Outlook Drafts.

---

## What lives where

| File | Purpose | Edit on new machine? |
|---|---|---|
| `config.json` | Machine-specific paths + email | **Yes** — only file you edit |
| `config.example.json` | Template for `config.json` | No |
| `requirements.txt` | Python dependencies | No |
| `Run Revenue Model Agent.bat` | Launcher (preflight checks + runs agent) | No |
| `agent/config.py` | Data-specific config (tabs, rows, tolerances) — reads machine values from `config.json` | No |
| `agent/*.py` | Step modules | No |
| `agent-backup/` | Frozen snapshot of the agent before JSON refactor | No — keep as rollback |
| `tests/` | Unit tests (`py tests/test_*.py`) | No |
| `spec.md` / `spec.docx` | Full specification | No |

## Troubleshooting

| Symptom | Fix |
|---|---|
| `FileNotFoundError: Missing .../config.json` | Copy `config.example.json` to `config.json` and edit it. |
| `ValueError: config.json is missing required key 'X' or still has a <REPLACE_ME> placeholder` | Finish editing `config.json` — don't leave `<REPLACE_ME>` placeholders. |
| `PermissionError` during Step 1 | The workbook is open in Excel. Close Excel and re-run. |
| `ImportError: No module named win32com` | Run `py -m pip install -r requirements.txt`. |
| Step 3 fails with macro security prompt | Add the project folder to Excel Trust Center (see Step 5 above). |
| Draft never appears in Outlook | Outlook isn't installed or isn't signed in. Open Outlook manually first, then re-run `--only-step 7`. |
