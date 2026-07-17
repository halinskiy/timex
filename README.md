# Timex

Minimal time tracker for macOS. Dark TUI in a native window, keyboard-driven, no Electron.

![Python](https://img.shields.io/badge/Python-3.13-blue)
![Platform](https://img.shields.io/badge/Platform-macOS-lightgrey)
![License](https://img.shields.io/badge/License-MIT-green)

## Features

- Timer with start / pause / resume
- Multi-project support with independent histories
- Reports for a day, a week, a month or a custom range:
  - a self-contained HTML page that opens in a browser and can be sent to a client
  - an Excel workbook (.xlsx), also embedded in the page as a download
- Session history by date, statistics
- Automatic backups of every project's history
- Customization: accent color (10 presets + custom HEX), timezone, reminder intervals
- Menu bar widget showing timer status

## Reports

`/export` picks a period, then writes to `~/Downloads`:

- **Visual report (.html)** — one file with everything inlined: no network requests, no
  server, works offline forever. Summary tiles, the longest tasks, hours per day, and the
  full task table. The workbook is embedded as a data URI, so the page carries its own
  download. Follows the reader's light/dark setting.
- **Excel (.xlsx)** — `Report` sheet (summary + charts) and `Detail` sheet (one row per
  task, autofiltered). Durations appear both as `HH:MM:SS` and as decimal hours, because
  invoicing multiplies the decimal ones.

A single-day report charts the whole week around that day, so one day reads as one day
among seven rather than a full-width block.

## Architecture

3-process model:

1. **launcher.py** — native frameless window (pywebview, 400x720)
2. **serve.py** — WebSocket bridge (textual-serve / uvicorn)
3. **timex.py** — main TUI app (Textual + Rich)

Separately: **menubar.py** — menu bar widget (rumps), reads `~/.timex/state.json`.

The widget runs from the bundle's own second executable (`MacOS/TimexMenubar`, wired up
through `SCRIPT_MAP` in `__boot__.py`). A menu bar item needs an app-bundled binary — a
bare `python` cannot own one.

## Tech Stack

Python 3.13, Textual, Rich, textual-serve, pywebview, rumps, openpyxl, PyObjC

## Data

All data stored in `~/.timex/`:

```
~/.timex/
  state.json          # current session
  history.json        # completed sessions
  config.json         # preferences
  active_project      # current project name
  projects/           # per-project directories
    ProjectName/
      state.json
      history.json
  backups/            # automatic snapshots of projects/
    20260716-152459-daily/
```

`history.json` is the only record a session ever gets, so the whole `projects/` tree is
snapshotted on first launch each day and before anything destructive (deleting a session
or a project). The last 30 snapshots are kept; restoring is a plain folder copy back.

## Commands

| Command | Description |
|---|---|
| `/start` | Start timer |
| `/pause` | Pause timer |
| `/resume` | Resume timer |
| `/new` | Save session, start fresh |
| `/edit` | Rename or delete tasks |
| `/add 30m` | Add time to session |
| `/remove 10m` | Remove time from session |
| `/export` | Report a period: visual .html page or .xlsx |
| `/project` | Switch projects |
| `/date` | Browse history |
| `/stats` | View statistics |
| `/color` | Change accent color |
| `/timezone` | Set timezone |
| `/notification` | Configure reminders |
| `/update` | Check for a new version |
| `/help` | Show help |

## Setup

### Requirements

```
pip install textual rich textual-serve pywebview rumps openpyxl
```

### Run

```bash
python timex.py          # TUI only (terminal)
python launcher.py       # native window + menu bar
```

### Build (.app)

Requires py2app:
```bash
pip install py2app
python setup.py py2app
```

The built app is signed and notarized with `notarize.sh`. Editing anything inside the
bundle invalidates the signature, so the script has to be re-run for any build meant to
ship — and updates are delivered as a fresh signed app rather than by patching files in
place.

## License

MIT
