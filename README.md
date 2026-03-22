# Timex

Minimal time tracker for macOS. Dark TUI in a native window, keyboard-driven, no Electron.

![Python](https://img.shields.io/badge/Python-3.13-blue)
![Platform](https://img.shields.io/badge/Platform-macOS-lightgrey)
![License](https://img.shields.io/badge/License-MIT-green)

## Features

- Timer with start / pause / resume
- Multi-project support with independent histories
- Watch mode — automatic activity tracking via screenshots or window focus detection
- AI task labeling (analyzes screenshots to name tasks automatically)
- Export to Google Sheets (OAuth2) and Excel (.xlsx)
- Session history by date, statistics
- Customization: accent color (10 presets + custom HEX), timezone, reminder intervals
- Menu bar widget showing timer status

## Architecture

3-process model:

1. **launcher.py** — native frameless window (pywebview, 400x720)
2. **serve.py** — WebSocket bridge (textual-serve / uvicorn)
3. **timex.py** — main TUI app (Textual + Rich)

Separately: **menubar.py** — menu bar widget (rumps), reads `~/.timex/state.json`.

## Tech Stack

Python 3.13, Textual, Rich, textual-serve, pywebview, rumps, openpyxl, gspread, PyObjC

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
      sheets_config.json
```

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
| `/watch` | Auto-track activity |
| `/export` | Export to Sheets or Excel |
| `/project` | Switch projects |
| `/date` | Browse history |
| `/stats` | View statistics |
| `/color` | Change accent color |
| `/timezone` | Set timezone |
| `/notification` | Configure reminders |
| `/help` | Show help |

## Setup

### Requirements

```
pip install textual rich textual-serve pywebview rumps openpyxl
```

Optional (for Google Sheets export):
```
pip install gspread google-auth google-auth-oauthlib
```

Optional (for AI task labeling):
- Set `openai_api_key` in `~/.timex/config.json`

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

## License

MIT
