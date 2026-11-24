# teamsinterminal

Microsoft Teams files and Forms tests in the terminal.

- `teams_cli.py`: signs in with a normal school account in a hidden Chrome window (Playwright). Lists and downloads team files from SharePoint and fills Forms tests question by question. Needs no app registration or admin consent.
- `teams_tui.py`: Textual TUI over Microsoft Graph (MSAL login). Needs an Azure app client id via `TEAMS_TUI_CLIENT_ID`.

## Install

```sh
python -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/playwright install chromium
.venv/bin/python teams_cli.py login
```
