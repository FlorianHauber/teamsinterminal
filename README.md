# teamsinterminal

Microsoft Teams files and Forms tests in the terminal.

- `teams` (`teams_cli.py`): signs in with a normal school account in a hidden Chrome window (Playwright). Lists and downloads team files from SharePoint and fills Forms tests question by question. Needs no app registration or admin consent.
- `teams-tui` (`teams_tui.py`): Textual TUI over Microsoft Graph (MSAL login). Needs an Azure app client id via `TEAMS_TUI_CLIENT_ID`.

## Install

Needs Linux or macOS with Python 3.10+ and git. No root needed.

```sh
curl -fsSL https://raw.githubusercontent.com/FlorianHauber/teamsinterminal/main/install.sh | sh
```

The script does this:

1. Clones the repo into `~/.local/share/teamsinterminal` (or updates it).
2. Creates a virtual environment there and installs `requirements.txt`.
3. Downloads Chromium for Playwright, unless Google Chrome is installed.
4. Puts `teams` and `teams-tui` into `~/.local/bin`.

From a clone instead:

```sh
git clone https://github.com/FlorianHauber/teamsinterminal.git
cd teamsinterminal
./install.sh
```

Run the same command again to update.

## Use

```sh
teams login     # e-mail, password, MFA in the terminal
teams           # full-screen view, press ? for keys
teams shell     # command prompt, type "help"
```

The Courses tab lists only the teams you are a member of for the current school year. The list comes from Teams itself and is cached, so it shows at once on the next start. Press `A` to show earlier years and archived teams, `r` to reload. In the shell, `teams --all` and `teams --sites` widen the list.

Forms tests run full screen. ↑ ↓ pick an answer, Space ticks it, Enter goes on, ← goes back. A progress bar shows answered questions. Before submitting you see all answers, with missing required ones marked. After submitting, the score shows when the quiz reports one.

## Uninstall

```sh
rm -rf ~/.local/bin/teams ~/.local/bin/teams-tui ~/.local/share/teamsinterminal
```
