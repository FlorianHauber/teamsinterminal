#!/bin/sh
# Installs `teams` into ~/.local/bin. No root needed.
set -e
DIR="$(cd "$(dirname "$0")" && pwd)"
[ -x "$DIR/.venv/bin/python" ] || python3 -m venv "$DIR/.venv"
"$DIR/.venv/bin/python" -m pip install -q --upgrade playwright
# Use the installed Google Chrome if there is one, else download Chromium into ~/.cache/ms-playwright
if ! command -v google-chrome-stable >/dev/null 2>&1 && ! command -v google-chrome >/dev/null 2>&1; then
  "$DIR/.venv/bin/playwright" install chromium
fi
mkdir -p "$HOME/.local/bin"
cat > "$HOME/.local/bin/teams" <<EOS
#!/bin/sh
exec "$DIR/.venv/bin/python" "$DIR/teams_cli.py" "\$@"
EOS
chmod +x "$HOME/.local/bin/teams" "$DIR/teams_cli.py"
echo "installed: $HOME/.local/bin/teams"
