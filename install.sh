#!/bin/sh
# Installs `teams` and `teams-tui` into ~/.local/bin. No root needed.
#
#   curl -fsSL https://raw.githubusercontent.com/FlorianHauber/teamsinterminal/main/install.sh | sh
#
# Run from a clone, it installs that clone. Run through curl, it clones the
# repo into ~/.local/share/teamsinterminal (or updates an existing clone).
set -e

REPO="https://github.com/FlorianHauber/teamsinterminal.git"
BIN="$HOME/.local/bin"

die() { echo "error: $*" >&2; exit 1; }

command -v python3 >/dev/null 2>&1 || die "python3 not found. Install Python 3.10 or newer."
python3 -c 'import sys; sys.exit(sys.version_info < (3, 10))' || die "Python 3.10 or newer is required."
python3 -c 'import venv, ensurepip' 2>/dev/null || die "python3 venv module missing (Debian/Ubuntu: sudo apt install python3-venv)."

# Find the source: this script's folder if it holds the code, else a clone.
DIR=""
case "$0" in
  */install.sh|install.sh) DIR="$(cd "$(dirname "$0")" && pwd)" ;;
esac
if [ -z "$DIR" ] || [ ! -f "$DIR/teams_cli.py" ]; then
  command -v git >/dev/null 2>&1 || die "git not found. Install git first."
  DIR="${TEAMS_DIR:-$HOME/.local/share/teamsinterminal}"
  if [ -d "$DIR/.git" ]; then
    echo "updating $DIR"
    git -C "$DIR" pull -q --ff-only
  else
    echo "cloning into $DIR"
    mkdir -p "$(dirname "$DIR")"
    git clone -q "$REPO" "$DIR"
  fi
fi

echo "creating virtual environment in $DIR/.venv"
[ -x "$DIR/.venv/bin/python" ] || python3 -m venv "$DIR/.venv"
"$DIR/.venv/bin/python" -m pip install -q --upgrade pip
"$DIR/.venv/bin/python" -m pip install -q -r "$DIR/requirements.txt"

# Use the installed Google Chrome if there is one, else download Chromium into ~/.cache/ms-playwright
if ! command -v google-chrome-stable >/dev/null 2>&1 && ! command -v google-chrome >/dev/null 2>&1; then
  echo "downloading Chromium for Playwright"
  "$DIR/.venv/bin/playwright" install chromium
fi

mkdir -p "$BIN"
for pair in "teams:teams_cli.py" "teams-tui:teams_tui.py"; do
  name="${pair%%:*}"; script="${pair#*:}"
  cat > "$BIN/$name" <<EOS
#!/bin/sh
exec "$DIR/.venv/bin/python" "$DIR/$script" "\$@"
EOS
  chmod +x "$BIN/$name"
  echo "installed: $BIN/$name"
done

case ":$PATH:" in
  *":$BIN:"*) ;;
  *) echo
     echo "$BIN is not in your PATH. Add this line to ~/.bashrc or ~/.zshrc:"
     echo "  export PATH=\"\$HOME/.local/bin:\$PATH\"" ;;
esac

echo
echo "done. Sign in with: teams login"
