#!/usr/bin/env python3
"""
teams - Microsoft Teams files and Forms tests in the terminal.

Signs in with your normal school account in a hidden Chrome window (Playwright),
then reads what the web pages show you:

  * team files      from the team's SharePoint site (where Teams stores them)
  * Forms tests     by filling the real forms.office.com page question by question

No app registration, no admin consent, no root: the session is the same as
logging in at office.com in a browser. Downloads go to ~/Moodle/<subject>.
"""

import cmd
import difflib
import getpass
import json
import os
import re
import shlex
import shutil
import subprocess
import sys
import time
import urllib.parse
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path

VERSION = "1.3"
APP = "teams-cli"
CONFIG_DIR = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config") / APP
CONFIG_FILE = CONFIG_DIR / "config.json"
DATA_DIR = Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share") / APP
PROFILE_DIR = DATA_DIR / "chrome-profile"
STATE_DIR = Path(os.environ.get("XDG_STATE_HOME") or Path.home() / ".local" / "state") / APP
HISTORY_FILE = STATE_DIR / "history"
ASSIGN_CACHE = STATE_DIR / "assignments.json"
MOODLE_CFG = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config") / "moodle-cli" / "config.json"

# --------------------------------------------------------------------------- #
#  Output helpers (same look as moodle-cli)
# --------------------------------------------------------------------------- #

USE_COLOR = sys.stdout.isatty() and not os.environ.get("NO_COLOR") and os.environ.get("TERM") != "dumb"


def _c(code):
    return lambda t: f"\033[{code}m{t}\033[0m" if USE_COLOR else str(t)


bold, dim, red, green, yellow, blue, magenta, cyan = (
    _c("1"), _c("2"), _c("31"), _c("32"), _c("33"), _c("34"), _c("35"), _c("36"))


def err(msg):
    print(red("error: ") + str(msg), file=sys.stderr)


IN_TUI = False


def info(msg):
    if not IN_TUI:               # the full-screen view shows progress in its status line instead
        print(dim(msg), file=sys.stderr)


def width():
    return max(40, min(shutil.get_terminal_size((100, 24)).columns, 110))


def page(text):
    if not sys.stdout.isatty() or text.count("\n") + 2 < shutil.get_terminal_size((100, 24)).lines:
        print(text)
        return
    pager = os.environ.get("PAGER")
    try:
        p = subprocess.Popen(shlex.split(pager) if pager else ["less", "-R", "-F", "-X"], stdin=subprocess.PIPE)
        p.communicate(text.encode("utf-8", "replace"))
    except OSError:
        print(text)


def human(n):
    n = float(n or 0)
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024


def fmt_time(iso):
    if not iso:
        return ""
    try:
        return datetime.fromisoformat(iso.replace("Z", "+00:00")).astimezone().strftime("%d %b %Y %H:%M")
    except ValueError:
        return iso


def safe_name(name, limit=150):
    name = re.sub(r"[/\\\x00-\x1f]", "_", str(name or "")).strip().strip(".")
    return (name or "_")[:limit]


def confirm(question, default=False):
    if not sys.stdin.isatty():
        return default
    try:
        ans = input(f"{question} {'[Y/n]' if default else '[y/N]'} ").strip().lower()
    except EOFError:
        return default
    return default if not ans else ans in ("y", "yes", "j", "ja")


def args_of(line):
    try:
        return shlex.split(line)
    except ValueError:
        return line.split()


def parse_nums(tokens, maximum):
    """'1 3 5-7' -> [1,3,5,6,7]; 'all' -> 1..maximum"""
    out = []
    for t in tokens:
        if t == "all":
            return list(range(1, maximum + 1))
        m = re.fullmatch(r"(\d+)-(\d+)", t)
        if m:
            out += range(int(m[1]), int(m[2]) + 1)
        elif t.isdigit():
            out.append(int(t))
        else:
            raise TeamsError(f"not a number: {t}")
    bad = [n for n in out if not 1 <= n <= maximum]
    if bad:
        raise TeamsError(f"no item {bad[0]} (list has {maximum})")
    return list(dict.fromkeys(out))


def norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


# --------------------------------------------------------------------------- #
#  School years and course names ("2026_27_2IHIF_ ENGLISH" -> English, 26/27)
# --------------------------------------------------------------------------- #

YEAR_PAIR_RE = re.compile(r"(?<!\d)(?:20)?(\d{2})\s*[/_\-. ]?\s*(?:20)?(\d{2})(?!\d)")
CLASS_RE = re.compile(r"(?<![A-Za-z0-9])\d[A-Za-z]{3,5}(?![A-Za-z])")


def school_year(now=None):
    """Start year of the current school year (it starts in September; August already counts)."""
    now = now or datetime.now()
    return now.year - (0 if now.month >= 8 else 1)


def school_label(y=None):
    y = school_year() if y is None else y
    return f"{y % 100:02d}/{(y + 1) % 100:02d}"


def title_years(title):
    """School years named in a team title: '2526', '25-26', '2025/26', '2026_27' -> {2025} or {2026}."""
    out = set()
    for m in YEAR_PAIR_RE.finditer(str(title)):
        a, b = int(m[1]), int(m[2])
        if b == (a + 1) % 100:
            out.add(2000 + a)
    return out


def class_code(title):
    m = CLASS_RE.search(YEAR_PAIR_RE.sub(" ", str(title)))
    return m[0].upper() if m else ""


def course_name(title):
    """Readable subject name: year and class code removed, SHOUTING turned into Title Case."""
    t = YEAR_PAIR_RE.sub(lambda m: " " if int(m[2]) == (int(m[1]) + 1) % 100 else m[0], str(title))
    t = CLASS_RE.sub(" ", t)
    t = re.sub(r"\s+", " ", re.sub(r"[_\-/]+", " ", t)).strip(" .")
    if not t:
        return class_code(title) or str(title)
    words = []
    for w in t.split(" "):
        if w.isupper() and len(w) > 5 or w.islower() and len(w) > 3:
            w = w.capitalize()
        words.append(w)
    return " ".join(words)


# --------------------------------------------------------------------------- #
#  Single key presses (for the Forms runner)
# --------------------------------------------------------------------------- #

_ESC_KEYS = {b"[A": "up", b"[B": "down", b"[C": "right", b"[D": "left", b"OA": "up", b"OB": "down",
             b"OC": "right", b"OD": "left", b"[H": "home", b"[F": "end", b"[Z": "btab"}


def read_key():
    """One key press as a name ('up', 'enter', 'space', 'esc', 'tab', 'backspace') or the character."""
    import select
    import termios
    import tty
    fd = sys.stdin.fileno()
    old = termios.tcgetattr(fd)
    try:
        tty.setraw(fd)
        ch = os.read(fd, 1)
        if ch == b"\x1b":
            seq = b""
            while select.select([fd], [], [], 0.04)[0]:
                seq += os.read(fd, 1)
                if (len(seq) >= 2 and seq[-1:].isalpha()) or seq[-1:] == b"~":
                    break
            return _ESC_KEYS.get(seq, "esc" if not seq else "")
        if ch in (b"\r", b"\n"):
            return "enter"
        if ch == b" ":
            return "space"
        if ch == b"\t":
            return "tab"
        if ch in (b"\x7f", b"\x08"):
            return "backspace"
        if ch == b"\x03":
            raise KeyboardInterrupt
        if ch == b"\x04":
            raise EOFError
        if ch[0] >= 0xC0:                       # rest of a UTF-8 character
            ch += os.read(fd, 1 if ch[0] < 0xE0 else 2 if ch[0] < 0xF0 else 3)
        return ch.decode("utf-8", "replace")
    finally:
        termios.tcsetattr(fd, termios.TCSADRAIN, old)


def prefilled_input(prompt, text=""):
    """input() with `text` already typed in, so Enter keeps it and the arrow keys edit it."""
    try:
        import readline
    except ImportError:
        readline = None
    if readline and text:
        readline.set_startup_hook(lambda: readline.insert_text(text))
    try:
        s = input(prompt)
    finally:
        if readline:
            readline.set_startup_hook()
    return text if not readline and not s else s


class TeamsError(Exception):
    pass


def load_config():
    try:
        cfg = json.loads(CONFIG_FILE.read_text())
    except (OSError, ValueError):
        cfg = {}
    cfg.setdefault("subjects", {})   # site url -> Moodle subfolder name
    cfg.setdefault("hidden", [])     # site urls hidden from `teams`
    return cfg


def save_config(cfg):
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    tmp = CONFIG_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(cfg, indent=2))
    os.chmod(tmp, 0o600)
    tmp.replace(CONFIG_FILE)


def moodle_root(cfg):
    """Same download folder moodle-cli uses (its `dir` setting), default ~/Moodle."""
    if cfg.get("download_dir"):
        return Path(os.path.expanduser(cfg["download_dir"]))
    try:
        d = json.loads(MOODLE_CFG.read_text()).get("download_dir")
        if d:
            return Path(os.path.expanduser(d))
    except (OSError, ValueError):
        pass
    return Path.home() / "Moodle"


# --------------------------------------------------------------------------- #
#  Browser session
# --------------------------------------------------------------------------- #

def interactive(fn):
    """Run something that reads from the keyboard. The full-screen view replaces this to leave curses first."""
    return fn()


LOGIN_HOSTS = ("login.microsoftonline.com", "login.microsoft.com", "login.live.com")


class Browser:
    """One Chrome profile, kept in DATA_DIR, so the Microsoft login cookie survives restarts."""

    def __init__(self, headless=True):
        self.headless = headless
        self.pw = self.ctx = None

    def start(self):
        if self.ctx:
            return self
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            raise TeamsError("Playwright is missing: run install.sh again")
        PROFILE_DIR.mkdir(parents=True, exist_ok=True)
        # report fetches made by service workers too (Teams loads its data through one)
        os.environ.setdefault("PW_EXPERIMENTAL_SERVICE_WORKER_NETWORK_EVENTS", "1")
        self.pw = sync_playwright().start()
        opts = dict(headless=self.headless, accept_downloads=True, viewport={"width": 1280, "height": 900})
        last = None
        for extra in ({"channel": "chrome"}, {"channel": "chromium"}, {}):
            try:
                self.ctx = self.pw.chromium.launch_persistent_context(str(PROFILE_DIR), **opts, **extra)
                break
            except Exception as e:  # browser not installed in this variant, try next
                last = e
        if not self.ctx:
            self.pw.stop()
            raise TeamsError(f"could not start Chrome/Chromium: {last}\n"
                             "  install one without root:  ~/teamsinterminal/.venv/bin/playwright install chromium")
        self.ctx.set_default_timeout(30000)
        return self

    def page(self):
        return self.ctx.pages[0] if self.ctx.pages else self.ctx.new_page()

    def new_page(self):
        return self.ctx.new_page()

    def close(self):
        try:
            if self.ctx:
                self.ctx.close()
            if self.pw:
                self.pw.stop()
        except Exception:
            pass
        self.ctx = self.pw = None


def on_login_page(pg):
    return urllib.parse.urlparse(pg.url).hostname in LOGIN_HOSTS


def _visible(pg, sel):
    loc = pg.locator(sel)
    try:
        return loc.first if loc.count() and loc.first.is_visible() else None
    except Exception:
        return None


def _text(pg, sel):
    el = _visible(pg, sel)
    try:
        return el.inner_text().strip() if el else ""
    except Exception:
        return ""


def _click_next(pg):
    for sel in ("#idSIButton9", "input[type=submit]", "button[type=submit]", "#idSubmit_SAOTCC_Continue"):
        el = _visible(pg, sel)
        if el:
            el.click()
            return True
    return False


def terminal_login(pg, start_url, done):
    """Fill the Microsoft sign-in page from the terminal. `done(pg)` says when we are through."""
    pg.goto(start_url, wait_until="domcontentloaded")
    shown, last_change, last_url = set(), time.time(), ""
    while True:
        pg.wait_for_timeout(700)
        if done(pg):
            return True
        if pg.url != last_url:
            last_url, last_change = pg.url, time.time()
        if not on_login_page(pg):
            # federated school login (ADFS etc.) or a page we don't know
            if time.time() - last_change > 40:
                raise TeamsError("sign-in stopped on an unknown page: " + pg.url[:120] +
                                 "\n  use `login --window` to sign in in a visible window")
            continue

        error = (_text(pg, "#usernameError") or _text(pg, "#passwordError") or _text(pg, "#idTD_Error")
                 or _text(pg, "[role=alert]") or _text(pg, "#field-error"))
        if error and error not in shown:
            shown.add(error)
            print(red("  " + error))

        number = _text(pg, "#idRichContext_DisplaySign") or _text(pg, "[data-testid=displaySign]")
        if number:
            key = "num" + number
            if key not in shown:
                shown.add(key)
                print(yellow(f"  Open Microsoft Authenticator and tap/enter the number: {bold(number)}"))
                print(dim("  waiting for approval…"))
            continue

        otc = _visible(pg, "input[name=otc]") or _visible(pg, "input[autocomplete=one-time-code]")
        if otc:
            hint = _text(pg, "#idDiv_SAOTCC_Description") or "Enter the code from your authenticator app / SMS"
            print(dim("  " + hint))
            otc.fill(input("  code: ").strip())
            _click_next(pg)
            pg.wait_for_timeout(1500)
            continue

        email = _visible(pg, "input[type=email]") or _visible(pg, "input[name=loginfmt]")
        if email and (not email.input_value() or (error and error in shown and "email" not in shown)):
            shown.discard("email")
            email.fill(input("  school e-mail: ").strip())
            if error:
                shown.add("email")   # one retry per error message
                shown.discard(error)
            _click_next(pg)
            pg.wait_for_timeout(1500)
            continue

        pw = _visible(pg, "input[type=password]")
        if pw:
            pw.fill(getpass.getpass("  password: "))
            _click_next(pg)
            pg.wait_for_timeout(2000)
            continue

        # "Stay signed in?" -> yes, so the next start needs no password
        if _visible(pg, "#KmsiCheckboxField") or _visible(pg, "#KmsiDescription") or "kmsi" in pg.url.lower():
            box = _visible(pg, "#KmsiCheckboxField")
            if box and not box.is_checked():
                box.check()
            _click_next(pg)
            continue

        # "Pick an account": take the first tile
        tile = _visible(pg, "#tilesHolder .table[role=button]") or _visible(pg, "[data-test-id] .table")
        if tile and "Pick an account" in (pg.title() + _text(pg, "#loginHeader")):
            tile.click()
            continue

        if time.time() - last_change > 60:
            raise TeamsError("sign-in is stuck (" + (_text(pg, "#loginHeader") or pg.title()) +
                             ")\n  use `login --window` to sign in in a visible window")


# --------------------------------------------------------------------------- #
#  SharePoint (where Teams keeps every team's files)
# --------------------------------------------------------------------------- #

class SharePoint:
    def __init__(self, browser, cfg):
        self.b, self.cfg = browser, cfg

    @property
    def host(self):
        return self.cfg.get("host")            # e.g. school.sharepoint.com

    @property
    def my_host(self):
        h = self.host
        return h.replace(".sharepoint.com", "-my.sharepoint.com") if h else None

    def _req(self, url, **kw):
        r = self.b.ctx.request.get(url, headers={"Accept": "application/json;odata=nometadata"},
                                   max_redirects=0, fail_on_status_code=False, **kw)
        return r

    def get(self, url, params=None, _retry=True):
        if params:
            url += ("&" if "?" in url else "?") + urllib.parse.urlencode(params, quote_via=urllib.parse.quote)
        r = self._req(url)
        if r.status in (401, 403, 302) and _retry:
            self.refresh_cookie(url)
            return self.get(url, _retry=False)
        if r.status != 200:
            try:
                msg = r.json().get("odata.error", {}).get("message", {}).get("value") or r.status_text
            except Exception:
                msg = r.status_text
            raise TeamsError(f"SharePoint {r.status}: {msg}")
        return r.json()

    def refresh_cookie(self, url):
        """SharePoint cookies are short-lived; opening the site re-issues them via the saved login."""
        site = re.match(r"https://[^/]+(/(sites|teams)/[^/]+)?", url)[0]
        pg = self.b.page()
        pg.goto(site, wait_until="domcontentloaded")
        for _ in range(40):
            if urllib.parse.urlparse(pg.url).hostname.endswith("sharepoint.com"):
                return
            if on_login_page(pg) and (_visible(pg, "input[type=password]") or _visible(pg, "input[type=email]")):
                def again():
                    print(yellow("Your Microsoft session expired, please sign in again:"))
                    terminal_login(pg, site, lambda p: urllib.parse.urlparse(p.url).hostname.endswith("sharepoint.com"))
                interactive(again)
                return
            pg.wait_for_timeout(500)

    def me(self):
        return self.get(f"https://{self.my_host}/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
                        {"$select": "DisplayName,Email,AccountName"})

    def team_sites(self):
        """All Microsoft 365 group sites (= teams) and channel sites you can open."""
        q = "contentclass:STS_Site AND (WebTemplate:GROUP OR WebTemplate:TEAMCHANNEL)"
        out, start = [], 0
        while True:
            d = self.get(f"https://{self.my_host}/_api/search/query", {
                "querytext": f"'{q}'", "selectproperties": "'Title,Path,WebTemplate,Description,LastModifiedTime'",
                "rowlimit": "500", "startrow": str(start), "trimduplicates": "false"})
            table = d["PrimaryQueryResult"]["RelevantResults"]
            rows = table["Table"]["Rows"]
            for row in rows:
                c = {x["Key"]: x["Value"] for x in row["Cells"]}
                if c.get("Path"):
                    out.append({"title": c.get("Title") or c["Path"].rsplit("/", 1)[-1], "url": c["Path"].rstrip("/"),
                                "channel": c.get("WebTemplate") == "TEAMCHANNEL",
                                "desc": c.get("Description") or "", "modified": c.get("LastModifiedTime")})
            start += len(rows)
            if not rows or start >= int(table.get("TotalRows") or 0):
                break
        seen, uniq = set(), []
        for s in sorted(out, key=lambda s: s["title"].lower()):
            if s["url"].lower() not in seen:
                seen.add(s["url"].lower())
                uniq.append(s)
        return uniq

    def libraries(self, site):
        d = self.get(f"{site}/_api/web/lists", {
            "$filter": "BaseTemplate eq 101 and Hidden eq false", "$expand": "RootFolder",
            "$select": "Title,ItemCount,RootFolder/ServerRelativeUrl"})
        libs = [{"title": x["Title"], "path": x["RootFolder"]["ServerRelativeUrl"], "count": x.get("ItemCount", 0)}
                for x in d.get("value", [])]
        # Teams uses "Shared Documents" (title is translated, e.g. "Dokumente"); put it first
        libs.sort(key=lambda l: (not l["path"].endswith("/Shared Documents"), l["title"].lower()))
        return libs

    def folder(self, site, path):
        enc = path.replace("'", "''")
        base = f"{site}/_api/web/GetFolderByServerRelativePath(decodedurl=@p)"
        p = {"@p": f"'{enc}'"}
        folders = self.get(base + "/Folders", {**p, "$select": "Name,ServerRelativeUrl,ItemCount,TimeLastModified"})
        files = self.get(base + "/Files", {**p, "$select": "Name,ServerRelativeUrl,Length,TimeLastModified"})
        items = [{"kind": "dir", "name": f["Name"], "path": f["ServerRelativeUrl"], "count": f.get("ItemCount", 0),
                  "modified": f.get("TimeLastModified")}
                 for f in folders.get("value", []) if f["Name"] != "Forms"]
        items.sort(key=lambda x: x["name"].lower())
        fl = [{"kind": "file", "name": f["Name"], "path": f["ServerRelativeUrl"], "size": int(f.get("Length") or 0),
               "modified": f.get("TimeLastModified")} for f in files.get("value", [])]
        fl.sort(key=lambda x: x["name"].lower())
        return items + fl

    def walk(self, site, path):
        for it in self.folder(site, path):
            if it["kind"] == "dir":
                yield from ((it["name"] + "/" + rel, f) for rel, f in self.walk(site, it["path"]))
            else:
                yield it["name"], it

    def download(self, site, path):
        url = f"{site}/_api/web/GetFileByServerRelativePath(decodedurl=@p)/$value?@p=" + \
              urllib.parse.quote("'" + path.replace("'", "''") + "'")
        r = self._req(url)
        if r.status in (401, 403, 302):
            self.refresh_cookie(site)
            r = self._req(url)
        if r.status != 200:
            raise TeamsError(f"download failed ({r.status}): {path}")
        return r.body()

    def web_url(self, site, path):
        return f"https://{self.host}" + urllib.parse.quote(path)


# --------------------------------------------------------------------------- #
#  Microsoft Forms in the terminal (drives the real page)
# --------------------------------------------------------------------------- #

FORMS_SCAN = r"""
() => {
  const vis = e => !!(e && (e.offsetWidth || e.offsetHeight || e.getClientRects().length));
  const txt = e => (e ? (e.innerText || e.textContent || '') : '').replace(/\s+\n/g, '\n').trim();
  let qs = [...document.querySelectorAll('[data-automation-id="questionItem"]')].filter(vis);
  if (!qs.length) qs = [...document.querySelectorAll('[role="group"][aria-labelledby], [role="radiogroup"]')].filter(vis);
  document.querySelectorAll('[data-tcli]').forEach(e => e.removeAttribute('data-tcli'));
  return qs.map((q, i) => {
    q.setAttribute('data-tcli', 'q' + i);
    const titleEl = q.querySelector('[data-automation-id="questionTitle"]') ||
                    document.getElementById((q.getAttribute('aria-labelledby') || '').split(' ')[0]);
    let title = txt(titleEl) || txt(q).split('\n')[0];
    const required = !!q.querySelector('[aria-required="true"], [required]') || /\*\s*$/.test(title.split('\n')[0]);
    const imgs = q.querySelectorAll('img').length;
    const labelOf = (el) => {
      const item = el.closest('[data-automation-id="choiceItem"]') || el.closest('label') ||
                   (el.id && document.querySelector(`label[for="${el.id}"]`));
      return (el.getAttribute('aria-label') || txt(item) || el.value || '').trim();
    };
    let opts = [...q.querySelectorAll('input[type=radio]')];
    let type = 'radio';
    if (!opts.length) { opts = [...q.querySelectorAll('input[type=checkbox]')]; type = 'checkbox'; }
    if (!opts.length) { opts = [...q.querySelectorAll('[role=radio]')]; type = 'radio'; }
    if (!opts.length) { opts = [...q.querySelectorAll('[role=checkbox]')]; type = 'checkbox'; }
    const choices = opts.map((o, j) => {
      o.setAttribute('data-tcli', `q${i}o${j}`);
      return {label: labelOf(o), checked: o.checked === true || o.getAttribute('aria-checked') === 'true'};
    });
    const texts = [...q.querySelectorAll('textarea, input[type=text], input[type=number], input[type=date], input[type=time], input:not([type])')]
                  .filter(vis);
    texts.forEach((t, j) => t.setAttribute('data-tcli', `q${i}t${j}`));
    const drop = q.querySelector('[role=combobox], [aria-haspopup=listbox], select');
    if (drop) drop.setAttribute('data-tcli', `q${i}d`);
    const upload = !!q.querySelector('input[type=file], [data-automation-id*="upload" i]');
    if (choices.length) {
      // an "Other" option has its own text box
      return {i, title, required, imgs, type, choices, other: texts.length ? `q${i}t0` : null,
              otherValue: texts.length ? texts[0].value : ''};
    }
    if (texts.length) return {i, title, required, imgs, type: texts[0].tagName === 'TEXTAREA' ? 'long' : (texts[0].type || 'text'),
                              field: `q${i}t0`, value: texts[0].value};
    if (drop) return {i, title, required, imgs, type: 'dropdown', field: `q${i}d`, value: txt(drop)};
    if (upload) return {i, title, required, imgs, type: 'upload'};
    return {i, title, required, imgs, type: 'unknown'};
  });
}
"""

FORMS_BUTTONS = r"""
() => {
  const vis = e => !!(e && (e.offsetWidth || e.offsetHeight));
  const out = {};
  const pick = (k, ids, re) => {
    let b = ids.map(id => document.querySelector(`[data-automation-id="${id}"]`)).find(vis);
    if (!b) b = [...document.querySelectorAll('button, [role=button], input[type=submit]')]
                 .find(x => vis(x) && re.test((x.innerText || x.value || x.getAttribute('aria-label') || '').trim()));
    if (b) { b.setAttribute('data-tcli-btn', k); out[k] = (b.innerText || b.value || k).trim(); }
  };
  pick('submit', ['submitButton'], /^(submit|absenden|senden|einreichen)$/i);
  pick('next', ['nextButton'], /^(next|weiter|nächste)$/i);
  pick('back', ['backButton', 'previousButton'], /^(back|previous|zurück)$/i);
  pick('start', ['startButton', 'startNowButton'], /^(start( now)?|jetzt starten|starten|beginnen|begin)$/i);
  return out;
}
"""

FORMS_ERRORS = r"""
() => [...document.querySelectorAll('[data-automation-id="validationError"], [role=alert]')]
        .map(e => (e.innerText || '').trim()).filter(Boolean)
"""


FORM_CLEAR = "\033[H\033[2J\033[3J"
FORM_NAV = {"left": "back", "b": "back", "backspace": "back", "r": "review", "o": "open", "q": "quit", "esc": "quit"}


def _wrap(text, indent="   "):
    import textwrap
    return "\n".join(textwrap.fill(line, width() - len(indent), initial_indent=indent, subsequent_indent=indent)
                     if line.strip() else "" for line in text.split("\n"))


class FormRunner:
    def __init__(self, browser, url):
        self.b, self.url = browser, url
        self.pg = None
        self.title, self.desc = "Microsoft Form", ""
        self.flash = []           # messages shown on top of the next screen
        self.page_no = 1

    def q(self, tag):
        return self.pg.locator(f'[data-tcli="{tag}"]')

    def tick(self, tag):
        """Click a choice like a user would; Forms often hides the real <input> under its label."""
        el = self.q(tag).first
        try:
            el.click(force=True, timeout=3000)
        except Exception:
            el.evaluate("e => (e.closest('label') || e).click()")

    def open(self):
        self.pg = self.b.new_page()
        self.pg.goto(self.url, wait_until="domcontentloaded")
        for _ in range(60):
            self.pg.wait_for_timeout(500)
            if on_login_page(self.pg) and (_visible(self.pg, "input[type=password]") or _visible(self.pg, "input[type=email]")):
                print(yellow("Forms wants you to sign in:"))
                terminal_login(self.pg, self.url, lambda p: "forms." in (urllib.parse.urlparse(p.url).hostname or ""))
            if self.scan() or self.buttons().get("start"):
                return
        raise TeamsError("the form did not load (closed, no access, or not a Microsoft Forms link)")

    def scan(self):
        try:
            return self.pg.evaluate(FORMS_SCAN)
        except Exception:
            return []

    def buttons(self):
        try:
            return self.pg.evaluate(FORMS_BUTTONS)
        except Exception:
            return {}

    def click_btn(self, key):
        self.pg.locator(f'[data-tcli-btn="{key}"]').first.click()
        self.pg.wait_for_timeout(1500)

    def header(self):
        self.title, self.desc = self.pg.evaluate("""() => {
          const t = document.querySelector('[data-automation-id="formTitle"]') || document.querySelector('h1');
          const d = document.querySelector('[data-automation-id="formDescription"]');
          return [(t && t.innerText || document.title || '').trim(), (d && d.innerText || '').trim()];
        }""")
        self.title = self.title or "Microsoft Form"

    # -- screen ------------------------------------------------------------- #
    @staticmethod
    def answered(q):
        return any(c["checked"] for c in q.get("choices") or []) or bool(q.get("value"))

    @staticmethod
    def answer_text(q):
        if q.get("choices"):
            return ", ".join(c["label"] for c in q["choices"] if c["checked"])
        return q.get("value") or ""

    def frame(self, body, keys, qs=None, n=None):
        """Clear the terminal and draw: title, progress bar, body, key line."""
        W = min(width(), 92)
        right = f"page {self.page_no}" + (f" · question {n} of {len(qs)}" if n and qs else "")
        title = self.title[:max(10, W - len(right) - 5)]
        out = [FORM_CLEAR, " " + bold(magenta(title)) + " " * max(2, W - len(title) - len(right) - 2) + dim(right)]
        if qs:
            done = sum(1 for q in qs if self.answered(q))
            fill = round((W - 2) * done / len(qs))
            out.append(" " + green("━" * fill) + dim("━" * (W - 2 - fill)))
        else:
            out.append(" " + dim("─" * (W - 2)))
        out.append("")
        out += ["  " + red("! " + f) for f in self.flash]
        self.flash = []
        out += body
        out += ["", " " + "   ".join(bold(cyan(k)) + " " + dim(label) for k, label in keys)]
        sys.stdout.write("\n".join(out) + "\n")
        sys.stdout.flush()

    def q_body(self, n, q, cur=None):
        import textwrap
        W = min(width(), 92)
        lines = q["title"].splitlines() or [""]
        first = re.sub(r"^\d+[.)]\s*", "", lines[0]).rstrip(" *")
        req = red(" *") if q["required"] else ""
        wrapped = textwrap.wrap(first, W - 8) or [""]
        out = [f"  {cyan(bold(f'{n:>2}'))}  {bold(wrapped[0])}" + (req if len(wrapped) == 1 else "")]
        out += ["      " + bold(w) for w in wrapped[1:]]
        if len(wrapped) > 1:
            out[-1] += req
        rest = "\n".join(lines[1:]).strip()
        if rest:
            out += ["      " + w for line in rest.split("\n") for w in (textwrap.wrap(line, W - 8) or [""])]
        if q["imgs"]:
            out.append("      " + yellow(f"▣ {q['imgs']} image(s) in this question: press o to see them in your browser"))
        out.append("")
        multi = q["type"] == "checkbox"
        for j, c in enumerate(q.get("choices") or []):
            mark = (green("■") if c["checked"] else "□") if multi else (green("◉") if c["checked"] else "○")
            here = j == cur
            label = c["label"] or dim("(no text)")
            out.append(f"     {magenta('❯') if here else ' '} {mark} {dim(f'{j + 1}.')} "
                       + (bold(label) if here else green(label) if c["checked"] else label))
        return out

    # -- one question ------------------------------------------------------- #
    def refresh(self, q):
        return next((x for x in self.scan() if x["i"] == q["i"]), q)

    def answer_q(self, n, q, qs=None):
        """Returns 'next', 'back', 'review', 'open' or 'quit'."""
        qs = qs or [q]
        t = q["type"]
        if t in ("radio", "checkbox"):
            return self.ask_choice(n, q, qs, self.pick_choice)
        if t in ("text", "long", "number", "date", "time"):
            return self.ask_text(n, q, qs)
        if t == "dropdown":
            return self.ask_dropdown(n, q, qs)
        kind = "a file upload" if t == "upload" else "a question type (rating, ranking, Likert…)"
        body = self.q_body(n, q) + ["  " + yellow(f"This is {kind} the terminal cannot fill."),
                                    "  " + dim("Nothing has been submitted yet. Press o to finish the form in your browser.")]
        self.frame(body, [("o", "browser"), ("→", "skip"), ("←", "back"), ("r", "review"), ("q", "quit")], qs, n)
        k = read_key()
        return FORM_NAV.get(k) or "next"

    def ask_choice(self, n, q, qs, apply):
        multi = q["type"] == "checkbox"
        cur = next((j for j, c in enumerate(q["choices"]) if c["checked"]), 0)
        keys = [("↑↓", "move"), ("space", "tick" if multi else "choose"),
                ("enter", "next" if multi else "choose + next"), ("→", "skip"), ("←", "back"),
                ("r", "review"), ("o", "browser"), ("q", "quit")]
        while True:
            qs = [q if x["i"] == q["i"] else x for x in qs]
            self.frame(self.q_body(n, q, cur), keys, qs, n)
            k = read_key()
            nc = len(q["choices"])
            if k in ("up", "k"):
                cur = (cur - 1) % nc
            elif k in ("down", "j"):
                cur = (cur + 1) % nc
            elif k in ("home", "g"):
                cur = 0
            elif k in ("end", "G"):
                cur = nc - 1
            elif k.isdigit() and 1 <= int(k) <= nc:
                cur = int(k) - 1
                q = apply(q, cur)
            elif k == "space":
                q = apply(q, cur)
            elif k == "enter":
                if not multi and not q["choices"][cur]["checked"]:
                    q = apply(q, cur)
                return "next"
            elif k in ("right", "tab"):
                return "next"
            elif k in FORM_NAV:
                return FORM_NAV[k]

    def pick_choice(self, q, j):
        self.tick(f"q{q['i']}o{j}")
        self.pg.wait_for_timeout(150)
        q2 = self.refresh(q)
        c = q2["choices"][j] if j < len(q2.get("choices") or []) else None
        if c and c["checked"] and q2.get("other") and re.match(r"^(other|sonstige|andere)", c["label"], re.I):
            s = prefilled_input("     other: ", q2.get("otherValue") or "")
            self.q(q2["other"]).first.fill(s)
            q2 = self.refresh(q2)
        return q2

    def ask_text(self, n, q, qs):
        t = q["type"]
        hint = {"date": "as the form shows it, e.g. 24.11.2026", "time": "e.g. 14:30",
                "number": "a number"}.get(t, "")
        keys = [("enter", "save + next"), (":b", "back"), (":r", "review"), (":o", "browser"), (":q", "quit")]
        if t == "long":
            keys.insert(1, (":e", "write in $EDITOR"))
        body = self.q_body(n, q)
        if hint:
            body.append("      " + dim(hint))
        self.frame(body, keys, qs, n)
        old = q.get("value") or ""
        s = prefilled_input("     › ", old)
        cmd = s.strip().lower()
        if cmd in (":b", ":r", ":o", ":q"):
            return {":b": "back", ":r": "review", ":o": "open", ":q": "quit"}[cmd]
        if t == "long" and cmd == ":e":
            s = edit_text(old)
        if s != old:
            self.q(q["field"]).first.fill(s)
        return "next"

    def ask_dropdown(self, n, q, qs):
        field = self.q(q["field"]).first
        field.click()
        self.pg.wait_for_timeout(500)
        opts = self.pg.locator("[role=option]")
        labels = [opts.nth(k).inner_text().strip() for k in range(opts.count())]
        self.pg.keyboard.press("Escape")
        if not labels:
            self.flash.append("this drop-down has no options")
            return "next"
        fake = dict(q, type="radio", choices=[{"label": l, "checked": l == q.get("value")} for l in labels])

        def apply(fq, j):
            field.click()
            self.pg.wait_for_timeout(400)
            self.pg.locator("[role=option]").nth(j).click()
            self.pg.wait_for_timeout(200)
            return dict(fq, value=labels[j], choices=[dict(c, checked=k == j) for k, c in enumerate(fq["choices"])])
        return self.ask_choice(n, fake, qs, apply)

    # -- whole form --------------------------------------------------------- #
    def leave(self):
        return confirm("  Leave without submitting? Your answers are not sent.", True)

    def run(self):
        if not sys.stdin.isatty():
            raise TeamsError("Forms tests need a terminal")
        print(dim("opening the form…"))
        self.open()
        self.header()
        btn = self.buttons()
        if btn.get("start") and not self.scan():
            body = self.pg.inner_text("body")
            m = re.search(r"(\d+\s*(min|minute|minuten|hour|stunde)\w*)", body, re.I)
            desc = [_wrap(self.desc, "  ")] if self.desc else []
            self.frame(desc + ["", "  " + yellow("⏱  This form starts a timer when you begin" + (f" ({m[1]})." if m else ".")),
                               "  " + dim("Answers are only sent when you submit at the end.")],
                       [("enter", "start now"), ("q", "cancel")])
            if read_key() != "enter":
                return
            self.click_btn("start")
            self.pg.wait_for_timeout(1500)
        self.page_no = 0
        while True:
            qs = self.scan()
            self.page_no += 1
            if not qs:
                raise TeamsError("no questions found on this page (type of form not supported)")
            k, went_back = 0, False
            while k < len(qs):
                act = self.answer_q(k + 1, qs[k], qs)
                qs = self.scan() or qs
                if act == "quit":
                    if self.leave():
                        return
                    continue
                if act == "open":
                    return self.open_in_browser()
                if act == "back":
                    if k == 0 and self.page_no > 1 and self.buttons().get("back"):
                        self.click_btn("back")
                        self.page_no -= 2
                        went_back = True
                        break
                    k = max(0, k - 1)
                elif act == "review":
                    break
                else:
                    k += 1
            if went_back:
                continue
            btn = self.buttons()
            if btn.get("next") and not btn.get("submit"):
                first = qs[0]["title"]
                self.click_btn("next")
                errs = self.pg.evaluate(FORMS_ERRORS)
                now = self.scan()
                if errs and now and now[0]["title"] == first:
                    self.flash += errs
                    self.page_no -= 1
                continue
            if not self.review_and_submit():
                self.page_no -= 1
                continue
            return

    def review_and_submit(self):
        import textwrap
        while True:
            qs = self.scan()
            W = min(width(), 92)
            body, missing = ["  " + bold("Check your answers"), ""], []
            for n, q in enumerate(qs, 1):
                ans = self.answer_text(q)
                if not ans and q["required"]:
                    missing.append(n)
                mark = green("✓") if ans else red("✗") if q["required"] else dim("·")
                first = re.sub(r"^\d+[.)]\s*", "", q["title"].splitlines()[0]).rstrip(" *")
                body.append(f"  {mark} {dim(f'{n:>2}')}  {first[:W - 10]}")
                shown = textwrap.shorten(ans, W - 10, placeholder="…") if ans else ""
                body.append("        " + (green(shown) if ans else red("no answer, required") if q["required"]
                                           else dim("no answer")))
            if missing:
                body += ["", "  " + yellow(f"{len(missing)} required question(s) without an answer: "
                                           + ", ".join(map(str, missing)))]
            self.frame(body, [("s", "submit"), ("1-" + str(len(qs)), "change an answer"), ("o", "browser"),
                              ("q", "quit")], qs)
            s = input("  › ").strip().lower()
            if s == "q":
                if self.leave():
                    return True
                continue
            if s == "o":
                self.open_in_browser()
                return True
            if s.isdigit() and 1 <= int(s) <= len(qs):
                self.answer_q(int(s), qs[int(s) - 1], qs)
                continue
            if s != "s":
                continue
            if not self.buttons().get("submit"):
                raise TeamsError("no Submit button on this page")
            if missing and not confirm(yellow(f"  {len(missing)} required answer(s) missing. Submit anyway?")):
                continue
            if not confirm(bold("  Submit the form now? This cannot be undone.")):
                continue
            break
        self.click_btn("submit")
        self.pg.wait_for_timeout(2500)
        errs = self.pg.evaluate(FORMS_ERRORS)
        if errs and self.buttons().get("submit"):
            self.flash += errs + ["Not submitted. Fix the answers above."]
            return False
        self.pg.wait_for_timeout(1500)
        body = self.pg.inner_text("body").strip()
        lines = [l.strip() for l in body.splitlines() if l.strip()]
        score = re.search(r"(\d+(?:[.,]\d+)?)\s*/\s*(\d+(?:[.,]\d+)?)\s*(?:points?|punkte?|pkt|pts)", body, re.I) or \
            re.search(r"(?:points?|punkte?)\s*:?\s*(\d+(?:[.,]\d+)?)\s*/\s*(\d+(?:[.,]\d+)?)", body, re.I)
        out = ["  " + green(bold("✓ Submitted")), ""]
        if score:
            out += ["  " + bold(f"Score  {score[1]} / {score[2]}"), ""]
        out += [_wrap("\n".join(lines[:12]), "  ")]
        self.frame(out, [])
        return True

    def open_in_browser(self):
        print(dim("Opening in your browser. Answers typed here were not submitted and are not carried over."))
        subprocess.Popen(["xdg-open", self.url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    def close(self):
        try:
            if self.pg:
                self.pg.close()
        except Exception:
            pass


# --------------------------------------------------------------------------- #
#  Teams web app: channels, posts, assignments and the Forms links in them
# --------------------------------------------------------------------------- #

FORMS_RE = re.compile(r"https?://(?:forms\.office\.com|forms\.cloud\.microsoft|forms\.microsoft\.com)/[^\s\"'<>\\)\]]+", re.I)
GUID_RE = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", re.I)
TEAMS_HOSTS = ("teams.microsoft.com", "teams.cloud.microsoft", "teams.live.com", "skype.com",
               "onenote.com", "office.com", "office.net", "microsoft.com")
WEB_BUTTONS = re.compile(r"^(continue on this browser|use the web app instead|use web app|"
                         r"im browser fortfahren|stattdessen die web-app verwenden|web-app verwenden)$", re.I)


def forms_links(text):
    """Microsoft Forms answer links in any text/JSON (design and analysis links are left out)."""
    t = str(text).replace("\\u0026", "&").replace("\\/", "/").replace("&amp;", "&")
    out = []
    for m in FORMS_RE.finditer(t):
        u = m.group(0).rstrip(".,;:")
        if re.search(r"/(r|e)/|ResponsePage|[?&]id=", u, re.I) and not re.search(r"DesignPage|AnalysisPage", u, re.I):
            if u not in out:
                out.append(u)
    return out


def walk_dicts(o):
    stack = [o]
    while stack:
        x = stack.pop()
        if isinstance(x, dict):
            yield x
            stack.extend(x.values())
        elif isinstance(x, list):
            stack.extend(x)


def find_key(o, pattern):
    rx = re.compile(pattern, re.I)
    for d in walk_dicts(o):
        for k, v in d.items():
            if rx.fullmatch(k) and isinstance(v, str) and v:
                return v
    return None


class _H2T(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.out, self.links, self.href, self.skip = [], [], None, 0

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag in ("br", "p", "div", "li", "tr", "h1", "h2", "h3"):
            self.out.append("\n")
        if tag == "li":
            self.out.append("• ")
        if tag == "a":
            self.href = a.get("href")
        if tag == "img":
            self.out.append("[image]")
        if tag in ("script", "style"):
            self.skip += 1

    def handle_endtag(self, tag):
        if tag == "a" and self.href:
            if self.href.startswith("http"):
                self.links.append(self.href)
                self.out.append(f" [{len(self.links)}]")
            self.href = None
        if tag in ("script", "style"):
            self.skip -= 1

    def handle_data(self, data):
        if not self.skip:
            self.out.append(data)


def html2text(html):
    p = _H2T()
    try:
        p.feed(str(html or ""))
    except Exception:
        return re.sub(r"<[^>]+>", "", str(html or "")), []
    text = re.sub(r"[ \t]+", " ", "".join(p.out))
    text = re.sub(r"\n\s*\n\s*\n+", "\n\n", text).strip()
    return text, p.links


class TeamsWeb:
    """Opens teams.microsoft.com in the hidden browser and reads the JSON the page downloads to show itself."""
    ASSIGNMENTS_APP = "66aeee93-507d-479a-a3ef-8f494af43945"

    def __init__(self, browser):
        self.b = browser
        self.pg = None
        self.pending, self.caps = [], []
        self.hooked = False
        self.ready = False
        self.error = None
        self.forms = {}          # url -> {"label", "where"}
        self.auth = None         # (csa base url, headers) copied from a request the page made
        self.team_list = None    # the teams you are a member of, as Teams itself lists them

    def _on_request(self, r):
        if self.auth or "/api/csa/" not in r.url:
            return
        try:
            h = r.headers
        except Exception:
            return
        if "authorization" in h:
            base = r.url.split("/api/csa/")[0] + "/api/csa/" + r.url.split("/api/csa/")[1].split("/")[0]
            keep = ("authorization", "x-ms-client-version", "x-ms-region", "x-ringoverride")
            self.auth = (base, {k: v for k, v in h.items() if k in keep})

    def fetch_team_list(self):
        """Ask Teams for the full list of your teams (archived flag, class type, SharePoint site, channels)."""
        if not self.auth:
            return None
        base, h = self.auth
        r = self.b.ctx.request.get(base + "/api/v3/teams/users/me?isPrefetch=false&enableMembershipSummary=true",
                                   headers=h, fail_on_status_code=False)
        if r.status != 200:
            if r.status == 401:
                self.auth = None          # token expired; the page makes a new request soon
            return None
        text = r.text()
        data = json.loads(text)
        self.caps.append({"url": base + "/api/v3/teams/users/me", "host": urllib.parse.urlparse(base).hostname,
                          "data": data, "text": text})
        self.team_list = data.get("teams") or []
        return self.team_list

    def joined(self):
        out = []
        for t in self.team_list or []:
            if t.get("isDeleted") or not isinstance(t, dict):
                continue
            tsi = t.get("teamSiteInformation") if isinstance(t.get("teamSiteInformation"), dict) else {}
            out.append({"name": t.get("displayName") or "", "id": t.get("id"), "groupId": tsi.get("groupId"),
                        "site": (tsi.get("sharepointSiteUrl") or "").rstrip("/"), "archived": bool(t.get("isArchived")),
                        "type": t.get("teamType"), "desc": t.get("description") or ""})
        return out

    def reload_list(self):
        if not self.ready:
            return self.boot()
        for _ in range(30):
            if self.fetch_team_list() is not None:
                return
            self.wait(self.pg, 1)
        raise TeamsError("Teams did not answer; try again in a moment")

    # -- capture ------------------------------------------------------------ #
    def _on_response(self, r):
        self.pending.append(r)

    def drain(self):
        rs, self.pending = self.pending, []
        for r in rs:
            try:
                host = urllib.parse.urlparse(r.url).hostname or ""
                if not host.endswith(TEAMS_HOSTS) or "json" not in (r.headers.get("content-type") or ""):
                    continue
                body = r.text()
                data = json.loads(body)
            except Exception:
                continue
            self.caps.append({"url": r.url.split("?")[0], "host": host, "data": data, "text": body})

    def wait(self, pg, seconds, until=None, stable=3):
        """Wait while the page loads; stop when until() has returned the same thing for `stable` seconds."""
        end, last, same = time.time() + seconds, None, 0
        while time.time() < end:
            pg.wait_for_timeout(1000)
            self.drain()
            self.click_web_button(pg)
            if on_login_page(pg) and (_visible(pg, "input[type=password]") or _visible(pg, "input[type=email]")):
                url = pg.url
                interactive(lambda: (print(yellow("Teams wants you to sign in:")),
                                     terminal_login(pg, url, lambda p: (urllib.parse.urlparse(p.url).hostname or "")
                                                    .endswith(("teams.microsoft.com", "teams.cloud.microsoft")))))
            if until:
                v = until()
                if v:
                    same = same + 1 if v == last else 0
                    last = v
                    if same >= stable:
                        return v
        return last

    def click_web_button(self, pg):
        """Deep links show 'Open the Teams app?' first; take the web version."""
        try:
            for b in pg.get_by_role("button").all()[:30] + pg.get_by_role("link").all()[:30]:
                t = (b.inner_text(timeout=500) or "").strip()
                if WEB_BUTTONS.match(t):
                    b.click(timeout=2000)
                    return True
        except Exception:
            pass
        return False

    # -- boot --------------------------------------------------------------- #
    def boot(self):
        if self.ready:
            return
        if self.error:
            raise TeamsError(self.error)
        self.b.start()
        if not self.hooked:
            self.b.ctx.on("response", self._on_response)
            self.b.ctx.on("request", self._on_request)
            self.hooked = True
        self.pg = self.b.new_page()
        try:
            self.pg.goto("https://teams.microsoft.com/v2/", wait_until="commit", timeout=60000)
        except Exception:
            pass
        end = time.time() + 90
        while time.time() < end and not self.team_list:
            self.wait(self.pg, 1)
            try:
                self.fetch_team_list()
            except Exception:
                pass
        if not self.team_list:     # older way: wait until the page has loaded its own copy
            self.wait(self.pg, max(5, end - time.time()), lambda: len(self.teams()), stable=4)
        if self.teams():
            self.ready = True
            return
        self.error = ("the Teams web app did not load in the hidden browser, so posts and assignments are not "
                      "available (files still work). Run `teams doctor` for details.")
        raise TeamsError(self.error)

    def teams(self):
        out = {}
        for c in self.caps:
            for d in walk_dicts(c["data"]):
                chans = d.get("channels")
                if not (isinstance(chans, list) and chans and isinstance(d.get("displayName"), str)):
                    continue
                cl = [{"id": x["id"], "name": x.get("displayName") or "General"} for x in chans
                      if isinstance(x, dict) and "@thread" in str(x.get("id", ""))]
                if not cl:
                    continue
                gid = find_key(d, r"(aad)?groupid|teamsitegroupid")
                tid = d.get("id") or d.get("threadId") or cl[0]["id"]
                cl.sort(key=lambda x: (x["id"] != tid, x["name"].lower()))   # "General" first
                out[tid] = {"id": tid, "name": d["displayName"], "groupId": gid if gid and GUID_RE.match(gid) else None,
                            "channels": cl, "raw": d}
        return out

    def match(self, title):
        ts = list(self.teams().values())
        if not ts:
            return None
        n = norm(title)
        exact = [t for t in ts if norm(t["name"]) == n]
        if exact:
            return exact[0]
        best = max(ts, key=lambda t: difflib.SequenceMatcher(None, n, norm(t["name"])).ratio())
        return best if difflib.SequenceMatcher(None, n, norm(best["name"])).ratio() > 0.75 else None

    # -- posts -------------------------------------------------------------- #
    def messages(self, conv_id, since=0):
        cid = conv_id.lower()
        msgs = {}
        for c in self.caps[since:] if since else self.caps:
            for d in walk_dicts(c["data"]):
                if not isinstance(d.get("content"), str):
                    continue
                when = d.get("composetime") or d.get("originalarrivaltime") or d.get("createdDateTime")
                conv = str(d.get("conversationid") or d.get("conversationLink") or d.get("conversationId") or "")
                if not when or cid not in conv.lower():
                    continue
                mtype = str(d.get("messagetype") or d.get("messageType") or "")
                if mtype.startswith(("ThreadActivity", "Control", "Event")):
                    continue
                author = d.get("imdisplayname") or d.get("fromDisplayNameInToken") or \
                    (find_key(d.get("from") or {}, "displayname") if isinstance(d.get("from"), dict) else None) or ""
                m = re.search(r";messageid=(\d+)", conv)
                mid = str(d.get("id") or d.get("clientmessageid") or when)
                props = d.get("properties") if isinstance(d.get("properties"), dict) else {}
                text, links = html2text(d["content"])
                subject = props.get("subject") or d.get("subject") or ""
                msgs[mid] = {"id": mid, "root": m[1] if m and m[1] != mid else None, "author": author,
                             "time": when, "subject": subject, "text": text, "links": links,
                             "forms": forms_links(json.dumps(d, ensure_ascii=False)),
                             "deleted": bool(props.get("deletetime"))}
        roots = sorted((m for m in msgs.values() if not m["root"] and not m["deleted"]), key=lambda m: m["time"],
                       reverse=True)
        for r in roots:
            r["replies"] = sorted((m for m in msgs.values() if m["root"] == r["id"] and not m["deleted"]),
                                  key=lambda m: m["time"])
        return roots

    def channel(self, team, ch):
        self.boot()
        url = (f"https://teams.microsoft.com/l/channel/{urllib.parse.quote(ch['id'], safe='')}/"
               f"{urllib.parse.quote(ch['name'], safe='')}" + (f"?groupId={team['groupId']}" if team.get("groupId") else ""))
        self.pg.goto(url, wait_until="commit", timeout=60000)
        self.wait(self.pg, 40, lambda: len(self.messages(ch["id"])), stable=3)
        posts = self.messages(ch["id"])
        where = f"{team['name']} › {ch['name']}"
        for p in posts:
            for m in [p] + p["replies"]:
                for u in m["forms"]:
                    self.forms.setdefault(u, {"label": (m["subject"] or m["text"].split("\n")[0])[:70], "where": where,
                                              "time": m["time"]})
        # Forms tabs at the top of the channel
        for c in self.caps:
            if ch["id"].lower() in c["text"].lower():
                for d in walk_dicts(c["data"]):
                    name = d.get("displayName") or d.get("name")
                    if isinstance(name, str):
                        for u in forms_links(json.dumps({k: v for k, v in d.items() if isinstance(v, (str, dict))})):
                            self.forms.setdefault(u, {"label": f"tab: {name}", "where": where, "time": ""})
        return posts

    def team_forms(self, team):
        return {u: f for u, f in self.forms.items() if f["where"].startswith(team["name"] + " ›")}

    # -- assignments -------------------------------------------------------- #
    EDU_API = "https://assignments.edu.cloud.microsoft/api/v1.0/edu/me/work"

    def _on_edu_request(self, r):
        if "/edu/me/" in r.url:
            try:
                h = r.headers
            except Exception:
                return
            if "authorization" in h:
                self.edu_auth = {k: v for k, v in h.items() if k == "authorization" or k.startswith("x-")}

    def edu_token(self):
        """The Assignments app's sign-in header: open the app once from the Teams sidebar and copy it."""
        if getattr(self, "edu_auth", None):
            return self.edu_auth
        self.edu_auth = None
        if not getattr(self, "edu_hooked", False):
            self.b.ctx.on("request", self._on_edu_request)
            self.edu_hooked = True
        btn = self.pg.locator('[aria-label*="Assignments"], [aria-label*="Aufgaben"], '
                              'button:has-text("Assignments"), button:has-text("Aufgaben")').first
        try:
            btn.click(timeout=15000)
        except Exception:
            raise TeamsError("the Assignments button was not found in Teams")
        end = time.time() + 25
        while not self.edu_auth and time.time() < end:
            self.pg.wait_for_timeout(250)
        if not self.edu_auth:
            raise TeamsError("the Assignments app did not load")
        return self.edu_auth

    def assignments(self):
        """Only your own assignments, each with a bucket: 'due', 'missed' or 'done'."""
        self.boot()
        url = self.EDU_API + "?$top=100&$expand=" + urllib.parse.quote("submissions($expand=outcomes)")
        raw, retried = [], False
        while url:
            r = self.b.ctx.request.get(url, headers=self.edu_token(), fail_on_status_code=False)
            if r.status == 401 and not retried:
                self.edu_auth, retried = None, True
                continue
            if r.status != 200:
                raise TeamsError(f"Assignments {r.status}: {r.status_text}")
            d = r.json()
            raw += d.get("value") or []
            url = d.get("@odata.nextLink")
        classes = {j["groupId"]: course_name(j["name"]) for j in self.joined() if j.get("groupId")}
        now = datetime.now().astimezone()
        out = []
        for d in raw:
            if d.get("status") == "draft":
                continue
            subs = d.get("submissions") or []
            sub = subs[0] if subs and isinstance(subs[0], dict) else {}
            sstat = sub.get("status") or ""
            due = sub.get("extendedDueDateTime") or d.get("dueDateTime")
            try:
                late = bool(due) and datetime.fromisoformat(due.replace("Z", "+00:00")) < now
            except ValueError:
                late = False
            done = sstat in ("submitted", "returned", "excused") or d.get("isCompleted")
            bucket = "done" if done else "missed" if late else "due"
            points, maxp = None, (d.get("grading") or {}).get("maxPoints") if isinstance(d.get("grading"), dict) else None
            for o in sub.get("outcomes") or []:
                pts = o.get("points") if isinstance(o, dict) else None
                if isinstance(pts, dict) and pts.get("points") is not None:
                    points = pts["points"]
            instr = d.get("instructions")
            text = html2text(instr.get("content"))[0] if isinstance(instr, dict) else ""
            links = forms_links(json.dumps(d, ensure_ascii=False))
            fid = find_key(d, r"formid|formsid")
            if fid and len(fid) > 20:
                links.append(f"https://forms.office.com/Pages/ResponsePage.aspx?id={fid}")
            out.append({"id": d.get("id"), "title": d.get("displayName") or "", "due": due,
                        "class": classes.get(d.get("classId"), ""), "status": sstat or d.get("status") or "",
                        "bucket": bucket, "done": bool(done), "text": text, "forms": list(dict.fromkeys(links)),
                        "points": points, "max": maxp, "web": sub.get("webUrl") or d.get("webUrl")})
        for a in out:
            for u in a["forms"]:
                self.forms.setdefault(u, {"label": a["title"], "where": a["class"] or "Assignments", "time": a["due"] or ""})
        return sorted(out, key=lambda a: a["due"] or "9999")

    # -- diagnostics -------------------------------------------------------- #
    def doctor(self):
        self.b.start()
        if not self.hooked:
            self.b.ctx.on("response", self._on_response)
            self.hooked = True
        pg = self.b.new_page()
        print("opening teams.microsoft.com (up to 90 s)…")
        try:
            pg.goto("https://teams.microsoft.com/v2/", wait_until="commit", timeout=60000)
        except Exception as e:
            print(red(f"  could not open it: {e}"))
        self.wait(pg, 90, lambda: len(self.teams()), stable=4)
        hosts = {}
        for c in self.caps:
            hosts[c["host"]] = hosts.get(c["host"], 0) + 1
        print(f"  page:   {pg.url[:100]}")
        print(f"  title:  {pg.title()}")
        print(f"  JSON responses read: {len(self.caps)}")
        for h, n in sorted(hosts.items(), key=lambda x: -x[1])[:10]:
            print(dim(f"    {n:>4}  {h}"))
        ts = self.teams()
        print((green if ts else red)(f"  teams found in the web app: {len(ts)}"))
        for t in list(ts.values())[:5]:
            print(dim(f"    {t['name']}  ({len(t['channels'])} channels)"))
        try:
            shown = pg.inner_text("body", timeout=3000).strip().split("\n")[:5]
        except Exception:
            shown = []
        print(dim("  page text: " + " | ".join(x.strip() for x in shown if x.strip())[:160]))
        log = STATE_DIR / "doctor.txt"
        STATE_DIR.mkdir(parents=True, exist_ok=True)
        log.write_text("\n".join(c["url"] for c in self.caps))   # addresses only, no content, no query strings
        print(dim(f"  list of loaded addresses: {log}"))
        pg.close()
        self.ready = bool(ts)


def edit_text(initial=""):
    import tempfile
    editor = os.environ.get("EDITOR") or os.environ.get("VISUAL") or "nano"
    with tempfile.NamedTemporaryFile("w+", suffix=".txt", delete=False) as f:
        f.write(initial)
        name = f.name
    try:
        subprocess.call(shlex.split(editor) + [name])
        return Path(name).read_text().rstrip("\n")
    finally:
        os.unlink(name)


# --------------------------------------------------------------------------- #
#  Shell
# --------------------------------------------------------------------------- #

HELP = f"""
{bold('teams')} — Microsoft Teams in the terminal

  {cyan('teams')}            full-screen interface (arrow keys, Enter, Esc — press ? inside)
  {cyan('teams shell')}      command prompt (below)
  {cyan('teams <command>')}  run one command, e.g.  teams sync all

{bold('Account')}
  login                  sign in here in the terminal (e-mail, password, MFA number/code)
  login --window         sign in in a visible Chrome window (for unusual school logins)
  logout                 delete the saved session
  whoami                 show who is signed in

{bold('Teams & files')}           (numbers refer to the last list shown)
  teams [filter]         your classes of this school year
  teams --all            also earlier years and archived teams
  teams --sites          every team site of the school you can open
  team <n|name>          open a team and show its channel folders
  ls                     list the current folder again
  cd <n|..|/>            go into a folder, up one, or back to the team's top
  get <n...|all>         download files/folders into ~/Moodle/<subject>
  sync [team|all]        download everything new or changed of a team (all = every mapped team)
  view <n>               download a file to /tmp and open it
  url <n>                print the web link of an item
  libs                   other document libraries of this team (e.g. Class Materials)
  hide <n...> / unhide   hide teams you don't care about from `teams`

{bold('Moodle folders')}
  subject [folder]       show/set the ~/Moodle subfolder of the open team
  subjects               show all team → folder mappings
  dir [path]             show/set the download root (default: same as moodle-cli)

{bold('Posts, assignments, Forms tests')}
  posts [n]              channels of the open team / posts of channel n (Forms links marked ▶)
  assignments [--done]   your assignments: to do, missed (and done)
  forms                  all Forms links found so far
  form <url>             answer a Microsoft Forms quiz/test in the terminal
  doctor                 check that the Teams web app loads (needed for posts/assignments)

{bold('Other')}
  web [n]                open the team or an item in your normal browser
  refresh                reload the team list      quit / exit / Ctrl-D
"""


class Shell(cmd.Cmd):
    def __init__(self, browser):
        super().__init__()
        self.cfg = load_config()
        self.b = browser
        self.sp = SharePoint(browser, self.cfg)
        self.web = TeamsWeb(browser)
        self._teams = None
        self._joined = None
        self.courses_note = ""
        self.listed = []          # last list of teams shown
        self.team = None
        self.lib = None
        self.stack = []           # folder paths below the library root
        self.items = []
        self.update_prompt()

    # -- plumbing ----------------------------------------------------------- #
    def update_prompt(self):
        where = ""
        if self.team:
            where = self.team["title"][:30] + ("/" + "/".join(p.rsplit("/", 1)[-1] for p in self.stack) if self.stack else "")
        self.prompt = (f"\001\033[1;35m\002teams\001\033[0m\002{':' + where if where else ''}> "
                       if USE_COLOR else f"teams{':' + where if where else ''}> ")

    def onecmd(self, line):
        try:
            return super().onecmd(line)
        except TeamsError as e:
            err(e)
        except KeyboardInterrupt:
            print()
        except EOFError:
            print()
        except Exception as e:  # Playwright errors, network trouble
            err(f"{type(e).__name__}: {str(e).splitlines()[0] if str(e) else ''}")
        return False

    def emptyline(self):
        return False

    def default(self, line):
        err(f"unknown command: {line.split()[0]}  (type `help`)")

    def do_help(self, arg):
        if arg:
            return super().do_help(arg)
        page(HELP.strip())

    def do_quit(self, arg):
        """Leave the program."""
        return True

    do_exit = do_quit

    def do_EOF(self, arg):
        print()
        return True

    def need(self):
        if not self.cfg.get("host"):
            raise TeamsError("not signed in — run `login`")
        self.b.start()
        return self.sp

    def need_team(self):
        self.need()
        if not self.team:
            raise TeamsError("open a team first: `teams`, then `team <n>`")

    # -- account ------------------------------------------------------------ #
    def do_login(self, arg):
        """login [--window] — sign in with your school Microsoft account."""
        window = "--window" in arg
        start = "https://www.office.com/launch/onedrive"   # ends on <school>-my.sharepoint.com

        def done(pg):
            h = urllib.parse.urlparse(pg.url).hostname or ""
            return h.endswith("-my.sharepoint.com")

        if window:
            self.b.close()
            vis = Browser(headless=False).start()
            try:
                pg = vis.page()
                pg.goto(start)
                print("Sign in in the Chrome window. It closes by itself when you are done.")
                while not done(pg):
                    pg.wait_for_timeout(1000)
                host = urllib.parse.urlparse(pg.url).hostname
            finally:
                vis.close()
            self.b.start()
        else:
            self.b.start()
            pg = self.b.page()
            print(dim("Signing in to Microsoft 365 (your password goes only to login.microsoftonline.com)."))
            terminal_login(pg, start, done)
            host = urllib.parse.urlparse(pg.url).hostname
        self.cfg["host"] = host.replace("-my.sharepoint.com", ".sharepoint.com")
        save_config(self.cfg)
        try:
            me = self.sp.me()
            print(green(f"✓ signed in as {me.get('DisplayName')} <{me.get('Email')}>"))
        except TeamsError:
            print(green(f"✓ signed in ({self.cfg['host']})"))

    def do_logout(self, arg):
        """Delete the saved browser session."""
        self.b.close()
        shutil.rmtree(PROFILE_DIR, ignore_errors=True)
        self.cfg.pop("host", None)
        save_config(self.cfg)
        print("signed out")

    def do_whoami(self, arg):
        """Show the signed-in account."""
        me = self.need().me()
        print(f"{bold(me.get('DisplayName'))} <{me.get('Email')}>  ·  {self.cfg['host']}")
        print(dim(f"downloads: {moodle_root(self.cfg)}  ·  config: {CONFIG_FILE}"))

    # -- teams -------------------------------------------------------------- #
    def teams(self, reload=False):
        if self._teams is None or reload:
            info("loading teams…")
            self._teams = self.need().team_sites()
        return self._teams

    def joined(self, reload=False):
        """Teams you are a member of. Cached in the config so the list shows at once next time."""
        cached = self.cfg.get("joined")
        if self._joined is None and cached and not reload:
            self._joined = cached.get("teams") or []
        if self._joined is None or reload:
            info("finding the classes you are in…")
            self.need()
            self.web.reload_list() if reload else self.web.boot()
            j = self.web.joined()
            if not j:
                raise TeamsError("Teams did not list your teams")
            self._joined = j
            self.cfg["joined"] = {"time": int(time.time()), "teams": j}
            save_config(self.cfg)
        return self._joined

    def is_hidden(self, url):
        return url.lower() in {h.lower() for h in self.cfg["hidden"]}

    def courses(self, reload=False):
        """Your teams as site dicts. `old` marks archived teams and teams of an earlier school year."""
        keys = {k.lower(): k for k in self.cfg["subjects"]}
        cy = school_year()
        self.courses_note = ""
        try:
            src = [(j["name"], j["site"], j["archived"], j["desc"]) for j in self.joined(reload) if j.get("site")]
        except TeamsError as e:
            self.courses_note = f"{e}. Showing every SharePoint team site of {school_label()} instead."
            src = [(t["title"], t["url"], False, t["desc"]) for t in self.teams(reload)
                   if cy in title_years(t["title"])]
        out = []
        for title, url, archived, desc in src:
            years = title_years(title)
            out.append({"title": title, "url": keys.get(url.lower(), url), "channel": False, "desc": desc,
                        "modified": None, "archived": archived, "old": archived or bool(years and cy not in years),
                        "name": course_name(title), "code": class_code(title)})
        out.sort(key=lambda t: t["name"].lower())
        return out

    def do_refresh(self, arg):
        """Reload the team list."""
        self._teams = None
        self.courses(reload=True)
        self.do_teams("")

    def do_teams(self, arg):
        """teams [--all|--sites] [filter] — list your classes of this school year."""
        a = args_of(arg)
        everything, sites = "--all" in a or "-a" in a, "--sites" in a
        f = norm(" ".join(x for x in a if not x.startswith("-")))
        if sites:
            ts = self.teams()
        else:
            ts = [t for t in self.courses() if everything or not t["old"]]
            if self.courses_note:
                print(yellow(self.courses_note))
        ts = [t for t in ts if not self.is_hidden(t["url"]) and (not f or f in norm(t["title"]))]
        self.listed = ts
        if not ts:
            print(dim("no teams" + (" match" if f else "")))
            return
        nw = min(30, max(len(t.get("name") or t["title"]) for t in ts))
        for n, t in enumerate(ts, 1):
            subj = self.cfg["subjects"].get(t["url"])
            name = t.get("name") or t["title"]
            tag = yellow("  archived") if t.get("archived") else dim("  earlier year") if t.get("old") else ""
            print(f"{cyan(f'{n:>3}')}  {bold(name.ljust(nw))}  {dim(t['title'])}{tag}"
                  + (f"  {green('→ ' + subj)}" if subj else ""))
        if not sites and not everything:
            old = sum(1 for t in self.courses() if t["old"])
            if old:
                print(dim(f"     {old} from earlier years or archived · `teams --all` shows them"))
        pool = self.teams() if sites else self.courses()
        hidden = sum(1 for t in pool if self.is_hidden(t["url"]))
        if hidden:
            print(dim(f"     {hidden} hidden · `unhide` to show them"))

    def pick_team(self, arg):
        arg = arg.strip()
        if not arg:
            raise TeamsError("which team? give a number from `teams` or part of its name")
        if arg.isdigit():
            if not self.listed:
                self.do_teams("")
            n = int(arg)
            if not 1 <= n <= len(self.listed):
                raise TeamsError(f"no team {n}")
            return self.listed[n - 1]
        hits = [t for t in self.courses() if norm(arg) in norm(t["title"]) or norm(arg) in norm(t["name"])]
        if not hits:
            hits = [t for t in self.teams() if norm(arg) in norm(t["title"])]
        if len(hits) == 1:
            return hits[0]
        if not hits:
            raise TeamsError(f"no team matches '{arg}'")
        self.listed = hits
        self.do_teams(arg)
        raise TeamsError("several teams match, pick a number")

    def do_team(self, arg):
        """team <n|name> — open a team."""
        t = self.pick_team(arg)
        self.team, self.stack = t, []
        libs = self.sp.libraries(t["url"])
        if not libs:
            raise TeamsError("this team has no files you can see")
        self.lib = libs[0]
        self.update_prompt()
        print(bold(t["title"]) + (dim("  " + t["desc"][:80]) if t["desc"] else ""))
        subj = self.cfg["subjects"].get(t["url"])
        print(dim(f"files → {moodle_root(self.cfg) / subj}" if subj else "files → (no Moodle folder chosen yet, asked on first download)"))
        self.do_ls("")

    do_open = do_team

    def cur_path(self):
        return self.stack[-1] if self.stack else self.lib["path"]

    def do_ls(self, arg):
        """List the current folder."""
        self.need_team()
        self.items = self.sp.folder(self.team["url"], self.cur_path())
        if not self.items:
            print(dim("  (empty folder)"))
        for n, it in enumerate(self.items, 1):
            if it["kind"] == "dir":
                print(f"{cyan(f'{n:>3}')}  {blue(it['name'] + '/')}  {dim(str(it['count']) + ' items')}")
            else:
                print(f"{cyan(f'{n:>3}')}  {it['name']}  {dim(human(it['size']) + '  ' + fmt_time(it['modified']))}")

    def do_cd(self, arg):
        """cd <n|..|/> — enter a folder, go up, or go to the top."""
        self.need_team()
        a = arg.strip()
        if a in ("..", "-"):
            if self.stack:
                self.stack.pop()
        elif a in ("/", "", "~"):
            self.stack = []
        else:
            n = parse_nums([a], len(self.items))[0] if a.isdigit() else \
                next((i + 1 for i, it in enumerate(self.items) if it["kind"] == "dir" and norm(a) == norm(it["name"])), 0)
            if not n or self.items[n - 1]["kind"] != "dir":
                raise TeamsError("not a folder")
            self.stack.append(self.items[n - 1]["path"])
        self.update_prompt()
        self.do_ls("")

    def do_libs(self, arg):
        """libs [n] — list or switch document libraries of the open team."""
        self.need_team()
        libs = self.sp.libraries(self.team["url"])
        if arg.strip().isdigit():
            self.lib = libs[parse_nums([arg.strip()], len(libs))[0] - 1]
            self.stack = []
            self.update_prompt()
            return self.do_ls("")
        for n, l in enumerate(libs, 1):
            cur = green(" ◀") if l["path"] == self.lib["path"] else ""
            print(f"{cyan(f'{n:>3}')}  {l['title']}  {dim(str(l['count']) + ' items')}{cur}")
        print(dim("libs <n> to switch"))

    def do_hide(self, arg):
        """hide <n...> — hide teams from the `teams` list."""
        for n in parse_nums(args_of(arg), len(self.listed)):
            self.cfg["hidden"].append(self.listed[n - 1]["url"])
        save_config(self.cfg)
        self.do_teams("")

    def do_unhide(self, arg):
        """Show all hidden teams again."""
        self.cfg["hidden"] = []
        save_config(self.cfg)
        self.do_teams("")

    # -- Moodle subject folders --------------------------------------------- #
    def moodle_dirs(self):
        root = moodle_root(self.cfg)
        return sorted([p.name for p in root.iterdir() if p.is_dir() and not p.name.startswith(".")],
                      key=str.lower) if root.is_dir() else []

    ALIASES = {"chemie": "chemistry", "geschichte": "history", "deutsch": "german", "englisch": "english",
               "mathematik": "am", "mathe": "am", "rechnungswesen": "accounting", "buchhaltung": "accounting",
               "physik": "physics", "biologie": "biology", "geografie": "geography", "geographie": "geography",
               "wirtschaft": "wlc", "programmieren": "pose", "softwareentwicklung": "pose", "netzwerk": "nscs"}

    def guess_subject(self, title):
        dirs = self.moodle_dirs()
        t = norm(title)
        for de, en in self.ALIASES.items():
            if de in t:
                t += en
        tokens = [norm(w) for w in re.split(r"[\s_\-–|.,()]+", title) if len(norm(w)) >= 3]
        years = re.findall(r"\d{4}", title.replace("/", "").replace("_", ""))

        def score(d):
            nd = norm(d)
            s = difflib.SequenceMatcher(None, t, nd).ratio()
            word = re.match(r"[a-z]+", nd)
            if word and len(word[0]) >= 2 and word[0] in t:
                s += 0.6 if len(word[0]) >= 3 else 0.3               # "English2627" vs "2IHIF English 26/27"
            s += 0.15 * sum(1 for w in tokens if w in nd)
            if years and any(y in nd for y in years):
                s += 0.3
            return s
        ranked = sorted(dirs, key=score, reverse=True)
        return ranked, (score(ranked[0]) if ranked else 0)

    def subject_for(self, team, ask=True):
        subj = self.cfg["subjects"].get(team["url"])
        if subj or not ask:
            return subj
        ranked, sc = self.guess_subject(team["title"])
        root = moodle_root(self.cfg)
        print(f"Where should files of {bold(team['title'])} go?")
        top = ranked[:8]
        for n, d in enumerate(top, 1):
            print(f"  {cyan(str(n))}. {root / d}" + (green("  (best match)") if n == 1 and sc > 0.8 else ""))
        print(f"  {cyan('n')}. new folder    {cyan('l')}. list all folders")
        while True:
            s = input(f"choose [{'1' if top else 'n'}]: ").strip() or ("1" if top else "n")
            if s == "l":
                for n, d in enumerate(ranked, 1):
                    print(f"  {cyan(str(n))}. {d}")
                top = ranked
                continue
            if s == "n":
                default = safe_name(re.sub(r"\s+", "", team["title"]))
                subj = safe_name(input(f"folder name [{default}]: ").strip() or default)
                break
            if s.isdigit() and 1 <= int(s) <= len(top):
                subj = top[int(s) - 1]
                break
        self.cfg["subjects"][team["url"]] = subj
        save_config(self.cfg)
        print(green(f"→ {root / subj}") + dim("  (change with `subject`)"))
        return subj

    def do_subject(self, arg):
        """subject [folder] — show or change the ~/Moodle subfolder of the open team."""
        self.need_team()
        if arg.strip():
            self.cfg["subjects"][self.team["url"]] = safe_name(arg.strip())
            save_config(self.cfg)
        elif self.cfg["subjects"].get(self.team["url"]) and not confirm(
                f"now: {self.cfg['subjects'][self.team['url']]} — choose a different folder?"):
            return
        else:
            self.cfg["subjects"].pop(self.team["url"], None)
        self.subject_for(self.team)

    def do_subjects(self, arg):
        """Show all team → folder mappings."""
        names = {t["url"]: t["title"] for t in (self._teams or []) + (self.courses() if self._joined else [])}
        if not self.cfg["subjects"]:
            print(dim("none yet — they are asked on the first download from a team"))
        for url, d in sorted(self.cfg["subjects"].items(), key=lambda x: x[1].lower()):
            print(f"  {green(d):<30} ← {names.get(url, url)}")

    def do_dir(self, arg):
        """dir [path] — show or set the download root."""
        if arg.strip():
            self.cfg["download_dir"] = str(Path(os.path.expanduser(arg.strip())).resolve())
            save_config(self.cfg)
        print(moodle_root(self.cfg))

    # -- downloads ---------------------------------------------------------- #
    def save_file(self, team, it, dest, quiet=False):
        """Download one file; skip it if the same size is already there. Returns 'new'|'updated'|'same'."""
        dest.parent.mkdir(parents=True, exist_ok=True)
        status = "new"
        if dest.exists():
            if dest.stat().st_size == it["size"]:
                return "same"
            status = "updated"
        data = self.sp.download(team["url"], it["path"])
        tmp = dest.with_name(dest.name + ".part")
        tmp.write_bytes(data)
        tmp.replace(dest)
        try:
            ts = datetime.fromisoformat(it["modified"].replace("Z", "+00:00")).timestamp()
            os.utime(dest, (ts, ts))
        except (TypeError, ValueError, AttributeError):
            pass
        if not quiet:
            print(f"  {green('✓') if status == 'new' else yellow('↻')} {dest}  {dim(human(len(data)))}")
        return status

    def do_get(self, arg):
        """get <n...|all> — download files or whole folders into ~/Moodle/<subject>."""
        self.need_team()
        if not self.items:
            self.do_ls("")
        picks = parse_nums(args_of(arg) or ["all"], len(self.items))
        base = moodle_root(self.cfg) / self.subject_for(self.team)
        counts = {"new": 0, "updated": 0, "same": 0}
        for n in picks:
            it = self.items[n - 1]
            if it["kind"] == "file":
                counts[self.save_file(self.team, it, base / safe_name(it["name"]))] += 1
            else:
                for rel, f in self.sp.walk(self.team["url"], it["path"]):
                    dest = base / safe_name(it["name"]) / Path(*[safe_name(p) for p in rel.split("/")])
                    counts[self.save_file(self.team, f, dest)] += 1
        print(dim(f"{counts['new']} new, {counts['updated']} updated, {counts['same']} already there  →  {base}"))

    def do_sync(self, arg):
        """sync [team|all] — download all new/changed files of a team (all = every team with a Moodle folder)."""
        self.need()
        a = arg.strip()
        if a == "all":
            by_url = {t["url"].lower(): t for t in self.courses()}
            if any(u.lower() not in by_url for u in self.cfg["subjects"]):
                by_url = {**{t["url"].lower(): t for t in self.teams()}, **by_url}
            targets = [by_url[u.lower()] for u in self.cfg["subjects"] if u.lower() in by_url]
            if not targets:
                raise TeamsError("no team has a Moodle folder yet — `team <n>` then `get` or `sync` first")
        elif a:
            targets = [self.pick_team(a)]
        else:
            self.need_team()
            targets = [self.team]
        for t in targets:
            subj = self.subject_for(t)
            base = moodle_root(self.cfg) / subj
            print(bold(f"{t['title']}") + dim(f"  →  {base}"))
            counts = {"new": 0, "updated": 0, "same": 0}
            for lib in self.sp.libraries(t["url"]):
                for rel, f in self.sp.walk(t["url"], lib["path"]):
                    parts = [safe_name(p) for p in rel.split("/")]
                    # channel folder level ("General", "Allgemein") is dropped: files land in the subject folder
                    if len(parts) > 1:
                        parts = parts[1:]
                    counts[self.save_file(t, f, base.joinpath(*parts))] += 1
            print(dim(f"  {counts['new']} new, {counts['updated']} updated, {counts['same']} unchanged"))

    def do_view(self, arg):
        """view <n> — open a file with the default program."""
        self.need_team()
        it = self.items[parse_nums([arg.strip()], len(self.items))[0] - 1]
        if it["kind"] != "file":
            raise TeamsError("that is a folder, use `cd`")
        import tempfile
        d = Path(tempfile.mkdtemp(prefix="teams-"))
        p = d / safe_name(it["name"])
        p.write_bytes(self.sp.download(self.team["url"], it["path"]))
        if p.suffix.lower() == ".pdf" and shutil.which("pdftotext") and not os.environ.get("DISPLAY") \
                and not os.environ.get("WAYLAND_DISPLAY"):
            page(subprocess.run(["pdftotext", "-layout", str(p), "-"], capture_output=True, text=True).stdout)
            return
        subprocess.Popen(["xdg-open", str(p)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print(dim(f"opened {p}"))

    def do_url(self, arg):
        """url <n> — print the web link of an item."""
        self.need_team()
        it = self.items[parse_nums([arg.strip()], len(self.items))[0] - 1]
        print(self.sp.web_url(self.team["url"], it["path"]))

    def do_web(self, arg):
        """web [n] — open the team (or item n) in your browser."""
        self.need_team()
        url = self.team["url"]
        if arg.strip():
            url = self.sp.web_url(self.team["url"], self.items[parse_nums([arg.strip()], len(self.items))[0] - 1]["path"])
        subprocess.Popen(["xdg-open", url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    # -- forms -------------------------------------------------------------- #
    def do_form(self, arg):
        """form <url> — take a Microsoft Forms test/quiz in the terminal."""
        url = arg.strip() or input("Forms link: ").strip()
        host = urllib.parse.urlparse(url).hostname or ""
        if not re.search(r"(^|\.)forms\.(office|cloud\.microsoft|microsoft)\.com$|^forms\.cloud\.microsoft$|^forms\.office\.com$", host):
            raise TeamsError("that is not a Microsoft Forms link (forms.office.com / forms.cloud.microsoft)")
        self.b.start()
        fr = FormRunner(self.b, url)
        try:
            fr.run()
        finally:
            fr.close()

    do_quiz = do_form

    def do_forms(self, arg):
        """List the Forms links found in Teams so far (open channels or `assignments` first)."""
        if not self.web.forms:
            print(dim("none yet — open a channel (`posts`) or run `assignments`; or use `form <link>`"))
        self.form_list = list(self.web.forms.items())
        for n, (u, f) in enumerate(self.form_list, 1):
            print(f"{cyan(f'{n:>3}')}  {f['label']}  {dim(f['where'])}\n       {dim(u)}")

    def do_posts(self, arg):
        """posts [channel n] — show the posts of a channel of the open team."""
        self.need_team()
        self.web.boot()
        wt = self.web.match(self.team["title"])
        if not wt:
            raise TeamsError("this team was not found in the Teams web app")
        if not arg.strip():
            for n, ch in enumerate(wt["channels"], 1):
                print(f"{cyan(f'{n:>3}')}  # {ch['name']}")
            print(dim("posts <n> to read a channel"))
            return
        ch = wt["channels"][parse_nums([arg.strip()], len(wt["channels"]))[0] - 1]
        out = []
        for p in self.web.channel(wt, ch):
            out.append(f"{bold(p['author'])}  {dim(fmt_time(p['time']))}" + (f"\n  {bold(p['subject'])}" if p["subject"] else ""))
            out.append(_wrap(p["text"], "  "))
            for r in p["replies"]:
                out.append(dim(f"    ↳ {r['author']}  {fmt_time(r['time'])}"))
                out.append(_wrap(r["text"], "      "))
            for m in [p] + p["replies"]:
                for u in m["forms"]:
                    out.append(yellow(f"  ▶ Form: {u}"))
            out.append("")
        page("\n".join(out) or dim("no posts loaded"))

    def do_assignments(self, arg):
        """assignments [--done] — your assignments: to do, missed, and (with --done) done."""
        old = (load_assign_cache() or {}).get("items") or []
        items = self.web.assignments()
        save_assign_cache(items)
        added, changed = diff_assignments(old, items) if old else (set(), set())
        if old:
            print(dim(f"{len(added)} new, {len(changed)} changed since last time"))
        show = ("due", "missed", "done") if "--done" in arg else ("due", "missed")
        for k, label, col in (("due", "To do", yellow), ("missed", "Missed", red), ("done", "Done", green)):
            group = [a for a in items if a["bucket"] == k]
            if k == "missed" or k == "done":
                group.reverse()
            print(col(bold(f"{label} ({len(group)})")))
            if k not in show:
                print(dim("  `assignments --done` lists them"))
                continue
            for a in group:
                tag = magenta("  new") if a["id"] in added else yellow("  changed") if a["id"] in changed else ""
                print(f"  {bold(a['title'])}  {cyan(a['class'])}  {dim(fmt_due(a['due']))}{tag}"
                      + (f"  {green(score_text(a))}" if score_text(a) else ""))
                for u in a["forms"]:
                    print(yellow(f"     ▶ Form: {u}"))

    def do_doctor(self, arg):
        """Check whether the Teams web app (posts, assignments) loads in the hidden browser."""
        self.need()
        self.web.doctor()

    def do_tui(self, arg):
        """Open the full-screen interface."""
        run_tui(self)


# =========================================================================== #
#  Full-screen TUI (curses), same keys as moodle-cli
# =========================================================================== #

try:
    import curses
except ImportError:
    curses = None

ANSI_RE = re.compile(r"\033\[([\d;]*)m")

TUI_HELP = """
Moving around
  ↑ ↓   j k            move up / down
  Enter  →  l          open
  Esc  ←  h  Backspace go back
  PgUp PgDn  Space     page up / down
  g  G  Home End       top / bottom
  1 2 3   Tab  S-Tab   switch tab
  /                    filter the list (Esc clears)
  r                    reload
  mouse                click selects, click again opens, wheel scrolls
  q                    quit

Actions (the bottom line shows the ones for the current screen)
  Courses              s  download all files of the team   m  choose its ~/Moodle folder   x  hide it
                       A  show classes of earlier years and archived teams
  Team / folder        Enter open   d  download   a  download everything here   o  open file
  Channel              Enter on a post: read it with replies   Enter on ▶ Form: take the test
  Forms                Enter take the test in the terminal   p  paste a Forms link
  Assignments          To do, Missed, Done.  Enter details   o open in Teams   D show done

Files go to ~/Moodle/<subject> (the same folder moodle-cli uses). The subject folder
is asked once per team and remembered; `m` changes it.
Courses lists the teams you are a member of for this school year, taken from
Teams itself. Press r on Courses to reload that list.

Forms tests: ↑ ↓ pick an answer, Space ticks it, Enter goes on, ← goes back.
At the end you see all answers and confirm before anything is submitted.
"""


def strip_ansi(s):
    return ANSI_RE.sub("", s)


def char_w(ch):
    import unicodedata
    if unicodedata.combining(ch):
        return 0
    return 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1


class Row:
    __slots__ = ("text", "data", "sel", "kind")

    def __init__(self, text, data=None, sel=None, kind=""):
        self.text = str(text).replace("\t", "    ")
        self.data = data
        self.sel = data is not None if sel is None else sel
        self.kind = kind          # "head" draws a rule after the text


def text_rows(text):
    if isinstance(text, (list, tuple)):
        text = "\n".join(text)
    return [Row(l, sel=False) for l in str(text).split("\n")]


def wrap_rows(text, indent="    ", limit=None):
    import textwrap
    out = []
    for line in str(text).split("\n"):
        for w in (textwrap.wrap(line, max(20, width() - len(indent) - 2)) or [""]):
            out.append(Row(indent + w, sel=False))
    if limit and len(out) > limit:
        out = out[:limit] + [Row(indent + dim("…"), sel=False)]
    return out


def head(text):
    return Row("  " + cyan(bold(text)), sel=False, kind="head")


class View:
    hint = ""

    def __init__(self, app, title):
        self.app, self.sh, self.title = app, app.sh, title
        self.all_rows, self.rows = [], []
        self.sel, self.top = -1, 0
        self.filter = ""
        self.loaded = False

    def build(self):
        return []

    def load(self):
        keep = self.sel if self.loaded else None
        self.all_rows = self.build() or text_rows(dim("(empty)"))
        self.loaded = True
        self.apply_filter(keep)

    def apply_filter(self, keep=None):
        if self.filter:
            q = self.filter.lower()
            self.rows = [r for r in self.all_rows if r.sel and q in strip_ansi(r.text).lower()]
        else:
            self.rows = self.all_rows
        sels = self.selectable()
        if keep is not None and keep in sels:
            self.sel = keep
        else:
            self.sel = sels[0] if sels else -1
            self.top = 0
        self.clamp()

    def selectable(self):
        return [i for i, r in enumerate(self.rows) if r.sel]

    @property
    def current(self):
        return self.rows[self.sel].data if 0 <= self.sel < len(self.rows) else None

    def enter(self, data):
        pass

    def key(self, k, data):
        return False

    def bh(self):
        return self.app.body_h()

    def clamp(self):
        self.top = max(0, min(self.top, len(self.rows) - self.bh()))

    def show_sel(self):
        bh = self.bh()
        if self.sel < 0:
            return
        if self.sel >= self.top + bh:
            self.top = self.sel - bh + 1
        if self.sel < self.top:
            self.top = self.sel
            j = self.sel - 1
            while j >= 0 and not self.rows[j].sel and self.sel - j <= 3 and strip_ansi(self.rows[j].text).strip():
                self.top = j
                j -= 1
        self.clamp()

    def move(self, d):
        bh = self.bh()
        if self.sel < 0:
            self.top += d
            return self.clamp()
        rng = range(self.sel + 1, len(self.rows)) if d > 0 else range(self.sel - 1, -1, -1)
        nxt = next((j for j in rng if self.rows[j].sel), None)
        if nxt is not None and abs(nxt - self.sel) <= max(4, bh // 2):
            self.sel = nxt
            return self.show_sel()
        old = self.top
        self.top += d
        self.clamp()
        if self.top == old and nxt is not None:
            self.sel = nxt
            self.show_sel()
        elif nxt is not None and self.top <= nxt < self.top + bh and (d > 0) == (nxt > self.sel):
            if (d > 0 and nxt == self.top + bh - 1) or (d < 0 and nxt == self.top):
                self.sel = nxt

    def scroll(self, n):
        self.top += n
        self.clamp()
        bh = self.bh()
        if self.sel >= 0 and not (self.top <= self.sel < self.top + bh):
            vis = [i for i in self.selectable() if self.top <= i < self.top + bh]
            if vis:
                self.sel = vis[0] if n > 0 else vis[-1]

    def home(self):
        self.top = 0
        s = self.selectable()
        if s:
            self.sel = s[0]
            self.show_sel()

    def end(self):
        self.top = len(self.rows)
        self.clamp()
        s = self.selectable()
        if s:
            self.sel = s[-1]
            self.show_sel()


class TextView(View):
    def __init__(self, app, title, builder, hint=""):
        super().__init__(app, title)
        self.builder, self.hint = builder, hint

    def build(self):
        b = self.builder()
        return b if b and isinstance(b[0], Row) else text_rows(b)


class HelpView(TextView):
    def __init__(self, app):
        super().__init__(app, "Help", lambda: TUI_HELP.strip("\n"))


def form_row(url, info=None, indent="    "):
    label = (info or {}).get("label") or "Microsoft Form"
    return Row(f"{indent}{yellow('▶ Form')}  {label}  {dim(url[:60])}", {"form": url})


class CoursesView(View):
    hint = "Enter open · s sync files · m folder · x hide · A earlier years · / filter"

    def __init__(self, app):
        super().__init__(app, "Courses")
        self.show_old = False
        self.fresh = False        # True: ask Teams again instead of using the saved list

    def build(self):
        if not self.sh._joined and not self.sh.cfg.get("joined"):
            self.app.status(dim("finding the classes you are in (the first time takes up to a minute)…"))
        every = self.sh.courses(reload=self.fresh)
        self.fresh = False
        courses = [t for t in every if not self.sh.is_hidden(t["url"])]
        hidden = len(every) - len(courses)
        cur = [t for t in courses if not t["old"]]
        old = [t for t in courses if t["old"]]
        nw = min(28, max([len(t["name"]) for t in courses] or [10]))
        rows = []
        if self.sh.courses_note:
            rows += wrap_rows(yellow(self.sh.courses_note), "  ") + [Row("")]
        rows.append(head(f"Your classes {school_label()}  ·  {len(cur)}"))
        rows += [self.course_row(t, nw) for t in cur] or [Row(dim("  no classes for this school year"))]
        if old:
            rows.append(Row(""))
            if self.show_old:
                rows.append(head(f"Earlier years and archived  ·  {len(old)}"))
                rows += [self.course_row(t, nw) for t in old]
            else:
                rows.append(Row(dim(f"  + {len(old)} from earlier years or archived  (A shows them)"), {"toggle": True}))
        if hidden:
            rows += [Row(""), Row(dim(f"  {hidden} hidden  (X shows them again)"))]
        return rows

    def course_row(self, t, nw):
        subj = self.sh.cfg["subjects"].get(t["url"])
        name = t["name"][:nw]
        tag = yellow("  archived") if t.get("archived") else ""
        chip = f"  {green('↓ ' + subj)}" if subj else ""
        dot = dim("○") if t["old"] else magenta("●")
        return Row(f"  {dot} {bold(name)}{' ' * (nw - len(name) + 2)}{dim(t['title'])}{tag}{chip}", t)

    def enter(self, t):
        if t.get("toggle"):
            return self.toggle()
        self.app.push(TeamView(self.app, t))

    def toggle(self):
        self.show_old = not self.show_old
        self.all_rows = self.build()
        self.apply_filter(self.sel)

    def key(self, k, t):
        if k == "X":
            self.sh.cfg["hidden"] = []
            save_config(self.sh.cfg)
            self.app.load(self)
            return True
        if k == "A":
            self.toggle()
            return True
        if k in ("r", curses.KEY_F5):
            self.fresh = True
            self.app.load(self)
            self.app.msg = self.app.msg or green("✓ class list reloaded from Teams")
            return True
        if t is None or t.get("toggle"):
            return False
        if k == "s":
            self.app.sync_team(t)
        elif k == "m":
            self.app.choose_subject(t, force=True)
            self.app.load(self)
        elif k == "x":
            self.sh.cfg["hidden"].append(t["url"])
            save_config(self.sh.cfg)
            self.app.load(self)
            self.app.msg = f"hidden {t['name']} (X shows hidden teams again)"
        else:
            return False
        return True


class TeamView(View):
    hint = "Enter open · d download · a download all files · m Moodle folder"

    def __init__(self, app, t):
        super().__init__(app, t.get("name") or t["title"])
        self.t = t

    def build(self):
        rows = []
        subj = self.sh.cfg["subjects"].get(self.t["url"])
        rows.append(Row(f"  {bold(magenta(self.t.get('name') or self.t['title']))}   {dim(self.t['title'])}"))
        rows.append(Row("  " + (green(f"↓ {moodle_root(self.sh.cfg) / subj}") if subj else
                                dim("↓ no ~/Moodle folder yet (asked on the first download, or press m)"))))
        rows.append(Row(""))
        rows.append(head("Channels"))
        web = self.sh.web
        if not self.t["channel"]:
            try:
                self.app.status(dim("loading Teams (the first time takes up to a minute)…"))
                web.boot()
                wt = web.match(self.t["title"])
                if wt:
                    for ch in wt["channels"]:
                        rows.append(Row(f"  {cyan('#')} {ch['name']}", {"channel": ch, "wt": wt}))
                else:
                    rows.append(Row(dim("  this team was not found in the Teams web app")))
            except TeamsError as e:
                rows += wrap_rows(dim(str(e)), "  ")
        forms = self.sh.web.team_forms(self.t) if self.sh.web.ready else {}
        if forms:
            rows += [Row(""), head(f"Forms seen in this team ({len(forms)})")]
            rows += [form_row(u, f, "  ") for u, f in forms.items()]
        rows += [Row(""), head("Files")]
        libs = self.sh.sp.libraries(self.t["url"])
        if libs:
            self.lib = libs[0]
            for it in self.sh.sp.folder(self.t["url"], self.lib["path"]):
                rows.append(item_row(it))
            for l in libs[1:]:
                r = item_row({"kind": "dir", "name": l["title"], "path": l["path"], "count": l["count"]})
                r.text = r.text.replace("▸", "▣", 1)
                rows.append(r)
        else:
            rows.append(Row(dim("  no files you can see")))
        return rows

    def enter(self, d):
        if "channel" in d:
            return self.app.push(ChannelView(self.app, self.t, d["wt"], d["channel"]))
        if "form" in d:
            return self.app.run_form(d["form"])
        if d.get("kind") == "dir":
            return self.app.push(FolderView(self.app, self.t, d))
        self.app.download(self.t, [d])

    def key(self, k, d):
        if k == "m":
            self.app.choose_subject(self.t, force=True)
            self.app.load(self)
        elif k == "a":
            self.app.sync_team(self.t)
        elif k == "d" and d and d.get("kind"):
            self.app.download(self.t, [d])
        elif k == "o" and d and d.get("kind") == "file":
            self.app.open_file(self.t, d)
        else:
            return False
        return True


FILE_ICONS = {"pdf": "▤", "doc": "¶", "docx": "¶", "odt": "¶", "ppt": "◧", "pptx": "◧", "xls": "▦", "xlsx": "▦",
              "csv": "▦", "png": "◩", "jpg": "◩", "jpeg": "◩", "gif": "◩", "zip": "◫", "mp4": "▶", "mp3": "♪"}


def item_row(it, indent="  "):
    nw = max(24, min(56, width() - 40))
    name = it["name"] if len(it["name"]) <= nw else it["name"][:nw - 1] + "…"
    pad = " " * (nw - len(name) + 2)
    if it["kind"] == "dir":
        return Row(f"{indent}{blue('▸')} {blue(bold(name))}{pad}{dim(str(it['count']) + ' items')}", it)
    icon = FILE_ICONS.get(it["name"].rsplit(".", 1)[-1].lower() if "." in it["name"] else "", "·")
    return Row(f"{indent}{dim(icon)} {name}{pad}{dim(human(it['size']).rjust(9) + '   ' + fmt_time(it['modified']))}", it)


class FolderView(View):
    hint = "Enter open/download · o open file · d download · a download all here"

    def __init__(self, app, t, folder):
        super().__init__(app, folder["name"])
        self.t, self.f = t, folder

    def build(self):
        return [item_row(it) for it in self.sh.sp.folder(self.t["url"], self.f["path"])] or \
            text_rows(dim("  (empty folder)"))

    def enter(self, it):
        if it["kind"] == "dir":
            return self.app.push(FolderView(self.app, self.t, it))
        self.app.download(self.t, [it])

    def key(self, k, it):
        if k == "a":
            self.app.download(self.t, [r.data for r in self.all_rows if r.data])
        elif k == "d" and it:
            self.app.download(self.t, [it])
        elif k == "o" and it and it["kind"] == "file":
            self.app.open_file(self.t, it)
        else:
            return False
        return True


class ChannelView(View):
    hint = "Enter read post / take Form · r reload"

    def __init__(self, app, t, wt, ch):
        super().__init__(app, ch["name"])
        self.t, self.wt, self.ch = t, wt, ch

    def build(self):
        self.app.status(dim("loading posts…"))
        posts = self.sh.web.channel(self.wt, self.ch)
        rows = []
        for p in posts:
            n = len(p["replies"])
            rows.append(Row(f"  {bold(p['author'] or '?')}  {dim(fmt_time(p['time']))}"
                            + (dim(f"  · {n} repl{'y' if n == 1 else 'ies'}") if n else ""), {"post": p}))
            if p["subject"]:
                rows.append(Row("    " + bold(p["subject"]), sel=False))
            rows += wrap_rows(p["text"], "    ", limit=5)
            for m in [p] + p["replies"]:
                rows += [form_row(u, {"label": (m["subject"] or m["text"].split("\n")[0])[:50]}) for u in m["forms"]]
            rows.append(Row(""))
        return rows or text_rows(dim("  no posts loaded (the channel is empty, or Teams did not show them; try r)"))

    def enter(self, d):
        if "form" in d:
            return self.app.run_form(d["form"])
        p = d["post"]
        self.app.push(PostView(self.app, p))


class PostView(View):
    hint = "Enter on ▶ Form: take the test · Enter on a link: open in browser"

    def __init__(self, app, p):
        super().__init__(app, (p["subject"] or p["text"].split("\n")[0])[:40] or "Post")
        self.p = p

    def build(self):
        rows = []
        for i, m in enumerate([self.p] + self.p["replies"]):
            if i == 1:
                rows += [Row(""), head(f"Replies ({len(self.p['replies'])})")]
            rows.append(Row(f"  {bold(m['author'] or '?')}  {dim(fmt_time(m['time']))}"))
            if m["subject"]:
                rows.append(Row("    " + bold(m["subject"])))
            rows += wrap_rows(m["text"], "    ")
            for u in m["forms"]:
                rows.append(form_row(u, {"label": "Microsoft Form"}))
            for j, u in enumerate(m["links"], 1):
                if u not in m["forms"]:
                    rows.append(Row(f"    {dim(f'[{j}]')} {blue(u[:90])}", {"link": u}))
            rows.append(Row(""))
        return rows

    def enter(self, d):
        if "form" in d:
            return self.app.run_form(d["form"])
        subprocess.Popen(["xdg-open", d["link"]], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                         start_new_session=True)
        self.app.msg = green("opened in browser")


class FormsView(View):
    hint = "Enter take the test · p paste a Forms link"

    def __init__(self, app):
        super().__init__(app, "Forms")

    def build(self):
        rows = [Row(f"  {green('+')} paste a Forms link from Teams", {"paste": True}), Row("")]
        forms = self.sh.web.forms
        if forms:
            rows.append(head(f"Forms found in Teams ({len(forms)})"))
            for u, f in sorted(forms.items(), key=lambda x: x[1].get("time") or "", reverse=True):
                rows.append(form_row(u, f, "  "))
                rows.append(Row(dim(f"      {f['where']}  {fmt_time(f.get('time'))}")))
        else:
            rows += text_rows(dim("  Forms links appear here when you open a channel or the Assignments tab.\n"
                                  "  You can also copy a link from Teams and paste it with p."))
        return rows

    def enter(self, d):
        if d.get("paste"):
            return self.paste()
        self.app.run_form(d["form"])

    def paste(self):
        s = self.app.prompt("Forms link: ")
        if s:
            links = forms_links(s) or [s.strip()]
            self.app.run_form(links[0])

    def key(self, k, d):
        if k == "p":
            self.paste()
            return True
        return False


def load_assign_cache():
    try:
        c = json.loads(ASSIGN_CACHE.read_text())
        return c if isinstance(c.get("items"), list) else None
    except (OSError, ValueError, AttributeError):
        return None


def save_assign_cache(items):
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    tmp = ASSIGN_CACHE.with_suffix(".tmp")
    tmp.write_text(json.dumps({"time": int(time.time()), "items": items}, ensure_ascii=False))
    os.chmod(tmp, 0o600)
    tmp.replace(ASSIGN_CACHE)


def rebucket(items):
    """Due dates pass while the cache sits on disk: work out to do / missed again from the clock."""
    now = datetime.now().astimezone()
    for a in items:
        if a.get("done"):
            a["bucket"] = "done"
            continue
        try:
            late = bool(a.get("due")) and datetime.fromisoformat(a["due"].replace("Z", "+00:00")) < now
        except ValueError:
            late = False
        a["bucket"] = "missed" if late else "due"
    return items


def diff_assignments(old, new):
    """Ids that are new, and ids whose status or points changed."""
    before = {a["id"]: a for a in old}
    added = {a["id"] for a in new if a["id"] not in before}
    changed = {a["id"] for a in new if a["id"] in before and
               (before[a["id"]].get("status"), before[a["id"]].get("points"), before[a["id"]].get("due"))
               != (a.get("status"), a.get("points"), a.get("due"))}
    return added, changed


def ago(ts):
    m = int((time.time() - ts) // 60)
    return "just now" if m < 1 else f"{m} min ago" if m < 60 else f"{m // 60} h ago" if m < 1440 else f"{m // 1440} days ago"


def fmt_due(iso):
    """'in 3 days', 'tomorrow 23:59', '2 days ago' next to the date."""
    if not iso:
        return "no due date"
    try:
        t = datetime.fromisoformat(iso.replace("Z", "+00:00")).astimezone()
    except ValueError:
        return iso
    days = (t.date() - datetime.now().astimezone().date()).days
    rel = ("today" if days == 0 else "tomorrow" if days == 1 else "yesterday" if days == -1 else
           f"in {days} days" if days > 0 else f"{-days} days ago")
    return f"{t.strftime('%a %d %b %H:%M')}  ·  {rel}"


def score_text(a):
    if a.get("points") is None:
        return ""
    p = f"{a['points']:g}" if isinstance(a["points"], (int, float)) else str(a["points"])
    return f"{p}/{a['max']:g} points" if isinstance(a.get("max"), (int, float)) else f"{p} points"


class AssignmentsView(View):
    hint = "Enter details · o open in Teams · D show done · r reload · / filter"
    BUCKETS = (("due", "To do"), ("missed", "Missed"), ("done", "Done"))

    def __init__(self, app):
        super().__init__(app, "Assignments")
        self.show_done = False
        self.items = None
        self.updated = None       # time of the last download from Teams
        self.pending = False      # True: cached list is on screen, check Teams after drawing it
        self.added, self.changed = set(), set()

    def update(self):
        """Download the list again and mark what is new or changed since the cached copy."""
        self.app.status(dim("checking Teams for new assignments…" if self.items else "loading your assignments…"))
        new = self.sh.web.assignments()
        if self.items:
            self.added, self.changed = diff_assignments(self.items, new)
        self.items, self.updated = new, time.time()
        save_assign_cache(new)
        n, c = len(self.added), len(self.changed)
        return green(f"✓ {n} new, {c} changed") if n or c else dim("✓ up to date")

    def check(self):
        """Run by the main loop right after the cached list was drawn."""
        self.pending = False
        msg = self.update()
        self.all_rows = self.build()
        self.apply_filter(self.sel)
        self.app.msg = msg

    def build(self):
        if self.items is None and not self.fresh:
            c = load_assign_cache()
            if c:
                self.items, self.updated, self.pending = c["items"], c["time"], True
        if self.items is None or self.fresh:
            self.app.msg = self.update() if self.items else ""
        self.fresh = False
        rebucket(self.items)
        groups = {k: [a for a in self.items if a["bucket"] == k] for k, _ in self.BUCKETS}
        groups["missed"].reverse()                      # most recent first
        groups["done"].sort(key=lambda a: a["due"] or "", reverse=True)
        rows = [Row(dim(f"  {len(self.items)} assignments · updated {ago(self.updated)}"
                        + ("  · checking for new ones…" if self.pending else "  · r checks again"))), Row("")]
        for k, label in self.BUCKETS:
            items = groups[k]
            if len(rows) > 2:
                rows.append(Row(""))
            rows.append(head(f"{label}  ·  {len(items)}"))
            if k == "done" and not self.show_done:
                if items:
                    rows.append(Row(dim(f"  + {len(items)} done  (D shows them)"), {"toggle": True}))
                continue
            if not items:
                rows.append(Row(dim("  nothing here" if k != "due" else "  nothing to do right now")))
            for a in items:
                rows.append(self.row(a))
        return rows

    def row(self, a):
        mark = {"due": yellow("○"), "missed": red("✗"), "done": green("✓")}[a["bucket"]]
        when = (red if a["bucket"] == "missed" else dim)(fmt_due(a["due"]))
        extra = (magenta(bold("  new")) if a["id"] in self.added else yellow("  changed") if a["id"] in self.changed else "")
        extra += (yellow("  ▶ Form") if a["forms"] else "") + (f"  {green(score_text(a))}" if score_text(a) else "")
        title = a["title"].strip()
        title = title if len(title) <= 46 else title[:45] + "…"
        cls = cyan(a["class"][:16].ljust(16)) if a["class"] else " " * 16
        return Row(f"  {mark} {bold(title.ljust(46))}  {cls}  {when}{extra}", a)

    def load(self):
        self.fresh = getattr(self, "fresh", False)
        super().load()

    def toggle(self):
        self.show_done = not self.show_done
        self.all_rows = self.build()
        self.apply_filter(self.sel)

    def key(self, k, a):
        if k == "D":
            self.toggle()
            return True
        if k in ("r", curses.KEY_F5):
            self.fresh = True
            self.app.load(self)
            return True
        if k == "o" and a and a.get("web"):
            subprocess.Popen(["xdg-open", a["web"]], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                             start_new_session=True)
            self.app.msg = green("opened in Teams in your browser")
            return True
        return False

    def enter(self, a):
        if a.get("toggle"):
            return self.toggle()

        def b():
            label = {"due": yellow("to do"), "missed": red("missed"), "done": green("done")}[a["bucket"]]
            rows = [Row(f"  {bold(magenta(a['title']))}"),
                    Row(f"  {label}  ·  {dim(a['class'] or 'class unknown')}  ·  {dim('due ' + fmt_due(a['due']))}"
                        + (f"  ·  {green(score_text(a))}" if score_text(a) else "")), Row("")]
            rows += wrap_rows(a["text"] or dim("(no instructions)"), "  ")
            if a["forms"]:
                rows += [Row(""), head("Forms test")] + [form_row(u, {"label": a["title"]}, "  ") for u in a["forms"]]
            if a.get("web"):
                rows += [Row(""), Row(f"  {blue('↗ open in Teams')}", {"link": a["web"]})]
            return rows

        def go(d):
            if "form" in d:
                return self.app.run_form(d["form"])
            subprocess.Popen(["xdg-open", d["link"]], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                             start_new_session=True)
            self.app.msg = green("opened in Teams in your browser")
        v = TextView(self.app, a["title"][:40], b, "Enter on ▶ Form: take the test")
        v.enter = go
        self.app.push(v)


class TUI:
    TABS = [("Courses", CoursesView), ("Forms", FormsView), ("Assignments", AssignmentsView)]

    def __init__(self, sh):
        self.sh = sh
        self.stacks = [None] * len(self.TABS)
        self.tab = 0
        self.msg = ""
        self.tab_pos = []
        self.h, self.w = 24, 80

    @property
    def view(self):
        return self.stacks[self.tab][-1]

    def body_h(self):
        return max(1, self.h - 4)

    def run(self):
        import locale
        locale.setlocale(locale.LC_ALL, "")
        os.environ.setdefault("ESCDELAY", "25")
        curses.wrapper(self.main)

    def main(self, scr):
        global interactive, IN_TUI
        self.scr = scr
        IN_TUI = True
        interactive = lambda fn: self.suspended(fn, pause=False)
        try:
            curses.curs_set(0)
        except curses.error:
            pass
        scr.keypad(True)
        try:
            curses.use_default_colors()
            for i in range(1, 8):
                curses.init_pair(i, i, -1)
        except curses.error:
            pass

        def cp(n):
            try:
                return curses.color_pair(n) if curses.has_colors() else 0
            except curses.error:
                return 0
        R, B, D = curses.A_REVERSE, curses.A_BOLD, curses.A_DIM
        self.attrs = {"brand": cp(5) | R | B, "tab_on": cp(6) | R | B, "tab_off": D, "dim": D, "bold": B,
                      "accent": cp(6) | B, "rule": cp(6) | D, "sel": cp(5) | R | B, "key": cp(6) | B}
        try:
            curses.mousemask(curses.ALL_MOUSE_EVENTS)
            curses.mouseinterval(0)
        except curses.error:
            pass
        self.h, self.w = scr.getmaxyx()
        self.switch(0)
        while True:
            self.draw()
            v = self.view
            if getattr(v, "pending", False):   # cached list is on screen: now ask Teams for changes
                try:
                    v.check()
                except Exception as e:
                    v.pending = False
                    self.msg = red("could not check for new assignments: " + str(e).split("\n")[0])
                continue
            try:
                k = scr.get_wch()
            except KeyboardInterrupt:
                self.msg = yellow("press q to quit")
                continue
            except curses.error:
                continue
            self.msg = ""
            try:
                if self.handle(k) == "quit":
                    break
            except TeamsError as e:
                self.msg = red("error: " + str(e).split("\n")[0])
            except KeyboardInterrupt:
                self.msg = yellow("interrupted")
            except Exception as e:
                self.msg = red(f"error: {type(e).__name__}: {str(e).splitlines()[0] if str(e) else ''}")

    # -- navigation --------------------------------------------------------- #
    def switch(self, i):
        self.tab = i % len(self.TABS)
        if self.stacks[self.tab] is None:
            v = self.TABS[self.tab][1](self)
            self.stacks[self.tab] = [v]
            self.load(v)
        elif self.tab == 1 and len(self.stacks[1]) == 1:
            self.load(self.stacks[1][0])     # Forms list grows while you browse

    def push(self, v):
        self.stacks[self.tab].append(v)
        self.load(v)

    def pop(self):
        if len(self.stacks[self.tab]) > 1:
            self.stacks[self.tab].pop()

    def load(self, v):
        loading = dim("loading…")
        self.status(loading)
        try:
            v.load()
            if self.msg == loading or self.msg.startswith(dim("loading")[:4]):
                self.msg = ""
        except Exception as e:
            v.all_rows = v.rows = wrap_rows(red("error: " + str(e)), "  ")
            v.sel, v.loaded = -1, True
            self.msg = red("error: " + str(e).split("\n")[0])

    def status(self, msg):
        self.msg = msg
        self.draw()

    # -- actions ------------------------------------------------------------ #
    def choose_subject(self, t, force=False):
        if force:
            self.sh.cfg["subjects"].pop(t["url"], None)
        if not self.sh.cfg["subjects"].get(t["url"]):
            self.suspended(lambda: self.sh.subject_for(t), pause=False)
        return self.sh.cfg["subjects"].get(t["url"])

    def download(self, t, items):
        subj = self.choose_subject(t)
        if not subj:
            self.msg = yellow("no Moodle folder chosen")
            return
        base = moodle_root(self.sh.cfg) / subj
        counts = {"new": 0, "updated": 0, "same": 0}
        for it in items:
            if it["kind"] == "file":
                self.status(dim(f"downloading {it['name']}…"))
                counts[self.sh.save_file(t, it, base / safe_name(it["name"]), quiet=True)] += 1
            else:
                for rel, f in self.sh.sp.walk(t["url"], it["path"]):
                    self.status(dim(f"downloading {rel}…"))
                    dest = base / safe_name(it["name"]) / Path(*[safe_name(p) for p in rel.split("/")])
                    counts[self.sh.save_file(t, f, dest, quiet=True)] += 1
        self.msg = green(f"✓ {counts['new']} new, {counts['updated']} updated, {counts['same']} already there → {base}")

    def sync_team(self, t):
        subj = self.choose_subject(t)
        if not subj:
            return
        base = moodle_root(self.sh.cfg) / subj
        counts = {"new": 0, "updated": 0, "same": 0}
        for lib in self.sh.sp.libraries(t["url"]):
            for rel, f in self.sh.sp.walk(t["url"], lib["path"]):
                self.status(dim(f"{t['title']}: {rel}"))
                parts = [safe_name(p) for p in rel.split("/")]
                counts[self.sh.save_file(t, f, base.joinpath(*(parts[1:] if len(parts) > 1 else parts)), quiet=True)] += 1
        self.msg = green(f"✓ {t['title']}: {counts['new']} new, {counts['updated']} updated, "
                         f"{counts['same']} unchanged → {base}")

    def open_file(self, t, it):
        import tempfile
        self.status(dim(f"downloading {it['name']}…"))
        p = Path(tempfile.mkdtemp(prefix="teams-")) / safe_name(it["name"])
        p.write_bytes(self.sh.sp.download(t["url"], it["path"]))
        if (os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY")) and shutil.which("xdg-open"):
            subprocess.Popen(["xdg-open", str(p)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                             start_new_session=True)
            self.msg = green(f"opened {p.name}")
        elif p.suffix.lower() == ".pdf" and shutil.which("pdftotext") and shutil.which("less"):
            def run():
                pr = subprocess.Popen(["pdftotext", "-layout", str(p), "-"], stdout=subprocess.PIPE)
                subprocess.call(["less", "-R"], stdin=pr.stdout)
                pr.wait()
            self.suspended(run, pause=False)
        else:
            self.msg = yellow(f"saved to {p}")

    def run_form(self, url):
        def go():
            fr = FormRunner(self.sh.b, url)
            try:
                fr.run()
            finally:
                fr.close()
        self.suspended(go)

    # -- input -------------------------------------------------------------- #
    def handle(self, k):
        v = self.view
        d = v.current
        K = curses
        if k == K.KEY_RESIZE:
            curses.update_lines_cols()
            return
        if k == K.KEY_MOUSE:
            return self.mouse()
        if v.key(k, d):
            return
        if k in ("q", "Q"):
            return "quit"
        if isinstance(k, str) and len(k) == 1 and k in "123"[:len(self.TABS)]:
            return self.switch(int(k) - 1)
        if k in (K.KEY_UP, "k"):
            v.move(-1)
        elif k in (K.KEY_DOWN, "j"):
            v.move(1)
        elif k in (K.KEY_NPAGE, " ", "\x04"):
            v.scroll(self.body_h() - 1)
        elif k in (K.KEY_PPAGE, "\x15"):
            v.scroll(-(self.body_h() - 1))
        elif k in (K.KEY_HOME, "g"):
            v.home()
        elif k in (K.KEY_END, "G"):
            v.end()
        elif k == "\t":
            self.switch(self.tab + 1)
        elif k == K.KEY_BTAB:
            self.switch(self.tab - 1)
        elif k in ("\n", "\r", K.KEY_ENTER, K.KEY_RIGHT, "l"):
            if d is not None:
                v.enter(d)
        elif k in (K.KEY_LEFT, "h", K.KEY_BACKSPACE, "\x7f", "\x08", "\x1b"):
            if v.filter:
                v.filter = ""
                v.apply_filter()
            else:
                self.pop()
        elif k == "/":
            s = self.prompt("/", v.filter)
            if s is not None:
                v.filter = s.strip()
                v.apply_filter()
        elif k in ("r", K.KEY_F5):
            self.load(v)
        elif k == "?":
            self.push(HelpView(self))

    def mouse(self):
        try:
            _, mx, my, _, b = curses.getmouse()
        except curses.error:
            return
        v = self.view
        if b & getattr(curses, "BUTTON4_PRESSED", 0):
            return v.scroll(-3)
        if b & getattr(curses, "BUTTON5_PRESSED", 0x200000):
            return v.scroll(3)
        if not b & (curses.BUTTON1_CLICKED | curses.BUTTON1_PRESSED | curses.BUTTON1_DOUBLE_CLICKED):
            return
        if my == 0:
            for i, (x0, x1) in enumerate(self.tab_pos):
                if x0 <= mx < x1:
                    return self.switch(i)
            return
        if my == 1:
            return self.pop()
        idx = v.top + my - 3
        if 3 <= my < 3 + self.body_h() and idx < len(v.rows) and v.rows[idx].sel:
            if idx == v.sel or b & curses.BUTTON1_DOUBLE_CLICKED:
                v.enter(v.current)
            else:
                v.sel = idx

    def prompt(self, label, initial=""):
        buf = list(initial or "")
        try:
            curses.curs_set(1)
        except curses.error:
            pass
        try:
            while True:
                self.h, self.w = self.scr.getmaxyx()
                shown = (label + "".join(buf))[-(self.w - 2):]
                self.scr.move(self.h - 1, 0)
                self.scr.clrtoeol()
                self.put(self.h - 1, 0, shown, curses.A_BOLD)
                self.scr.move(self.h - 1, min(self.w - 2, sum(char_w(c) for c in shown)))
                self.scr.refresh()
                ch = self.scr.get_wch()
                if ch in ("\n", "\r", curses.KEY_ENTER):
                    return "".join(buf)
                if ch == "\x1b":
                    return None
                if ch in (curses.KEY_BACKSPACE, "\x7f", "\x08"):
                    if buf:
                        buf.pop()
                elif ch == "\x15":
                    buf.clear()
                elif isinstance(ch, str) and ch.isprintable():
                    buf.append(ch)
        except KeyboardInterrupt:
            return None
        finally:
            try:
                curses.curs_set(0)
            except curses.error:
                pass

    def suspended(self, fn, pause=True):
        """Leave curses for a moment (Forms tests, sign-in, folder questions)."""
        curses.def_prog_mode()
        curses.endwin()
        print()
        result = None
        try:
            result = fn()
        except TeamsError as e:
            err(e)
            pause = True
        except (KeyboardInterrupt, EOFError):
            print("\ncancelled.")
        if pause:
            try:
                print(dim("\n[Enter] back to teams "), end="", flush=True)
                input()
            except (EOFError, KeyboardInterrupt):
                pass
        curses.reset_prog_mode()
        self.scr.clear()
        self.scr.refresh()
        return result

    # -- drawing ------------------------------------------------------------ #
    def put(self, y, x, text, attr=0):
        if y >= self.h or x >= self.w - 1:
            return x
        out, col = [], x
        for ch in text:
            if not ch.isprintable():
                continue
            cw = char_w(ch)
            if col + cw > self.w - 1:
                break
            out.append(ch)
            col += cw
        try:
            self.scr.addstr(y, x, "".join(out), attr)
        except curses.error:
            pass
        return col

    def put_ansi(self, y, x, text, extra=0, fill=False):
        """Draw text with ANSI colours; returns the column after it."""
        attr, pos, col = 0, 0, x
        segs = []
        for m in ANSI_RE.finditer(text):
            if m.start() > pos:
                segs.append((text[pos:m.start()], attr))
            for code in (m.group(1) or "0").split(";"):
                code = int(code or 0)
                if code == 0:
                    attr = 0
                elif code == 1:
                    attr |= curses.A_BOLD
                elif code == 2:
                    attr |= curses.A_DIM
                elif 30 <= code <= 37:
                    attr = (attr & ~curses.A_COLOR) | curses.color_pair(code - 30)
            pos = m.end()
        segs.append((text[pos:], attr))
        for seg, a in segs:
            if seg:
                col = self.put(y, col, seg, a | extra)
        if fill and col < self.w - 1:
            self.put(y, col, " " * (self.w - 1 - col), extra)
        return col

    def draw(self):
        scr, A = self.scr, self.attrs
        self.h, self.w = scr.getmaxyx()
        scr.erase()
        # top bar: brand, tabs, school year
        x = self.put(0, 0, " ◆ teams ", A["brand"]) + 1
        self.tab_pos = []
        for i, (name, _) in enumerate(self.TABS):
            nx = self.put(0, x, f" {i + 1} {name} ", A["tab_on"] if i == self.tab else A["tab_off"])
            self.tab_pos.append((x, nx))
            x = nx + 1
        right = f" school year {school_label()} "
        if x + len(right) < self.w - 1:
            self.put(0, self.w - 1 - len(right), right, A["dim"])
        if self.stacks[self.tab] is None:
            scr.refresh()
            return
        v = self.view
        # breadcrumb and rule
        stack = self.stacks[self.tab]
        x = 1
        for j, s in enumerate(stack):
            last = j == len(stack) - 1
            x = self.put(1, x, s.title, A["bold"] if last else A["dim"])
            if not last:
                x = self.put(1, x, "  ›  ", A["dim"])
        if v.filter:
            self.put(1, x + 3, f"/ {v.filter}", A["accent"])
        self.put(2, 0, "─" * (self.w - 1), A["rule"])
        # body
        v.clamp()
        bh = self.body_h()
        for i in range(bh):
            idx = v.top + i
            if idx >= len(v.rows):
                break
            r, y = v.rows[idx], 3 + i
            if idx == v.sel:
                t = strip_ansi(r.text)
                t = ("▸ " + t[2:]) if t.startswith("  ") else t
                self.put(y, 0, t + " " * max(0, self.w - 1 - sum(char_w(c) for c in t)), A["sel"])
            elif r.kind == "head":
                col = self.put_ansi(y, 0, r.text)
                if col + 2 < self.w - 1:
                    self.put(y, col + 1, "─" * (self.w - 2 - col), A["rule"])
            else:
                self.put_ansi(y, 0, r.text)
        # bottom line: message or key hints, position
        n = len(v.rows)
        right = f" {v.top + 1}-{min(n, v.top + bh)} of {n} " if n > bh else ""
        if self.msg:
            self.put_ansi(self.h - 1, 1, self.msg)
        else:
            hint = f"{v.hint}{' · ' if v.hint else ''}← back · ? help · q quit"
            x, stop = 1, self.w - 2 - len(right)
            for part in hint.split(" · "):
                k, _, label = part.partition(" ")
                if x + len(part) + 3 > stop:
                    break
                x = self.put(self.h - 1, x, k, A["key"])
                x = self.put(self.h - 1, x + 1, label, A["dim"]) + 3
        if right and len(right) < self.w - 2:
            self.put(self.h - 1, self.w - 1 - len(right), right, A["dim"])
        scr.refresh()


def run_tui(sh):
    if curses is None:
        raise TeamsError("Python curses module not available")
    if not sys.stdout.isatty():
        raise TeamsError("the full-screen interface needs a terminal")
    if not sh.cfg.get("host"):
        print("Not signed in yet — let's do that first.\n")
        sh.do_login("")
        if not sh.cfg.get("host"):
            return
    sh.need()
    TUI(sh).run()


def main():
    argv = sys.argv[1:]
    if argv and argv[0] in ("-h", "--help", "help"):
        print(HELP.strip())
        return 0
    if argv and argv[0] in ("-V", "--version"):
        print(f"{APP} {VERSION}")
        return 0
    b = Browser()
    sh = Shell(b)
    try:
        if not argv or argv[0] == "tui":
            if sys.stdout.isatty() and curses is not None:
                try:
                    run_tui(sh)
                except TeamsError as e:
                    err(e)
                return 0
            argv = ["shell"]
        if argv[0] != "shell":
            sh.onecmd(" ".join(shlex.quote(a) for a in argv))
            return 0
        try:
            import readline
            STATE_DIR.mkdir(parents=True, exist_ok=True)
            try:
                readline.read_history_file(HISTORY_FILE)
            except OSError:
                pass
        except ImportError:
            readline = None
        print(f"{bold('teams')} {dim(VERSION)}  —  type {cyan('help')} for commands, {cyan('tui')} for the full-screen view")
        if not sh.cfg.get("host"):
            print(yellow("Not signed in yet: type `login`."))
        try:
            sh.cmdloop()
        finally:
            if readline:
                try:
                    readline.write_history_file(HISTORY_FILE)
                except OSError:
                    pass
    except KeyboardInterrupt:
        print()
    finally:
        b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
