#!/usr/bin/env python3
"""teams-tui: a read-only Microsoft Teams client for the terminal.

Views your chats, team channels and class assignments through the official
Microsoft Graph API, using delegated permissions only (you see what you could
already see in Teams, nothing more).

    teams_tui.py login --client-id <APP_ID>   # first time, see README.md
    teams_tui.py                              # start the TUI
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import threading
import time
from datetime import datetime, timezone
from html.parser import HTMLParser
from pathlib import Path

import httpx
import msal
from rich.text import Text
from textual import work
from textual.app import App, ComposeResult
from textual.binding import Binding
from textual.containers import Horizontal
from textual.widgets import (
    Footer,
    Header,
    Label,
    ListItem,
    ListView,
    RichLog,
    TabbedContent,
    TabPane,
    Tree,
)
from textual.worker import get_current_worker

APP_NAME = "teams-tui"
CONFIG_DIR = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config")) / APP_NAME
CACHE_DIR = Path(os.environ.get("XDG_CACHE_HOME", Path.home() / ".cache")) / APP_NAME
CONFIG_FILE = CONFIG_DIR / "config.json"
TOKEN_FILE = CACHE_DIR / "token_cache.json"

GRAPH = "https://graph.microsoft.com"

# Delegated scopes. The first group can be consented to by a normal user
# (unless the tenant disables user consent entirely). The second group always
# needs tenant admin consent, so they are opt-in at login time.
BASE_SCOPES = ["User.Read", "Chat.Read", "Team.ReadBasic.All", "Channel.ReadBasic.All"]
CHANNEL_SCOPES = ["ChannelMessage.Read.All"]  # admin consent required
EDU_SCOPES = ["EduAssignments.ReadBasic"]  # admin consent required


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------
def load_config() -> dict:
    cfg = {"client_id": None, "tenant": "organizations"}
    if CONFIG_FILE.exists():
        try:
            cfg.update(json.loads(CONFIG_FILE.read_text()))
        except (OSError, ValueError):
            pass
    cfg["client_id"] = os.environ.get("TEAMS_TUI_CLIENT_ID") or cfg["client_id"]
    cfg["tenant"] = os.environ.get("TEAMS_TUI_TENANT") or cfg["tenant"]
    return cfg


def save_config(cfg: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2))


def parse_dt(s: str | None) -> datetime | None:
    if not s:
        return None
    s = re.sub(r"(\.\d{6})\d+", r"\1", s.replace("Z", "+00:00"))
    try:
        dt = datetime.fromisoformat(s)
    except ValueError:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone()


def fmt_time(s: str | None, short: bool = False) -> str:
    dt = parse_dt(s)
    if not dt:
        return ""
    if short:
        if dt.date() == datetime.now().astimezone().date():
            return dt.strftime("%H:%M")
        return dt.strftime("%d %b %H:%M")
    return dt.strftime("%Y-%m-%d %H:%M")


def relative(dt: datetime) -> str:
    secs = int((dt - datetime.now().astimezone()).total_seconds())
    past = secs < 0
    secs = abs(secs)
    days, rem = divmod(secs, 86400)
    hours = rem // 3600
    mins = (rem % 3600) // 60
    if days:
        txt = f"{days}d {hours}h"
    elif hours:
        txt = f"{hours}h {mins}m"
    else:
        txt = f"{mins}m"
    return f"overdue by {txt}" if past else f"due in {txt}"


_BLOCK = {"p", "div", "tr", "blockquote", "ul", "ol", "h1", "h2", "h3", "h4", "h5", "h6", "table"}


class _Html2Text(HTMLParser):
    """Tiny HTML -> plain text converter for Teams message bodies."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self._links: list[tuple[str, int]] = []
        self._skip = 0

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag in ("script", "style"):
            self._skip += 1
        elif tag == "br":
            self.parts.append("\n")
        elif tag == "li":
            self.parts.append("\n• ")
        elif tag in _BLOCK:
            self.parts.append("\n")
        elif tag == "img":
            alt = a.get("alt")
            self.parts.append(f"[image: {alt}]" if alt else "[image]")
        elif tag == "emoji":
            self.parts.append(a.get("alt") or "")
        elif tag == "a":
            self._links.append((a.get("href") or "", len(self.parts)))

    def handle_endtag(self, tag):
        if tag in ("script", "style"):
            self._skip = max(0, self._skip - 1)
        elif tag in _BLOCK:
            self.parts.append("\n")
        elif tag == "a" and self._links:
            href, start = self._links.pop()
            label = "".join(self.parts[start:]).strip()
            if href and not href.startswith("#") and href.strip() != label:
                self.parts.append(f" <{href}>")

    def handle_data(self, data):
        if not self._skip:
            self.parts.append(data)


def html_to_text(html: str) -> str:
    p = _Html2Text()
    try:
        p.feed(html)
        p.close()
    except Exception:  # malformed HTML: fall back to a crude strip
        return re.sub(r"<[^>]+>", "", html)
    out = "".join(p.parts).replace("\xa0", " ")
    out = re.sub(r"[ \t]+\n", "\n", out)
    return re.sub(r"\n{3,}", "\n\n", out).strip()


def body_text(body: dict | None) -> str:
    body = body or {}
    content = body.get("content") or ""
    return html_to_text(content) if body.get("contentType") == "html" else content.strip()


# --------------------------------------------------------------------------
# Auth + Graph client
# --------------------------------------------------------------------------
class GraphError(Exception):
    def __init__(self, status: int, message: str):
        super().__init__(f"{status}: {message}")
        self.status = status
        self.message = message


class Auth:
    def __init__(self, cfg: dict):
        if not cfg.get("client_id"):
            sys.exit("No client id configured. Run:  teams_tui.py login --client-id <APP_ID>  (see README.md)")
        self.cache = msal.SerializableTokenCache()
        if TOKEN_FILE.exists():
            self.cache.deserialize(TOKEN_FILE.read_text())
        self.app = msal.PublicClientApplication(
            cfg["client_id"],
            authority=f"https://login.microsoftonline.com/{cfg['tenant']}",
            token_cache=self.cache,
        )
        self._lock = threading.Lock()

    def _save(self) -> None:
        if self.cache.has_state_changed:
            CACHE_DIR.mkdir(parents=True, exist_ok=True, mode=0o700)
            fd = os.open(TOKEN_FILE, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
            with os.fdopen(fd, "w") as f:
                f.write(self.cache.serialize())

    def login(self, scopes: list[str], device_code: bool = False) -> dict:
        if device_code:
            flow = self.app.initiate_device_flow(scopes=scopes)
            if "user_code" not in flow:
                return {"error": "device_flow", "error_description": json.dumps(flow)}
            print(flow["message"], flush=True)
            result = self.app.acquire_token_by_device_flow(flow)
        else:
            result = self.app.acquire_token_interactive(scopes, prompt="select_account")
        self._save()
        return result

    def token(self, scopes: list[str]) -> str | None:
        with self._lock:
            accounts = self.app.get_accounts()
            if not accounts:
                return None
            result = self.app.acquire_token_silent(scopes, account=accounts[0])
            self._save()
        return result.get("access_token") if result else None

    def logout(self) -> None:
        TOKEN_FILE.unlink(missing_ok=True)


class Api:
    def __init__(self, auth: Auth):
        self.auth = auth
        self.http = httpx.Client(timeout=30.0, headers={"Accept": "application/json"})
        self.me_id: str | None = None
        self._team_names: dict[str, str] = {}

    def _get(self, path: str, scopes: list[str], params: dict | None = None, version: str = "v1.0") -> dict:
        token = self.auth.token(scopes)
        if not token:
            raise GraphError(0, f"Not signed in, or permission not granted: {', '.join(scopes)}")
        url = path if path.startswith("http") else f"{GRAPH}/{version}{path}"
        r = None
        for _ in range(3):
            r = self.http.get(url, params=params, headers={"Authorization": f"Bearer {token}"})
            if r.status_code == 429:  # throttled
                time.sleep(min(int(r.headers.get("Retry-After", "2")), 30))
                continue
            break
        assert r is not None
        if r.status_code >= 400:
            try:
                msg = r.json()["error"]["message"]
            except Exception:
                msg = r.text[:300]
            raise GraphError(r.status_code, msg)
        return r.json()

    def _paged(self, path: str, scopes: list[str], params: dict | None = None, max_items: int = 150) -> list[dict]:
        items: list[dict] = []
        url: str | None = path
        while url and len(items) < max_items:
            data = self._get(url, scopes, params)
            params = None  # nextLink already carries the query
            items += data.get("value", [])
            url = data.get("@odata.nextLink")
        return items

    # ---- profile --------------------------------------------------------
    def me(self) -> dict:
        d = self._get("/me", ["User.Read"], {"$select": "id,displayName,userPrincipalName"})
        self.me_id = d["id"]
        return d

    # ---- chats ----------------------------------------------------------
    def chats(self) -> list[dict]:
        scopes = ["Chat.Read"]
        try:
            items = self._paged(
                "/me/chats",
                scopes,
                {"$expand": "members,lastMessagePreview", "$top": "50"},
                max_items=100,
            )
        except GraphError as e:
            if e.status != 400:
                raise
            items = self._paged("/me/chats", scopes, {"$expand": "members", "$top": "50"}, max_items=100)

        def last(c: dict) -> str:
            return (c.get("lastMessagePreview") or {}).get("createdDateTime") or c.get("lastUpdatedDateTime") or ""

        return sorted(items, key=last, reverse=True)

    def chat_title(self, chat: dict) -> str:
        if chat.get("topic"):
            return chat["topic"]
        others = [
            m.get("displayName")
            for m in chat.get("members", [])
            if m.get("userId") != self.me_id and m.get("displayName")
        ]
        if chat.get("chatType") == "oneOnOne" and others:
            return others[0]
        if others:
            return ", ".join(others[:4]) + ("…" if len(others) > 4 else "")
        return "(unnamed chat)"

    def chat_messages(self, chat_id: str) -> list[dict]:
        msgs = self._paged(f"/me/chats/{chat_id}/messages", ["Chat.Read"], {"$top": "50"}, max_items=150)
        return sorted(msgs, key=lambda m: m.get("createdDateTime") or "")

    # ---- teams / channels -----------------------------------------------
    def teams_with_channels(self) -> list[tuple[dict, list[dict] | GraphError]]:
        teams = self._paged("/me/joinedTeams", ["Team.ReadBasic.All"])
        self._team_names = {t["id"]: t.get("displayName", "?") for t in teams}
        out: list[tuple[dict, list[dict] | GraphError]] = []
        for t in sorted(teams, key=lambda t: (t.get("displayName") or "").lower()):
            try:
                chans = self._paged(f"/teams/{t['id']}/channels", ["Channel.ReadBasic.All"])
            except GraphError as e:
                out.append((t, e))
            else:
                out.append((t, chans))
        return out

    def channel_messages(self, team_id: str, channel_id: str) -> list[dict]:
        msgs = self._paged(
            f"/teams/{team_id}/channels/{channel_id}/messages", CHANNEL_SCOPES, {"$top": "30"}, max_items=30
        )
        return sorted(msgs, key=lambda m: m.get("createdDateTime") or "")

    def channel_replies(self, team_id: str, channel_id: str, msg_id: str) -> list[dict]:
        msgs = self._paged(
            f"/teams/{team_id}/channels/{channel_id}/messages/{msg_id}/replies", CHANNEL_SCOPES, {"$top": "50"}
        )
        return sorted(msgs, key=lambda m: m.get("createdDateTime") or "")

    # ---- assignments ----------------------------------------------------
    def team_name(self, class_id: str) -> str:
        if not self._team_names:
            try:
                for t in self._paged("/me/joinedTeams", ["Team.ReadBasic.All"]):
                    self._team_names[t["id"]] = t.get("displayName", "?")
            except GraphError:
                pass
        return self._team_names.get(class_id, "")  # class id == backing team id

    def assignments(self) -> list[dict]:
        return self._paged("/education/me/assignments", EDU_SCOPES, max_items=300)

    def my_submission(self, class_id: str, assignment_id: str) -> dict | None:
        d = self._get(f"/education/classes/{class_id}/assignments/{assignment_id}/submissions", EDU_SCOPES)
        subs = d.get("value", [])
        return subs[0] if subs else None


# --------------------------------------------------------------------------
# Message rendering
# --------------------------------------------------------------------------
def sender_name(m: dict) -> str:
    frm = m.get("from") or {}
    who = frm.get("user") or frm.get("application") or frm.get("device") or {}
    return who.get("displayName") or "system"


def write_messages(log: RichLog, title: str, msgs: list[dict]) -> None:
    log.clear()
    log.write(Text(title, style="bold underline"))
    last_day = None
    shown = 0
    for m in msgs:
        if m.get("messageType") not in (None, "message"):
            continue
        dt = parse_dt(m.get("createdDateTime"))
        if dt and dt.date() != last_day:
            last_day = dt.date()
            log.write(Text(f"\n── {dt.strftime('%A, %d %B %Y')} ──", style="dim"))
        header = Text.assemble(
            (sender_name(m), "bold cyan"),
            (f"  {dt.strftime('%H:%M') if dt else ''}", "dim"),
        )
        log.write(header)
        if m.get("deletedDateTime"):
            log.write(Text("  (deleted)", style="dim italic"))
        else:
            if m.get("subject"):
                log.write(Text("  " + m["subject"], style="bold"))
            body = body_text(m.get("body")) or "(no text)"
            log.write(Text("  " + body.replace("\n", "\n  ")))
            for att in m.get("attachments") or []:
                if att.get("name"):
                    log.write(Text(f"  📎 {att['name']}", style="yellow"))
        shown += 1
    if not shown:
        log.write(Text("\n(no messages)", style="dim"))
    log.scroll_end(animate=False)


# --------------------------------------------------------------------------
# TUI
# --------------------------------------------------------------------------
class DataItem(ListItem):
    def __init__(self, label: str, data: dict):
        super().__init__(Label(label, markup=False))
        self.data = data


CHANNEL_HINT = (
    "Reading channel posts needs the ChannelMessage.Read.All permission, and Microsoft\n"
    "requires a tenant admin to approve it. Chats, team and channel names still work.\n\n"
    "If your school approves the app, run:  teams_tui.py login --channels"
)
EDU_HINT = (
    "Reading assignments needs the EduAssignments.ReadBasic permission, and Microsoft\n"
    "requires a tenant admin to approve it for every app.\n\n"
    "If your school approves the app, run:  teams_tui.py login --assignments"
)


class TeamsTui(App):
    TITLE = "Teams"
    CSS = """
    TabPane { padding: 0; }
    .left { width: 46; border-right: solid $primary; }
    .viewer { padding: 0 1; }
    """
    BINDINGS = [
        Binding("1", "tab('chats')", "Chats"),
        Binding("2", "tab('teams')", "Teams"),
        Binding("3", "tab('assignments')", "Assignments"),
        Binding("r", "refresh", "Refresh"),
        Binding("a", "toggle_all", "All/recent assignments"),
        Binding("q", "quit", "Quit"),
    ]

    def __init__(self, api: Api):
        super().__init__()
        self.api = api
        self.show_all_assignments = False
        self._assignments: list[dict] = []

    def compose(self) -> ComposeResult:
        yield Header()
        with TabbedContent(initial="chats", id="tabs"):
            with TabPane("1 Chats", id="chats"):
                with Horizontal():
                    yield ListView(id="chat-list", classes="left")
                    yield RichLog(id="chat-view", classes="viewer", wrap=True, markup=False, highlight=False)
            with TabPane("2 Teams", id="teams"):
                with Horizontal():
                    yield Tree("Teams", id="team-tree", classes="left")
                    yield RichLog(id="chan-view", classes="viewer", wrap=True, markup=False, highlight=False)
            with TabPane("3 Assignments", id="assignments"):
                with Horizontal():
                    yield ListView(id="asg-list", classes="left")
                    yield RichLog(id="asg-view", classes="viewer", wrap=True, markup=False, highlight=False)
        yield Footer()

    def on_mount(self) -> None:
        self.query_one("#team-tree", Tree).show_root = False
        self.load_chats()
        self.load_teams()
        self.load_assignments()
        self.query_one("#chat-list").focus()

    # ---- actions --------------------------------------------------------
    def action_tab(self, name: str) -> None:
        self.query_one(TabbedContent).active = name
        focus = {"chats": "#chat-list", "teams": "#team-tree", "assignments": "#asg-list"}[name]
        self.query_one(focus).focus()

    def action_refresh(self) -> None:
        active = self.query_one(TabbedContent).active
        {"chats": self.load_chats, "teams": self.load_teams, "assignments": self.load_assignments}[active]()

    def action_toggle_all(self) -> None:
        self.show_all_assignments = not self.show_all_assignments
        self.fill_assignments()

    # ---- helpers --------------------------------------------------------
    def show_text(self, view: str, text: str, style: str = "") -> None:
        log = self.query_one(view, RichLog)
        log.clear()
        log.write(Text(text, style=style))

    def fail(self, view: str, err: GraphError, hint: str | None = None) -> None:
        msg = f"Error {err.status}: {err.message}"
        if hint and err.status in (0, 401, 403):
            msg = hint + "\n\n(" + msg + ")"
        self.show_text(view, msg, "red")

    # ---- chats ----------------------------------------------------------
    @work(thread=True, group="chats-load")
    def load_chats(self) -> None:
        self.call_from_thread(self.show_text, "#chat-view", "Loading chats…", "dim")
        try:
            me = self.api.me()
            chats = self.api.chats()
        except GraphError as e:
            self.call_from_thread(self.fail, "#chat-view", e)
            return
        self.call_from_thread(self.fill_chats, chats, me)

    def fill_chats(self, chats: list[dict], me: dict) -> None:
        self.sub_title = me.get("displayName", "")
        items = []
        for c in chats:
            title = self.api.chat_title(c)
            prev = c.get("lastMessagePreview") or {}
            when = fmt_time(prev.get("createdDateTime") or c.get("lastUpdatedDateTime"), short=True)
            snippet = body_text(prev.get("body")).replace("\n", " ")[:30]
            tag = {"oneOnOne": "1:1", "group": "grp", "meeting": "mtg"}.get(c.get("chatType"), "?")
            label = f"{tag}  {title}\n     {when}  {snippet}".rstrip()
            items.append(DataItem(label, {"id": c["id"], "title": title}))
        lv = self.query_one("#chat-list", ListView)
        lv.clear()
        lv.extend(items)
        self.show_text("#chat-view", f"{len(items)} chats. Pick one and press Enter.", "dim")

    @work(thread=True, exclusive=True, group="chat-open")
    def open_chat(self, chat: dict) -> None:
        worker = get_current_worker()
        self.call_from_thread(self.show_text, "#chat-view", "Loading…", "dim")
        try:
            msgs = self.api.chat_messages(chat["id"])
        except GraphError as e:
            if not worker.is_cancelled:
                self.call_from_thread(self.fail, "#chat-view", e)
            return
        if not worker.is_cancelled:
            self.call_from_thread(write_messages, self.query_one("#chat-view", RichLog), chat["title"], msgs)

    # ---- teams ----------------------------------------------------------
    @work(thread=True, group="teams-load")
    def load_teams(self) -> None:
        self.call_from_thread(self.show_text, "#chan-view", "Loading teams…", "dim")
        try:
            data = self.api.teams_with_channels()
        except GraphError as e:
            self.call_from_thread(self.fail, "#chan-view", e)
            return
        self.call_from_thread(self.fill_teams, data)

    def fill_teams(self, data) -> None:
        tree = self.query_one("#team-tree", Tree)
        tree.clear()
        for team, chans in data:
            node = tree.root.add(Text(team.get("displayName", "?")), data={"kind": "team"})
            if isinstance(chans, GraphError):
                node.add_leaf(Text(f"(channels unavailable: {chans.status})", style="red"))
                continue
            for ch in chans:
                node.add_leaf(Text(ch.get("displayName", "?")), data={"kind": "channel", "team": team, "channel": ch})
        tree.root.expand()
        self.show_text("#chan-view", f"{len(data)} teams. Expand one and press Enter on a channel.", "dim")

    def on_tree_node_selected(self, event: Tree.NodeSelected) -> None:
        d = event.node.data
        if d and d.get("kind") == "channel":
            self.open_channel(d["team"], d["channel"])

    @work(thread=True, exclusive=True, group="chan-open")
    def open_channel(self, team: dict, channel: dict) -> None:
        worker = get_current_worker()
        title = f"{team.get('displayName')} / {channel.get('displayName')}"
        self.call_from_thread(self.show_text, "#chan-view", "Loading…", "dim")
        try:
            posts = self.api.channel_messages(team["id"], channel["id"])
            for p in posts:  # include replies under each post
                if worker.is_cancelled:
                    return
                if p.get("messageType") == "message":
                    p["_replies"] = self.api.channel_replies(team["id"], channel["id"], p["id"])
        except GraphError as e:
            if not worker.is_cancelled:
                self.call_from_thread(self.fail, "#chan-view", e, CHANNEL_HINT)
            return
        flat: list[dict] = []
        for p in posts:
            flat.append(p)
            flat.extend(p.get("_replies", []))
        if not worker.is_cancelled:
            self.call_from_thread(write_messages, self.query_one("#chan-view", RichLog), title, flat)

    # ---- assignments ----------------------------------------------------
    @work(thread=True, group="asg-load")
    def load_assignments(self) -> None:
        self.call_from_thread(self.show_text, "#asg-view", "Loading assignments…", "dim")
        try:
            items = self.api.assignments()
        except GraphError as e:
            self.call_from_thread(self.fail, "#asg-view", e, EDU_HINT)
            return
        for a in items:
            a["_class"] = self.api.team_name(a.get("classId", ""))
        self._assignments = items
        self.call_from_thread(self.fill_assignments)

    def fill_assignments(self) -> None:
        now = datetime.now().astimezone()
        rows = []
        for a in self._assignments:
            due = parse_dt(a.get("dueDateTime"))
            if not self.show_all_assignments and due and (now - due).days > 7:
                continue
            rows.append((a, due))
        if self.show_all_assignments:
            rows.sort(key=lambda r: (r[1] is None, -(r[1].timestamp() if r[1] else 0)))
        else:
            rows.sort(key=lambda r: (r[1] is None, r[1].timestamp() if r[1] else 0))
        items = []
        for a, due in rows:
            when = due.strftime("%d %b %H:%M") if due else "no due date"
            cls = a.get("_class") or "class"
            items.append(DataItem(f"{when}  {a.get('displayName', '?')}\n     {cls}", a))
        lv = self.query_one("#asg-list", ListView)
        lv.clear()
        lv.extend(items)
        mode = "all" if self.show_all_assignments else "due within the last 7 days or later"
        self.show_text("#asg-view", f"{len(items)} assignments ({mode}). Press 'a' to toggle.", "dim")

    @work(thread=True, exclusive=True, group="asg-open")
    def open_assignment(self, a: dict) -> None:
        worker = get_current_worker()
        sub = None
        try:
            if a.get("classId") and a.get("id"):
                sub = self.api.my_submission(a["classId"], a["id"])
        except GraphError:
            pass  # submission info is best-effort
        if not worker.is_cancelled:
            self.call_from_thread(self.render_assignment, a, sub)

    def render_assignment(self, a: dict, sub: dict | None) -> None:
        log = self.query_one("#asg-view", RichLog)
        log.clear()
        log.write(Text(a.get("displayName", "?"), style="bold underline"))
        if a.get("_class"):
            log.write(Text(a["_class"], style="cyan"))
        due = parse_dt(a.get("dueDateTime"))
        if due:
            log.write(Text(f"Due: {due.strftime('%A %d %B %Y, %H:%M')}  ({relative(due)})", style="bold"))
        else:
            log.write(Text("Due: not set", style="dim"))
        if sub:
            line = f"Your status: {sub.get('status', '?')}"
            if sub.get("submittedDateTime"):
                line += f" (submitted {fmt_time(sub['submittedDateTime'])})"
            log.write(Text(line, style="green" if sub.get("status") in ("submitted", "returned") else "yellow"))
        grading = (a.get("grading") or {}).get("maxPoints")
        if grading:
            log.write(Text(f"Max points: {grading}"))
        if a.get("allowLateSubmissions") is not None:
            log.write(Text(f"Late submissions: {'allowed' if a['allowLateSubmissions'] else 'not allowed'}"))
        log.write(Text(""))
        log.write(Text(body_text(a.get("instructions")) or "(no instructions)"))
        if a.get("webUrl"):
            log.write(Text(f"\nOpen in Teams: {a['webUrl']}", style="blue"))
        log.scroll_home(animate=False)

    # ---- list selection -------------------------------------------------
    def on_list_view_selected(self, event: ListView.Selected) -> None:
        item = event.item
        if not isinstance(item, DataItem):
            return
        if event.list_view.id == "chat-list":
            self.open_chat(item.data)
        elif event.list_view.id == "asg-list":
            self.open_assignment(item.data)


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def cmd_login(args: argparse.Namespace) -> int:
    cfg = load_config()
    if args.client_id:
        cfg["client_id"] = args.client_id
    if args.tenant:
        cfg["tenant"] = args.tenant
    scopes = list(BASE_SCOPES)
    if args.channels:
        scopes += CHANNEL_SCOPES
    if args.assignments:
        scopes += EDU_SCOPES
    auth = Auth(cfg)
    result = auth.login(scopes, device_code=args.device_code)
    if "access_token" not in result:
        print("Login failed:", result.get("error_description") or result, file=sys.stderr)
        print(
            "\nIf this says 'need admin approval' / AADSTS65001 / AADSTS90094, your tenant does not let\n"
            "students consent to this app. Only your school's IT can change that.",
            file=sys.stderr,
        )
        return 1
    save_config(cfg)
    who = (result.get("id_token_claims") or {}).get("preferred_username", "?")
    print(f"Signed in as {who}")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(prog=APP_NAME, description="Read-only Microsoft Teams client for the terminal")
    sub = ap.add_subparsers(dest="cmd")
    lp = sub.add_parser("login", help="sign in to Microsoft 365")
    lp.add_argument("--client-id", help="Application (client) ID of your Entra app registration")
    lp.add_argument("--tenant", help="tenant id/domain (default: organizations)")
    lp.add_argument("--device-code", action="store_true", help="use device-code flow instead of opening a browser")
    lp.add_argument("--channels", action="store_true", help="also request ChannelMessage.Read.All (needs admin consent)")
    lp.add_argument("--assignments", action="store_true", help="also request EduAssignments.ReadBasic (needs admin consent)")
    sub.add_parser("logout", help="delete cached tokens")
    args = ap.parse_args()

    if args.cmd == "login":
        return cmd_login(args)
    if args.cmd == "logout":
        TOKEN_FILE.unlink(missing_ok=True)
        print("Token cache removed.")
        return 0

    TeamsTui(Api(Auth(load_config()))).run()
    return 0


if __name__ == "__main__":
    sys.exit(main())
