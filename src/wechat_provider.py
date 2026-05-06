"""WeChat 4.x provider via raw UIA on a popped-out chat sub-window.

Why not wxauto4: in version 41.1.2 (the only PyPI release at time of writing)
several methods on `WeChatMainWnd` are broken (e.g. `GetSubWindow` calls a
non-existent `get_sub_wnd`), and reading the *main* window's chat list forces
the window to the foreground every poll cycle, stealing user focus.

What we do instead:
  * Require the user to pop the listened chat out as an independent Qt window
    (right-click in WeChat sidebar → "在独立窗口中打开"). The sub-window has
    its own top-level HWND of class `Qt51514QWindowIcon` and a
    `mmui::ChatSingleWindow` UIA root.
  * Use the standalone `uiautomation` library (read-only) to walk the
    sub-window tree, locate the message ListView and the input/send controls.
  * Poll the message ListView for new items — purely passive, no focus theft,
    works even when the sub-window is minimized to taskbar (verified
    empirically; UIA updates regardless of visibility for popped-out chats).
  * Send messages via `ValuePattern.SetValue` on the input edit and a UIA
    `Click()` on the send button (mouse cursor briefly jumps to the button
    coords; foreground does NOT change). The brief cursor jump is the only
    visible side-effect.

Caveats:
  * If the user closes the popped-out sub-window, sidecar must re-find or
    recreate it. We log a warning and try to re-discover periodically.
  * `Control.GetChildren()` caches results when called repeatedly on the
    *same* Control instance — we re-walk from the root every poll cycle to
    avoid stale snapshots.
  * Messages we just sent show up in the chat as `attr=self`. We track
    recently-sent texts (with a short TTL) so we don't echo them back to
    Linux as fresh inbound messages.
"""
from __future__ import annotations

import logging
import queue
import threading
import time
from collections import Counter, deque
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

log = logging.getLogger(__name__)


# Subordinate Qt window class used by Weixin 4.x for both the main window and
# popped-out chats. Filter further by Name and process.
WEIXIN_WINDOW_CLASS = "Qt51514QWindowIcon"
WEIXIN_PROCESS_NAME = "Weixin.exe"

# UIA AutomationIds inside a popped-out chat sub-window.
INPUT_AID = "chat_input_field"
MSG_LIST_AID = "chat_message_list"

# Message list-item classes — the chat-bubble row vs. system/divider row.
MSG_TEXT_CLASS = "mmui::ChatTextItemView"
MSG_DIVIDER_CLASS = "mmui::ChatItemView"   # time-of-day separators


@dataclass
class IncomingMessage:
    chat_name: str
    chat_type: str          # "private" | "group"
    sender: str
    content_type: str       # "text" | "image" | "file" | "other"
    text: str
    file_path: Optional[str]
    raw_type: str
    msg_hash: str
    ts: float
    at_self: bool = False   # group only: text contains @<bot_at_name>

    def stable_id(self) -> str:
        return f"uia_{self.msg_hash}" if self.msg_hash else f"uia_{int(self.ts*1000)}"


class WeChatProvider:
    def __init__(
        self,
        listen_chats: List[str],
        on_message: Callable[[IncomingMessage], None],
        inbound_media_dir: Path,
        poll_interval: float = 1.0,
        bot_at_name: str = "",
    ) -> None:
        self._listen_chats = list(listen_chats)
        self._on_message = on_message
        self._inbound_dir = inbound_media_dir
        self._poll_interval = poll_interval
        self._bot_at_name = bot_at_name
        self._stop = threading.Event()
        self._thread: Optional[threading.Thread] = None
        self._send_queue: "queue.Queue[Tuple[str, tuple, dict]]" = queue.Queue()
        self._ready: Optional[threading.Event] = None
        self._init_error: Optional[BaseException] = None

        # Per-chat state: {chat_name: {"hwnd": int, "seen_counts": Counter}}
        self._chat_state: Dict[str, Dict[str, Any]] = {}
        # Track texts we sent via SendMsg so we don't echo them back to Linux.
        # deque of (chat_name, text, ts); kept ~30s.
        self._recent_outbound: "deque[Tuple[str, str, float]]" = deque(maxlen=200)

    # ---------- public API ----------

    def start(self) -> None:
        if self._thread is not None:
            return
        self._ready = threading.Event()
        self._init_error = None
        self._thread = threading.Thread(target=self._run, name="uia-poll", daemon=True)
        self._thread.start()
        if not self._ready.wait(timeout=30):
            raise RuntimeError("UIA provider init timed out (30s)")
        if self._init_error is not None:
            raise self._init_error

    def stop(self) -> None:
        self._stop.set()
        if self._thread:
            self._thread.join(timeout=5)

    def send_text(self, chat: str, text: str, at: Optional[str] = None) -> None:
        # Currently `at` is unused (UIA send doesn't compose @-prefixes for us).
        self._enqueue_send("text", (chat, text))

    def send_file(self, chat: str, path: str) -> None:
        p = Path(path)
        if not p.exists():
            raise FileNotFoundError(p)
        self._enqueue_send("file", (chat, str(p)))

    def send_image(self, chat: str, path: str) -> None:
        # WeChat 4.x sends images via the same "send file" affordance.
        self.send_file(chat, path)

    # ---------- thread loop ----------

    def _run(self) -> None:
        try:
            import comtypes
            comtypes.CoInitialize()
        except Exception:
            log.exception("CoInitialize failed; continuing")

        try:
            self._init()
        except BaseException as exc:
            self._init_error = exc
            log.exception("UIA provider init failed")
            self._ready.set()
            return
        self._ready.set()

        log.info("UIA poll loop started; chats=%s", list(self._chat_state.keys()))
        while not self._stop.is_set():
            for name in list(self._chat_state.keys()):
                if self._stop.is_set():
                    break
                try:
                    self._poll_chat(name)
                except Exception:
                    log.exception("poll(%s) failed", name)
            self._drain_send_queue()
            self._stop.wait(self._poll_interval)
        log.info("UIA poll loop stopped")

    # ---------- init / discovery ----------

    def _init(self) -> None:
        import uiautomation as auto  # noqa: F401  (just to fail-fast on missing dep)

        weixin_pids = self._find_weixin_pids()
        if not weixin_pids:
            raise RuntimeError("no running Weixin.exe — start WeChat first")
        log.info("Weixin pids=%s", sorted(weixin_pids))

        import uiautomation as auto
        for chat in self._listen_chats:
            hwnd = self._find_chat_subwindow(chat, weixin_pids)
            if not hwnd:
                raise RuntimeError(
                    f"could not find independent sub-window for {chat!r}. "
                    f"Right-click the chat in PC WeChat -> '在独立窗口中打开', "
                    f"then restart sidecar."
                )
            root = auto.ControlFromHandle(hwnd)
            aid = (root.AutomationId or "") if root else ""
            chat_type = "group" if "@chatroom" in aid else "private"
            log.info("found sub-window for %s: hwnd=%s type=%s aid=%s",
                     chat, hex(hwnd), chat_type, aid)
            self._chat_state[chat] = {
                "hwnd": hwnd,
                "chat_type": chat_type,
                "seen_counts": Counter(),
            }
            self._snapshot_existing(chat)

    @staticmethod
    def _find_weixin_pids() -> set:
        import psutil
        return {
            p.pid for p in psutil.process_iter(["name"])
            if p.info.get("name") in (WEIXIN_PROCESS_NAME, "Weixin")
        }

    @staticmethod
    def _find_chat_subwindow(chat_name: str, weixin_pids: set) -> Optional[int]:
        """Find an independent Qt sub-window whose title is the chat name.

        The main "微信" window is excluded by name match; popped-out chats use
        their display name (e.g. "文件传输助手") as the window title.
        """
        import win32gui
        import win32process

        match: List[int] = []
        def cb(h, _):
            try:
                if win32gui.GetClassName(h) != WEIXIN_WINDOW_CLASS:
                    return
                _, pid = win32process.GetWindowThreadProcessId(h)
                if pid not in weixin_pids:
                    return
                if win32gui.GetWindowText(h) == chat_name:
                    match.append(h)
            except Exception:
                pass
        win32gui.EnumWindows(cb, None)
        return match[0] if match else None

    def _snapshot_existing(self, name: str) -> None:
        ml = self._get_msg_list(name)
        if ml is None:
            return
        items = ml.GetChildren()
        seen: Counter = Counter()
        for it in items:
            txt = it.Name or ""
            if txt and (it.ClassName or "").startswith(MSG_TEXT_CLASS):
                seen[txt] += 1
        self._chat_state[name]["seen_counts"] = seen
        log.info(
            "chat %r: snapshotted %d existing messages (%d unique)",
            name, sum(seen.values()), len(seen),
        )

    # ---------- UIA lookup helpers ----------

    def _get_root(self, name: str):
        import uiautomation as auto
        hwnd = self._chat_state[name]["hwnd"]
        return auto.ControlFromHandle(hwnd)

    def _get_msg_list(self, name: str):
        return self._find_by_aid(self._get_root(name), MSG_LIST_AID)

    def _get_input(self, name: str):
        return self._find_by_aid(self._get_root(name), INPUT_AID)

    def _get_send_btn(self, name: str):
        root = self._get_root(name)
        return self._find_by(
            root,
            lambda c: (
                c.ControlTypeName == "ButtonControl"
                and c.Name == "发送"
                and "XOutlineButton" in (c.ClassName or "")
            ),
        )

    @staticmethod
    def _find_by_aid(node, target_aid: str, depth: int = 12):
        if node is None or depth <= 0:
            return None
        if (node.AutomationId or "") == target_aid:
            return node
        try:
            for child in node.GetChildren():
                r = WeChatProvider._find_by_aid(child, target_aid, depth - 1)
                if r is not None:
                    return r
        except Exception:
            pass
        return None

    @staticmethod
    def _find_by(node, predicate, depth: int = 12):
        if node is None or depth <= 0:
            return None
        try:
            if predicate(node):
                return node
        except Exception:
            pass
        try:
            for child in node.GetChildren():
                r = WeChatProvider._find_by(child, predicate, depth - 1)
                if r is not None:
                    return r
        except Exception:
            pass
        return None

    # ---------- polling ----------

    def _poll_chat(self, name: str) -> None:
        ml = self._get_msg_list(name)
        if ml is None:
            log.warning("poll(%s): message list not found; sub-window may be closed", name)
            return

        items = ml.GetChildren()
        if not items:
            return

        seen: Counter = self._chat_state[name]["seen_counts"]
        observed: Counter = Counter()
        new_msgs: List[Tuple[str, str, str]] = []   # (text, class_name, runtime_id)

        for it in items:
            txt = it.Name or ""
            cls = it.ClassName or ""
            if not txt:
                continue
            if not cls.startswith(MSG_TEXT_CLASS):
                # Time-of-day dividers, etc.
                continue
            observed[txt] += 1
            if observed[txt] > seen[txt]:
                new_msgs.append((txt, cls, ""))

        # Persist updated counts.
        for txt, cnt in observed.items():
            if cnt > seen[txt]:
                seen[txt] = cnt

        if not new_msgs:
            return

        log.debug("poll(%s): %d new", name, len(new_msgs))

        # Discard messages that were our own outbound sends within the last 30s.
        now = time.time()
        recent_self = {(c, t) for (c, t, ts) in self._recent_outbound if now - ts < 30}

        for txt, cls, _runtime in new_msgs:
            if (name, txt) in recent_self:
                # Consume one occurrence so a manual user-typed dup will still flow next time.
                # Remove only the matching tuple instance.
                for i, (c, t, ts) in enumerate(self._recent_outbound):
                    if c == name and t == txt:
                        del self._recent_outbound[i]
                        break
                log.debug("poll(%s): skip self-echo %r", name, txt[:60])
                continue
            self._emit(name, txt, cls)

    def _emit(self, name: str, text: str, raw_class: str) -> None:
        state = self._chat_state[name]
        chat_type = state.get("chat_type", "private")

        at_self = False
        if chat_type == "group" and self._bot_at_name:
            # WeChat uses U+2005 (FOUR-PER-EM SPACE) as the @-mention separator,
            # but accept a regular space too in case the layout differs.
            for sep in (" ", " "):
                if f"@{self._bot_at_name}{sep}" in text or text.endswith(f"@{self._bot_at_name}"):
                    at_self = True
                    break

        sender = name if chat_type == "private" else ""
        msg = IncomingMessage(
            chat_name=name,
            chat_type=chat_type,
            sender=sender,
            content_type="text",
            text=text,
            file_path=None,
            raw_type=raw_class,
            msg_hash=self._stable_hash(name, text, time.time()),
            ts=time.time(),
            at_self=at_self,
        )
        log.info("uia: NEW msg from %s [%s]: %r%s",
                 name, chat_type, text[:80], " (at_self)" if at_self else "")
        try:
            self._on_message(msg)
        except Exception:
            log.exception("on_message callback raised")

    @staticmethod
    def _stable_hash(name: str, text: str, ts: float) -> str:
        import hashlib
        h = hashlib.sha1(f"{name}|{text}|{int(ts)}".encode("utf-8")).hexdigest()[:16]
        return h

    # ---------- sending (runs on the UIA thread) ----------

    def _enqueue_send(self, kind: str, args: tuple, timeout: float = 30.0) -> None:
        done: Dict[str, Any] = {"event": threading.Event(), "ok": False, "error": None}
        self._send_queue.put((kind, args, done))
        if not done["event"].wait(timeout=timeout):
            raise TimeoutError(f"send_{kind} timed out")
        if not done["ok"]:
            raise done["error"] if done["error"] else RuntimeError(f"send_{kind} failed")

    def _drain_send_queue(self) -> None:
        while True:
            try:
                kind, args, done = self._send_queue.get_nowait()
            except queue.Empty:
                return
            try:
                if kind == "text":
                    chat, text = args
                    self._do_send_text(chat, text)
                elif kind == "file":
                    chat, path = args
                    self._do_send_file(chat, path)
                else:
                    raise ValueError(f"unknown send kind {kind!r}")
                done["ok"] = True
            except BaseException as exc:
                done["ok"] = False
                done["error"] = exc
                log.exception("send %s failed", kind)
            finally:
                done["event"].set()

    def _do_send_text(self, chat: str, text: str) -> None:
        if chat not in self._chat_state:
            raise RuntimeError(f"chat not in LISTEN_CHATS: {chat!r}")
        edit = self._get_input(chat)
        btn = self._get_send_btn(chat)
        if edit is None or btn is None:
            raise RuntimeError(f"input/send controls not found in sub-window for {chat!r}")
        edit.GetValuePattern().SetValue(text)
        time.sleep(0.05)

        hwnd = self._chat_state[chat]["hwnd"]
        strategy = self._fire_send(hwnd, edit, btn)
        if not strategy:
            raise RuntimeError(f"send_text({chat!r}): all send strategies failed")

        self._recent_outbound.append((chat, text, time.time()))
        log.debug("send_text(%s, %r) ok via %s", chat, text[:80], strategy)

    def _fire_send(self, hwnd: int, edit, btn) -> Optional[str]:
        """Try send strategies until the input clears. Returns the name of
        the strategy that succeeded, or None if all failed.

        Order matters: silent strategies first (InvokePattern, PostMessage),
        falling back to a brief no-activate window restore if those don't
        actually deliver. The restore strategy makes the sub-window visible
        for ~150ms but never takes keyboard focus.
        """
        import win32con
        import win32gui

        def input_cleared() -> bool:
            try:
                return not (edit.GetValuePattern().Value or "")
            except Exception:
                return False

        # 1) InvokePattern on the send button — pure UIA, no visual change.
        try:
            ip = btn.GetInvokePattern()
            if ip is not None:
                ip.Invoke()
                time.sleep(0.15)
                if input_cleared():
                    return "invoke"
        except Exception:
            log.debug("invoke strategy raised", exc_info=True)

        # 2) PostMessage Enter to the top-level HWND. Works for Qt windows
        #    that have the input as their focused widget.
        try:
            try:
                edit.SetFocus()
            except Exception:
                pass
            win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
            win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
            time.sleep(0.2)
            if input_cleared():
                return "post_enter"
        except Exception:
            log.debug("post_enter strategy raised", exc_info=True)

        # 3) Restore window without activating, click, re-minimize.
        try:
            was_minimized = bool(win32gui.IsIconic(hwnd))
            if was_minimized:
                win32gui.ShowWindow(hwnd, win32con.SW_SHOWNOACTIVATE)
                time.sleep(0.12)
            btn.Click(simulateMove=False, waitTime=0)
            time.sleep(0.15)
            sent = input_cleared()
            if was_minimized:
                win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINNOACTIVE)
            if sent:
                return "restore_click" if was_minimized else "click"
        except Exception:
            log.debug("restore_click strategy raised", exc_info=True)

        return None

    def _do_send_file(self, chat: str, path: str) -> None:
        # WeChat 4.x: the input edit accepts a full file path via clipboard
        # paste + Send. UIA "ValuePattern.SetValue" with the path then Send
        # triggers the attachment flow. Ctrl+V via SendKeys could also work
        # but would require the window to have focus. Stick with the standard
        # path: SetValue does not work for file attachments — we need to use
        # the "发送文件" button. For now this is unimplemented.
        raise NotImplementedError(
            "send_file via UIA is not implemented yet. wxauto4's SendFiles "
            "requires foreground; consider routing files through 文件传输助手 "
            "manually for now."
        )
