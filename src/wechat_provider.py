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

        # `seen` holds the (text -> count) snapshot from the *previous* poll —
        # NOT a monotonic high-water mark. WeChat's chat_message_list is a
        # virtualised Qt ListView that only exposes the most-recent ~10 rows;
        # if we accumulated counts across all of history, repeating an old
        # text (e.g. retrying the same prompt) would be mis-classified as
        # already-seen because UIA can no longer show all the past instances.
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

        # Replace the snapshot wholesale so counts can decrease as old rows
        # scroll out of the UIA window.
        self._chat_state[name]["seen_counts"] = observed

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
        """Send a file by clicking the toolbar's '发送文件' button to bring
        up the standard Win32 file-open dialog, then driving that dialog
        via UIA: write the absolute path into its filename field, click
        '打开', and let WeChat attach the file the same way it does for a
        human user. This is the most reliable path because it is exactly
        the one a person would take.

        Trade-off: while the file dialog is open it has keyboard focus
        (Win32 file dialogs are always foreground). When the dialog
        closes, focus returns to whatever was active before."""
        if chat not in self._chat_state:
            raise RuntimeError(f"chat not in LISTEN_CHATS: {chat!r}")
        p = Path(path)
        if not p.exists() or not p.is_file():
            raise FileNotFoundError(p)
        abspath = str(p.resolve())

        file_btn = self._get_send_file_btn(chat)
        if file_btn is None:
            raise RuntimeError(f"toolbar '发送文件' button not found in {chat!r}")

        # Snapshot existing top-level Weixin windows so we can detect the new
        # one that the file dialog opens. ClassName-based detection (#32770)
        # is too narrow: WeChat 4.x sometimes pops a Qt-skinned dialog that
        # shares the Qt51514QWindowIcon class with regular sub-windows.
        existing_hwnds = self._weixin_top_level_hwnds()

        # Qt drops WM_* mouse messages on minimized windows, so a stealth
        # PostMessage click no-ops while the chat window is iconic. We
        # un-minimize via SetWindowPlacement (showCmd = SW_SHOWNOACTIVATE)
        # which un-minimizes WITHOUT activating and KEEPS the original
        # off-screen rect intact — so the window stays invisible to the
        # user but is "live" enough to process synthetic input.
        import win32con
        import win32gui as _w32g
        sub_hwnd = self._chat_state[chat]["hwnd"]

        orig_placement = _w32g.GetWindowPlacement(sub_hwnd)
        was_minimized = bool(_w32g.IsIconic(sub_hwnd))
        placement_changed = False
        if was_minimized:
            try:
                un_min_placement = (
                    orig_placement[0],
                    win32con.SW_SHOWNOACTIVATE,
                    orig_placement[2],
                    orig_placement[3],
                    orig_placement[4],
                )
                _w32g.SetWindowPlacement(sub_hwnd, un_min_placement)
                placement_changed = True
                time.sleep(0.15)
            except Exception:
                log.debug("SetWindowPlacement to un-minimize failed", exc_info=True)

        # Re-resolve the button after placement change (UIA element handles
        # may need a refresh once the window is active again).
        file_btn = self._get_send_file_btn(chat) or file_btn

        try:
            stealth_ok = self._post_click(sub_hwnd, file_btn)
            log.debug("send_file: stealth post-click ok=%s", stealth_ok)
            dlg = self._wait_new_weixin_window(
                existing_hwnds, exclude={sub_hwnd}, timeout_s=6.0,
            )

            # Fallback only if PostMessage didn't surface a dialog: try a
            # real UIA Click. Window stays off-screen; if Click insists on
            # screen coordinates the next layer will surface a clear error.
            if dlg is None:
                log.debug("send_file: stealth produced no dialog; trying UIA Click")
                try:
                    file_btn.Click(simulateMove=False, waitTime=0)
                except Exception:
                    log.debug("file_btn.Click raised; trying InvokePattern", exc_info=True)
                    try:
                        ip = file_btn.GetInvokePattern()
                        if ip is not None:
                            ip.Invoke()
                    except Exception as exc:
                        raise RuntimeError(f"could not invoke 发送文件 button: {exc}")
                dlg = self._wait_new_weixin_window(
                    existing_hwnds, exclude={sub_hwnd}, timeout_s=6.0,
                )

            if dlg is None:
                snapshot = self._weixin_top_level_hwnds()
                added = snapshot - existing_hwnds - {sub_hwnd}
                log.debug(
                    "send_file: no new window; existing=%d snapshot=%d added=%s",
                    len(existing_hwnds), len(snapshot), [hex(h) for h in added],
                )
                raise RuntimeError(
                    "file dialog did not appear after clicking 发送文件 (both stealth and visible)"
                )

            log.debug("send_file: detected new window hwnd=%s class=%r title=%r",
                      hex(getattr(dlg, "NativeWindowHandle", 0) or 0),
                      _w32g.GetClassName(dlg.NativeWindowHandle) if dlg.NativeWindowHandle else "",
                      _w32g.GetWindowText(dlg.NativeWindowHandle) if dlg.NativeWindowHandle else "")

            try:
                self._fill_file_dialog(dlg, abspath)
                self._wait_dialog_closed(dlg, timeout_s=6.0)
            except Exception:
                self._dismiss_file_dialog(dlg)
                raise

            # Composition area now has the attachment chip; brief settle delay
            # before firing send so the button's IsEnabled state stabilises.
            time.sleep(0.4)

            edit = self._get_input(chat)
            btn = self._get_send_btn(chat)
            if edit is None or btn is None:
                raise RuntimeError("send button or input vanished after dialog closed")

            def btn_enabled() -> Optional[bool]:
                try:
                    return bool(btn.IsEnabled)
                except Exception:
                    return None

            strategy = self._fire_send_attached(sub_hwnd, edit, btn, btn_enabled)
            if not strategy:
                # Last resort: blind invoke. Dialog closed cleanly so the
                # composition area has the attachment regardless of whether
                # IsEnabled is a reliable signal on this build.
                try:
                    ip = btn.GetInvokePattern()
                    if ip is not None:
                        ip.Invoke()
                        strategy = "invoke_blind"
                except Exception:
                    pass
            if not strategy:
                raise RuntimeError(
                    f"send_file({chat!r}, {path!r}): attached via toolbar but send failed"
                )
            log.debug("send_file(%s, %r) ok via toolbar+%s", chat, abspath, strategy)
        finally:
            # Restore the original placement — back to minimized if it was.
            if placement_changed:
                try:
                    _w32g.SetWindowPlacement(sub_hwnd, orig_placement)
                except Exception:
                    pass

    @staticmethod
    def _post_click(sub_hwnd: int, btn) -> bool:
        """Send a synthetic mouse click to a UIA Button via PostMessage,
        using the sub-window's client-relative coordinates derived from
        the button's BoundingRectangle. Doesn't touch cursor / focus /
        z-order, so it's invisible to the user. Returns True on success."""
        import ctypes
        import win32api
        import win32con
        import win32gui

        try:
            rect = btn.BoundingRectangle
        except Exception:
            return False
        if rect is None:
            return False
        # uiautomation Rect has .left/.top/.right/.bottom attrs.
        try:
            cx = (rect.left + rect.right) // 2
            cy = (rect.top + rect.bottom) // 2
        except Exception:
            return False
        try:
            client_x, client_y = win32gui.ScreenToClient(sub_hwnd, (cx, cy))
        except Exception:
            return False
        lparam = ((client_y & 0xFFFF) << 16) | (client_x & 0xFFFF)
        MK_LBUTTON = 0x0001
        try:
            win32api.PostMessage(sub_hwnd, win32con.WM_LBUTTONDOWN, MK_LBUTTON, lparam)
            win32api.PostMessage(sub_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
            return True
        except Exception:
            return False

    def _fire_send_attached(self, hwnd: int, edit, btn,
                            btn_enabled: Callable[[], Optional[bool]]) -> Optional[str]:
        """Send when something is attached. Success = button flips back
        from enabled to disabled (composition area emptied)."""
        import win32con
        import win32gui

        def cleared() -> bool:
            return btn_enabled() is False

        try:
            ip = btn.GetInvokePattern()
            if ip is not None:
                ip.Invoke()
                time.sleep(0.4)
                if cleared():
                    return "invoke"
        except Exception:
            log.debug("invoke strategy raised", exc_info=True)

        try:
            try:
                edit.SetFocus()
            except Exception:
                pass
            win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
            win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
            time.sleep(0.4)
            if cleared():
                return "post_enter"
        except Exception:
            log.debug("post_enter strategy raised", exc_info=True)

        try:
            was_iconic = bool(win32gui.IsIconic(hwnd))
            if was_iconic:
                win32gui.ShowWindow(hwnd, win32con.SW_SHOWNOACTIVATE)
                time.sleep(0.12)
            btn.Click(simulateMove=False, waitTime=0)
            time.sleep(0.4)
            sent = cleared()
            if was_iconic:
                win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINNOACTIVE)
            if sent:
                return "restore_click" if was_iconic else "click"
        except Exception:
            log.debug("restore_click strategy raised", exc_info=True)

        return None

    def _get_send_file_btn(self, chat: str):
        """Locate the '发送文件' button in the chat sub-window's toolbar."""
        root = self._get_root(chat)
        toolbar = self._find_by_aid(root, "tool_bar_accessible")
        if toolbar is None:
            return None
        return self._find_by(
            toolbar,
            lambda c: c.ControlTypeName == "ButtonControl" and c.Name == "发送文件",
        )

    def _weixin_top_level_hwnds(self) -> set:
        """Set of every top-level window hwnd owned by a Weixin process."""
        import win32gui
        import win32process

        weixin_pids = self._find_weixin_pids()
        out: set = set()

        def cb(h, _):
            try:
                _, pid = win32process.GetWindowThreadProcessId(h)
                if pid in weixin_pids and win32gui.IsWindowVisible(h):
                    out.add(h)
            except Exception:
                pass

        win32gui.EnumWindows(cb, None)
        return out

    def _wait_new_weixin_window(self, existing: set, exclude: set,
                                 timeout_s: float = 12.0):
        """Poll for a newly-visible Weixin window that contains an Edit
        control — that's our signal it's the file-open dialog. As soon as
        ANY new top-level Weixin window appears we move it off-screen via
        SetWindowPos (no activate, no z-order change), so the user never
        sees the dialog flash on screen. The dialog continues to function
        normally at coordinates (-30000, -30000) — UIA + PostMessage
        operations don't depend on window position."""
        import ctypes
        import uiautomation as auto
        import win32gui

        SWP_NOACTIVATE = 0x0010
        SWP_NOZORDER = 0x0004
        SWP_NOSIZE = 0x0001
        SWP_FLAGS = SWP_NOACTIVATE | SWP_NOZORDER | SWP_NOSIZE
        SetWindowPos = ctypes.windll.user32.SetWindowPos

        examined: set = set()
        deadline = time.time() + timeout_s
        while time.time() < deadline:
            now = self._weixin_top_level_hwnds()
            candidates = sorted((now - existing) - exclude, reverse=True)
            for hwnd in candidates:
                if hwnd in examined:
                    continue
                examined.add(hwnd)
                # Hide ASAP — including SoPY_Status splash and any other
                # transient window so the user perceives no flicker.
                try:
                    SetWindowPos(hwnd, 0, -30000, -30000, 0, 0, SWP_FLAGS)
                except Exception:
                    pass
                try:
                    cls = win32gui.GetClassName(hwnd)
                    title = win32gui.GetWindowText(hwnd)
                except Exception:
                    cls, title = "", ""
                if cls in ("SoPY_Status", "IME", "tooltips_class32",
                            "Internet Explorer_Hidden", "OleMainThreadWndClass"):
                    log.debug("send_file: hid non-dialog window cls=%r", cls)
                    continue
                ctrl = auto.ControlFromHandle(hwnd)
                if ctrl is None:
                    continue
                if self._has_edit_control(ctrl):
                    log.debug("send_file: matched dialog hwnd=%s cls=%r title=%r",
                              hex(hwnd), cls, title[:40])
                    return ctrl
                else:
                    log.debug("send_file: window has no Edit, hidden cls=%r title=%r",
                              cls, title[:40])
            time.sleep(0.03)
        return None

    @staticmethod
    def _has_edit_control(node, depth: int = 0, max_depth: int = 10) -> bool:
        if node is None or depth > max_depth:
            return False
        try:
            if node.ControlTypeName == "EditControl":
                return True
            for k in node.GetChildren():
                if WeChatProvider._has_edit_control(k, depth + 1, max_depth):
                    return True
        except Exception:
            pass
        return False

    def _fill_file_dialog(self, dlg, abspath: str) -> None:
        """Write *abspath* into the dialog's filename field and click 打开."""
        import uiautomation as auto

        # Collect every Edit + Button in the dialog for diagnosis and to
        # pick the right one(s) by name rather than position.
        edits: List[Any] = []
        buttons: List[Any] = []

        def visit(node, depth: int = 0):
            if depth > 10 or node is None:
                return
            try:
                ctl = node.ControlTypeName
                if ctl == "EditControl":
                    edits.append(node)
                elif ctl == "ButtonControl":
                    buttons.append(node)
                for k in node.GetChildren():
                    visit(k, depth + 1)
            except Exception:
                pass

        visit(dlg)
        log.debug("send_file: dialog has %d edits, %d buttons", len(edits), len(buttons))
        for i, e in enumerate(edits):
            log.debug("  edit[%d]: name=%r aid=%r",
                      i, (e.Name or "")[:50], (e.AutomationId or "")[:30])
        for i, b in enumerate(buttons):
            log.debug("  button[%d]: name=%r aid=%r",
                      i, (b.Name or "")[:50], (b.AutomationId or "")[:30])

        if not edits:
            raise RuntimeError("file dialog has no Edit controls")

        # Prefer the edit whose name mentions "文件名" / "File name"; else last.
        fname_edit = None
        for e in edits:
            nm = e.Name or ""
            if "文件名" in nm or "File name" in nm or "file name" in nm.lower():
                fname_edit = e
                break
        if fname_edit is None:
            fname_edit = edits[-1]

        # Try ValuePattern first; verify; fall back to SendKeys.
        try:
            fname_edit.GetValuePattern().SetValue(abspath)
        except Exception as exc:
            log.debug("send_file: SetValue raised: %s", exc)
        time.sleep(0.2)
        try:
            actual = fname_edit.GetValuePattern().Value or ""
        except Exception:
            actual = ""
        log.debug("send_file: filename edit value after SetValue: %r", actual[:120])

        if actual != abspath:
            try:
                fname_edit.SetFocus()
                time.sleep(0.1)
                # Select-all + delete first so we don't append.
                auto.SendKeys("{Ctrl}a{Delete}", waitTime=0.05)
                # Send the path. {} bracket parsing in uiautomation.SendKeys
                # treats backslashes literally; colon and slash are fine.
                auto.SendKeys(abspath, waitTime=0.02)
                time.sleep(0.25)
                actual = ""
                try:
                    actual = fname_edit.GetValuePattern().Value or ""
                except Exception:
                    pass
                log.debug("send_file: filename edit after SendKeys: %r", actual[:120])
            except Exception as exc:
                log.debug("send_file: SendKeys fallback raised: %s", exc)

        # The Win32 file dialog's Open button is a Split-button. UIA exposes
        # only the dropdown halves (aid='DropDown'), not a clean IDOK
        # ButtonControl, so clicking what UIA calls 'Open' just opens the
        # dropdown menu — the dialog never submits. The reliable approach
        # is the Win32 protocol: PostMessage(WM_COMMAND, IDOK) to the dialog
        # hwnd; the OK handler reads the filename edit and dismisses.
        import ctypes
        WM_COMMAND = 0x0111
        IDOK = 1
        dlg_hwnd = getattr(dlg, "NativeWindowHandle", 0) or 0
        if not dlg_hwnd:
            raise RuntimeError("dialog has no NativeWindowHandle for IDOK")

        log.debug("send_file: posting WM_COMMAND IDOK to dialog hwnd=%s", hex(dlg_hwnd))
        ok = ctypes.windll.user32.PostMessageW(dlg_hwnd, WM_COMMAND, IDOK, 0)
        if not ok:
            log.debug("send_file: PostMessage WM_COMMAND failed; trying Click fallback")
            # Fallback: click any button named 打开/Open (may open dropdown).
            for b in buttons:
                nm = (b.Name or "").strip()
                if nm in ("打开(O)", "打开", "Open", "Open(O)") or "打开" in nm:
                    try:
                        b.Click(simulateMove=False, waitTime=0)
                        return
                    except Exception:
                        pass
            raise RuntimeError("could not submit file dialog (IDOK)")

    @staticmethod
    def _wait_dialog_closed(dlg, timeout_s: float = 6.0) -> None:
        import win32gui
        hwnd = getattr(dlg, "NativeWindowHandle", 0) or 0
        if not hwnd:
            return
        deadline = time.time() + timeout_s
        while time.time() < deadline:
            if not win32gui.IsWindow(hwnd):
                return
            time.sleep(0.15)
        raise RuntimeError("file dialog did not close after Open")

    def _dismiss_file_dialog(self, dlg) -> None:
        """Best-effort: click Cancel/取消 to close a stuck dialog so the
        chat doesn't end up with a dangling modal blocking the input."""
        cancel = self._find_by(
            dlg,
            lambda c: c.ControlTypeName == "ButtonControl" and (
                "取消" in (c.Name or "") or "Cancel" in (c.Name or "")
            ),
        )
        if cancel is None:
            return
        try:
            ip = cancel.GetInvokePattern()
            if ip is not None:
                ip.Invoke()
        except Exception:
            pass
