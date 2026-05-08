"""Top-level orchestrator.

Owns the mutable state (chat_state, recent_outbound, send_queue), runs
the polling thread, and dispatches sends. All actual UIA/Win32 work lives
in the sibling modules — this file is glue.

Why not wxauto4: in version 41.1.2 several methods on `WeChatMainWnd`
are broken (e.g. `GetSubWindow` calls a non-existent `get_sub_wnd`), and
reading the *main* window's chat list forces the window to the
foreground every poll cycle, stealing user focus.

What we do instead:
  * Auto-pop each LISTEN_CHATS entry into its own Qt sub-window if not
    already popped (see popout.py).
  * Use the standalone `uiautomation` library (read-only) to walk each
    sub-window tree, locating the message ListView and the input/send
    controls.
  * Poll the message ListView for new items — purely passive, no focus
    theft, works even when the sub-window is minimized to taskbar.
  * Send messages via `ValuePattern.SetValue` on the input edit and a
    silent send strategy (Invoke / PostMessage Enter / no-activate Click).

Caveats:
  * If the user closes a popped-out sub-window, the next poll logs a
    warning and that chat goes silent until restart.
  * `Control.GetChildren()` caches results when called repeatedly on the
    *same* Control instance — we re-walk from the root every poll cycle
    to avoid stale snapshots.
  * Messages we just sent show up in the chat as `attr=self`. We track
    recently-sent texts (with a short TTL) so we don't echo them back
    to Linux as fresh inbound messages.
"""
from __future__ import annotations

import logging
import queue
import threading
import time
from collections import Counter, deque
from pathlib import Path
from typing import Any, Callable, Deque, Dict, List, Optional, Tuple

from ._types import IncomingMessage
from . import file_sender, poller, popout, text_sender, window_finder

log = logging.getLogger(__name__)


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

        # Per-chat state: {chat_name: {"hwnd": int, "chat_type": str, "seen_counts": Counter}}
        self._chat_state: Dict[str, Dict[str, Any]] = {}
        # Track texts we sent via SendMsg so we don't echo them back to Linux.
        # deque of (chat_name, text, ts); kept ~30s.
        self._recent_outbound: Deque[Tuple[str, str, float]] = deque(maxlen=200)

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
                    poller.poll_chat(
                        self._chat_state, self._recent_outbound,
                        self._on_message, self._bot_at_name, name,
                    )
                except Exception:
                    log.exception("poll(%s) failed", name)
            self._drain_send_queue()
            self._stop.wait(self._poll_interval)
        log.info("UIA poll loop stopped")

    # ---------- init / discovery ----------

    def _init(self) -> None:
        import uiautomation as auto  # noqa: F401  (fail-fast on missing dep)
        import win32con
        import win32gui

        weixin_pids = window_finder.find_weixin_pids()
        if not weixin_pids:
            raise RuntimeError("no running Weixin.exe — start WeChat first")
        log.info("Weixin pids=%s", sorted(weixin_pids))

        # If 微信 main window is minimized, un-minimize it off-screen so the
        # popout right-clicks can target it via UIA. We restore the original
        # placement after everything is set up. The whole un-minimize +
        # popout dance is invisible to the user (window stays at -30000)
        # and never steals focus.
        main_hwnd = window_finder.find_main_weixin_window_any_state(weixin_pids)
        main_restore_placement = None
        if main_hwnd is not None and win32gui.IsIconic(main_hwnd):
            try:
                orig_pl = win32gui.GetWindowPlacement(main_hwnd)
                onr = orig_pl[4]  # rcNormalPosition
                w = max(800, onr[2] - onr[0])
                h = max(600, onr[3] - onr[1])
                offscreen_pl = (
                    orig_pl[0],
                    win32con.SW_SHOWNOACTIVATE,
                    orig_pl[2],
                    orig_pl[3],
                    (-30000, -30000, -30000 + w, -30000 + h),
                )
                win32gui.SetWindowPlacement(main_hwnd, offscreen_pl)
                main_restore_placement = orig_pl
                log.info("main 微信 was minimized; temporarily un-minimized off-screen for popout")
                time.sleep(0.4)
            except Exception:
                log.debug("could not un-minimize main window off-screen", exc_info=True)

        try:
            self._init_chats(weixin_pids)
        finally:
            if main_restore_placement is not None:
                try:
                    win32gui.SetWindowPlacement(main_hwnd, main_restore_placement)
                    log.info("main 微信 restored to original placement (minimized)")
                except Exception:
                    log.debug("could not restore main window placement", exc_info=True)

    def _init_chats(self, weixin_pids: set) -> None:
        import uiautomation as auto

        for chat in self._listen_chats:
            hwnd = window_finder.find_chat_subwindow(chat, weixin_pids)
            if not hwnd:
                log.info("no independent sub-window for %r; popping out via main window UIA", chat)
                try:
                    popout.popout_chat(chat, weixin_pids)
                except Exception as exc:
                    raise RuntimeError(
                        f"could not auto-pop-out {chat!r}: {exc}. Make sure 微信主窗口 is "
                        f"visible (not minimized) and the chat is in the session list, "
                        f"or right-click → '独立窗口显示' manually and restart."
                    )
                # Wait up to 5s for the new sub-window to appear.
                for _ in range(25):
                    hwnd = window_finder.find_chat_subwindow(chat, weixin_pids)
                    if hwnd:
                        break
                    time.sleep(0.2)
                if not hwnd:
                    raise RuntimeError(
                        f"popped out {chat!r} but new sub-window did not appear within 5s"
                    )
                # Immediately minimize the freshly popped window so it
                # doesn't cover the main window's session_list for the
                # next chat's RightClick.
                try:
                    import win32con as _wc
                    import win32gui as _wg
                    _wg.ShowWindow(hwnd, _wc.SW_SHOWMINNOACTIVE)
                    time.sleep(0.4)
                except Exception:
                    log.debug("could not minimize freshly popped sub-window", exc_info=True)
            root = auto.ControlFromHandle(hwnd)
            aid = (root.AutomationId or "") if root else ""
            chat_type = "group" if "@chatroom" in aid else "private"
            log.info(
                "found sub-window for %s: hwnd=%s type=%s aid=%s",
                chat, hex(hwnd), chat_type, aid,
            )
            self._chat_state[chat] = {
                "hwnd": hwnd,
                "chat_type": chat_type,
                "seen_counts": Counter(),
            }
            poller.snapshot_existing(self._chat_state, chat)

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
                    text_sender.send_text(self._chat_state, self._recent_outbound, chat, text)
                elif kind == "file":
                    chat, path = args
                    file_sender.send_file(self._chat_state, chat, path)
                else:
                    raise ValueError(f"unknown send kind {kind!r}")
                done["ok"] = True
            except BaseException as exc:
                done["ok"] = False
                done["error"] = exc
                log.exception("send %s failed", kind)
            finally:
                done["event"].set()
