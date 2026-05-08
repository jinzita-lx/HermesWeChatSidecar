"""Process and window enumeration for WeChat 4.x.

Everything here is read-only — no input is synthesised, no windows are
shown/moved. Callers combine these with win32_utils + popout to do work.
"""
from __future__ import annotations

from typing import List, Optional, Set

from ._types import WEIXIN_PROCESS_NAME, WEIXIN_WINDOW_CLASS


def find_weixin_pids() -> Set[int]:
    """All currently-running Weixin.exe PIDs."""
    import psutil
    return {
        p.pid for p in psutil.process_iter(["name"])
        if p.info.get("name") in (WEIXIN_PROCESS_NAME, "Weixin")
    }


def find_chat_subwindow(chat_name: str, weixin_pids: Set[int]) -> Optional[int]:
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


def find_main_weixin_window(weixin_pids: Set[int]) -> Optional[int]:
    """Find the main 微信 window (the one with the chat list, NOT a popped-out
    single-chat window). Must be visible and not iconic."""
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
            if win32gui.GetWindowText(h) != "微信":
                return
            if win32gui.IsIconic(h):
                return
            match.append(h)
        except Exception:
            pass
    win32gui.EnumWindows(cb, None)
    return match[0] if match else None


def find_main_weixin_window_any_state(weixin_pids: Set[int]) -> Optional[int]:
    """Same as find_main_weixin_window but also matches a minimized (iconic)
    main window. Used at startup so we can detect a minimized main window
    and un-minimize it off-screen before popping out chats."""
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
            if win32gui.GetWindowText(h) != "微信":
                return
            match.append(h)
        except Exception:
            pass
    win32gui.EnumWindows(cb, None)
    return match[0] if match else None


def weixin_top_level_hwnds() -> Set[int]:
    """Set of every visible top-level window hwnd owned by a Weixin process."""
    import win32gui
    import win32process

    weixin_pids = find_weixin_pids()
    out: Set[int] = set()

    def cb(h, _):
        try:
            _, pid = win32process.GetWindowThreadProcessId(h)
            if pid in weixin_pids and win32gui.IsWindowVisible(h):
                out.add(h)
        except Exception:
            pass

    win32gui.EnumWindows(cb, None)
    return out
