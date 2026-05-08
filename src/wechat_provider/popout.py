"""Programmatically pop out a chat into its own Qt sub-window.

User flow we automate: right-click a session-list item in the main 微信
window → click "独立窗口显示" in the context menu. The result is a
top-level Qt window (class Qt51514QWindowIcon) whose title is the chat
display name — that's what the rest of the provider treats as a chat.

Stealth-first: PostMessage the right-click + menu click directly to the
window. Only fall back to a real foreground click + UIA RightClick if the
stealth path fails.
"""
from __future__ import annotations

import logging
import time
from typing import Any, List, Set

from . import win32_utils
from .uia_utils import find_by, find_by_aid
from .window_finder import find_main_weixin_window, weixin_top_level_hwnds

log = logging.getLogger(__name__)


def popout_chat(chat_name: str, weixin_pids: Set[int]) -> None:
    """Use main-window UIA to right-click the chat in the session list and
    select '独立窗口显示' from the context menu, so we don't require the
    user to set this up by hand."""
    import uiautomation as auto
    import win32api
    import win32con
    import win32gui

    main_hwnd = find_main_weixin_window(weixin_pids)
    if main_hwnd is None:
        raise RuntimeError(
            "微信主窗口 not found or minimized — un-minimize it once so "
            "sidecar can pop out its listened chats"
        )

    root = auto.ControlFromHandle(main_hwnd)
    item_aid = f"session_item_{chat_name}"
    item = find_by_aid(root, item_aid, depth=20)
    if item is None:
        # The session list is virtualized; chat may be off-screen. Try the
        # search box as a fallback to bring the chat into view.
        _search_chat_in_main(root, chat_name)
        time.sleep(0.4)
        item = find_by_aid(root, item_aid, depth=20)
        if item is None:
            raise RuntimeError(
                f"chat {chat_name!r} not found in session_list; scroll "
                "it into view in 微信主窗口 and retry"
            )

    log.info("popout(%s): right-clicking session item", chat_name)
    before = weixin_top_level_hwnds()
    menu_root = None
    saved_fore = win32gui.GetForegroundWindow()
    used_foreground = False

    # Stealth first: PostMessage WM_RBUTTONDOWN/UP + WM_CONTEXTMENU to the
    # main window's client coords. Doesn't move cursor, doesn't change
    # z-order, doesn't steal focus. Qt's session list responds to this even
    # when not foreground.
    for attempt in range(2):
        ok = win32_utils.post_right_click(main_hwnd, item)
        log.debug("popout: stealth right-click attempt %d ok=%s", attempt + 1, ok)
        if not ok:
            break
        menu_root = _wait_context_menu(before, timeout_s=1.5)
        if menu_root is not None:
            break
        time.sleep(0.2)

    # Fallback: only if stealth failed do we briefly take foreground.
    if menu_root is None:
        log.debug("popout(%s): stealth failed, falling back to foreground+RightClick", chat_name)
        try:
            win32_utils.bring_to_foreground(main_hwnd)
            used_foreground = True
            time.sleep(0.2)
        except Exception:
            log.debug("bring_to_foreground failed", exc_info=True)
        for attempt in range(3):
            try:
                item.RightClick(simulateMove=False, waitTime=0)
            except Exception as exc:
                log.debug("RightClick attempt %d raised: %s", attempt + 1, exc)
                time.sleep(0.3)
                continue
            menu_root = _wait_context_menu(before, timeout_s=2.5)
            if menu_root is not None:
                break
            time.sleep(0.3)

    if menu_root is None:
        try:
            if used_foreground and saved_fore and saved_fore != main_hwnd:
                win32_utils.bring_to_foreground(saved_fore)
        except Exception:
            pass
        raise RuntimeError("context menu did not appear after right-click")
    log.info("popout(%s): context menu detected", chat_name)

    # Dump menu items for diagnosis (helps spot wording differences).
    menu_items: List[Any] = []
    def collect(n, d=0):
        if d > 6 or n is None:
            return
        try:
            if n.ControlTypeName == "MenuItemControl":
                menu_items.append(n)
            for k in n.GetChildren():
                collect(k, d + 1)
        except Exception:
            pass
    collect(menu_root)
    for mi in menu_items:
        log.debug("popout: menu item %r", (mi.Name or "")[:40])

    popout = None
    for mi in menu_items:
        nm = (mi.Name or "").strip()
        if nm in ("独立窗口显示", "在独立窗口中打开", "Open in separate window"):
            popout = mi
            break
    if popout is None:
        try:
            mh = menu_root.NativeWindowHandle
            win32api.PostMessage(mh, win32con.WM_KEYDOWN, win32con.VK_ESCAPE, 0)
            win32api.PostMessage(mh, win32con.WM_KEYUP, win32con.VK_ESCAPE, 0)
        except Exception:
            pass
        names = [(mi.Name or "")[:30] for mi in menu_items]
        raise RuntimeError(f"'独立窗口显示' not in menu items: {names}")

    log.info("popout(%s): clicking %r", chat_name, (popout.Name or ""))
    menu_hwnd = getattr(menu_root, "NativeWindowHandle", 0) or 0
    clicked = False
    # Stealth first: PostMessage WM_LBUTTONDOWN/UP to menu window client
    # coords — invisible to the user.
    if menu_hwnd:
        clicked = win32_utils.post_click(menu_hwnd, popout)
        log.debug("popout: stealth menu click ok=%s", clicked)
        time.sleep(0.2)
    if not clicked:
        # Fallback: real Click then Invoke.
        try:
            popout.Click(simulateMove=False, waitTime=0)
            clicked = True
        except Exception:
            log.debug("popout Click raised; trying Invoke", exc_info=True)
        if not clicked:
            try:
                ip = popout.GetInvokePattern()
                if ip is not None:
                    ip.Invoke()
                    clicked = True
            except Exception as exc:
                raise RuntimeError(f"could not invoke 独立窗口显示: {exc}")
    if not clicked:
        raise RuntimeError("could not click 独立窗口显示")

    # Restore original foreground only if we touched it.
    try:
        if used_foreground and saved_fore and saved_fore != main_hwnd:
            win32_utils.bring_to_foreground(saved_fore)
    except Exception:
        pass


def _search_chat_in_main(main_root, chat_name: str) -> None:
    """Type the chat name into 微信主窗口's search box so the chat shows up
    in the session list. Best-effort; we don't fail if the search box can't
    be located — caller will check session list afterwards."""
    search_edit = find_by(
        main_root,
        lambda c: c.ControlTypeName == "EditControl"
        and (c.Name or "") == "搜索",
        depth=20,
    )
    if search_edit is None:
        return
    try:
        search_edit.GetValuePattern().SetValue(chat_name)
    except Exception:
        pass


def _wait_context_menu(existing: Set[int], timeout_s: float = 3.0):
    """Wait for the Qt context-menu top-level window (class
    Qt51514QWindowToolSaveBits) to appear; returns its UIA Control."""
    import uiautomation as auto
    import win32gui

    deadline = time.time() + timeout_s
    while time.time() < deadline:
        now = weixin_top_level_hwnds()
        for hwnd in now - existing:
            try:
                cls = win32gui.GetClassName(hwnd)
            except Exception:
                cls = ""
            if cls == "Qt51514QWindowToolSaveBits":
                return auto.ControlFromHandle(hwnd)
        time.sleep(0.08)
    return None
