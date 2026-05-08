"""Win32 input/window helpers used to drive WeChat without stealing focus.

The PostMessage-based click helpers translate a UIA control's screen
BoundingRectangle into client coords on a target hwnd, then post the
appropriate WM_*BUTTON* messages directly to that window. Qt/Weixin
respond to these even when the window is non-foreground or off-screen,
which is what makes the sidecar's "stealth" mode possible.
"""
from __future__ import annotations

import hashlib


def post_click(target_hwnd: int, ctrl) -> bool:
    """Synthetic left-click via PostMessage to *target_hwnd*'s client coords,
    centred on *ctrl*'s BoundingRectangle. Doesn't touch cursor / focus /
    z-order, so it's invisible to the user. Returns True on success."""
    import win32api
    import win32con
    import win32gui

    try:
        rect = ctrl.BoundingRectangle
    except Exception:
        return False
    if rect is None:
        return False
    try:
        cx = (rect.left + rect.right) // 2
        cy = (rect.top + rect.bottom) // 2
    except Exception:
        return False
    try:
        client_x, client_y = win32gui.ScreenToClient(target_hwnd, (cx, cy))
    except Exception:
        return False
    lparam = ((client_y & 0xFFFF) << 16) | (client_x & 0xFFFF)
    MK_LBUTTON = 0x0001
    try:
        win32api.PostMessage(target_hwnd, win32con.WM_LBUTTONDOWN, MK_LBUTTON, lparam)
        win32api.PostMessage(target_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
        return True
    except Exception:
        return False


def post_right_click(target_hwnd: int, ctrl) -> bool:
    """Synthetic right-click via PostMessage to the target's client coords.
    Sends WM_RBUTTONDOWN, WM_RBUTTONUP, then WM_CONTEXTMENU (Qt typically
    posts the context menu in response to the latter). Returns True if all
    messages posted; doesn't verify the menu actually opens — caller should
    poll for it."""
    import win32api
    import win32con
    import win32gui

    try:
        rect = ctrl.BoundingRectangle
    except Exception:
        return False
    if rect is None:
        return False
    try:
        cx = (rect.left + rect.right) // 2
        cy = (rect.top + rect.bottom) // 2
    except Exception:
        return False
    try:
        client_x, client_y = win32gui.ScreenToClient(target_hwnd, (cx, cy))
    except Exception:
        return False
    lparam_client = ((client_y & 0xFFFF) << 16) | (client_x & 0xFFFF)
    lparam_screen = ((cy & 0xFFFF) << 16) | (cx & 0xFFFF)
    MK_RBUTTON = 0x0002
    WM_CONTEXTMENU = 0x007B
    try:
        win32api.PostMessage(target_hwnd, win32con.WM_RBUTTONDOWN, MK_RBUTTON, lparam_client)
        win32api.PostMessage(target_hwnd, win32con.WM_RBUTTONUP, 0, lparam_client)
        win32api.PostMessage(target_hwnd, WM_CONTEXTMENU, target_hwnd, lparam_screen)
        return True
    except Exception:
        return False


def bring_to_foreground(hwnd: int) -> None:
    """Best-effort: bring hwnd to foreground using AttachThreadInput trick to
    bypass the foreground-lock timeout."""
    import ctypes
    import win32process

    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    fore_hwnd = user32.GetForegroundWindow()
    if fore_hwnd == hwnd:
        return
    cur = kernel32.GetCurrentThreadId()
    fore_thr = (
        win32process.GetWindowThreadProcessId(fore_hwnd)[0] if fore_hwnd else 0
    )
    target_thr = win32process.GetWindowThreadProcessId(hwnd)[0]
    attached_fore = False
    attached_target = False
    if fore_thr and fore_thr != cur:
        user32.AttachThreadInput(cur, fore_thr, True)
        attached_fore = True
    if target_thr and target_thr != cur:
        user32.AttachThreadInput(cur, target_thr, True)
        attached_target = True
    try:
        user32.BringWindowToTop(hwnd)
        user32.SetForegroundWindow(hwnd)
    finally:
        if attached_target:
            user32.AttachThreadInput(cur, target_thr, False)
        if attached_fore:
            user32.AttachThreadInput(cur, fore_thr, False)


def stable_hash(name: str, text: str, ts: float) -> str:
    """Short SHA-1 hex digest used as the message stable-id seed."""
    return hashlib.sha1(f"{name}|{text}|{int(ts)}".encode("utf-8")).hexdigest()[:16]
